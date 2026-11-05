// 纯逻辑层：展开 BOM 模板字段、自动匹配源表头、识别单元格类型。
(function (root, factory) {
  const Core = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = Core;
  root.Core = Core;
})(typeof self !== "undefined" ? self : globalThis, function () {
  "use strict";

  function colLetter(index) {
    let value = "";
    let number = index + 1;
    while (number > 0) {
      const remainder = (number - 1) % 26;
      value = String.fromCharCode(65 + remainder) + value;
      number = Math.floor((number - 1) / 26);
    }
    return value;
  }

  // 把模板树展开成叶子字段，只有叶子字段代表一列。
  function flattenFields(fields, parentKeys, parentLabels) {
    const keys = parentKeys || [];
    const labels = parentLabels || [];
    const result = [];
    for (const field of fields || []) {
      if (!field || field.visible === false || typeof field.key !== "string" || !field.key) continue;
      const path = keys.concat(field.key);
      const labelPath = labels.concat(field.label || field.key);
      if (Array.isArray(field.children) && field.children.length > 0) {
        result.push(...flattenFields(field.children, path, labelPath));
      } else {
        result.push({
          id: path.join("."),
          key: field.key,
          label: labelPath[labelPath.length - 1],
          path,
          labelPath,
          groupLabels: labelPath.slice(0, -1),
          aliases: Array.isArray(field.aliases) ? field.aliases.slice() : [],
        });
      }
    }
    return result;
  }

  // 兼容当前模板的 groups 和旧结构的 fields。
  function flattenTemplate(template) {
    const groups = Array.isArray(template.groups)
      ? template.groups
      : [{ key: "", label: "", fields: template.fields || [] }];
    const result = [];
    for (const group of groups) {
      if (!group || group.visible === false) continue;
      result.push(...flattenFields(
        group.fields,
        group.key ? [group.key] : [],
        group.label ? [group.label] : []
      ));
    }
    return result;
  }

  // 表头比较忽略空格和常见标点，不做包含式模糊匹配。
  function normalizeHeader(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[\s\u3000]+/g, "")
      .replace(/[()（）\[\]【】/\\_\-·.,，。:%％]/g, "");
  }

  // 自动匹配只建立一对一关系，作为手动列选择的初始值。
  function autoMatchHeaders(sourceHeaders, fields) {
    const candidates = [];
    (sourceHeaders || []).forEach((header, headerIndex) => {
      const sourceField = normalizeHeader(header.label || header.fullLabel);
      (fields || []).forEach((field, fieldIndex) => {
        const targetField = normalizeHeader(field.label);
        let score = 0;
        let reason = "";

        if (sourceField && sourceField === targetField) {
          score = 100;
          reason = "字段名称完全匹配";
        } else {
          const aliasIndex = (field.aliases || [])
            .map(normalizeHeader)
            .findIndex((alias) => alias && alias === sourceField);
          if (aliasIndex >= 0) {
            score = 90 - Math.min(aliasIndex, 5);
            reason = "通过字段别名匹配";
          }
        }

        if (score > 0) candidates.push({ headerIndex, fieldIndex, score, reason });
      });
    });

    candidates.sort((left, right) => right.score - left.score || left.headerIndex - right.headerIndex);
    const headerMatches = new Map();
    const usedFields = new Set();
    for (const candidate of candidates) {
      if (headerMatches.has(candidate.headerIndex) || usedFields.has(candidate.fieldIndex)) continue;
      headerMatches.set(candidate.headerIndex, candidate);
      usedFields.add(candidate.fieldIndex);
    }

    const matches = (sourceHeaders || []).map((header, headerIndex) => {
      const candidate = headerMatches.get(headerIndex);
      const field = candidate ? fields[candidate.fieldIndex] : null;
      return {
        headerIndex,
        header,
        fieldId: field ? field.id : null,
        field,
        reason: candidate ? candidate.reason : "未找到匹配字段",
      };
    });
    return {
      matches,
      matchedCount: matches.filter((match) => match.field).length,
      unmatchedTemplateFields: (fields || []).filter(
        (field) => !matches.some((match) => match.fieldId === field.id)
      ),
    };
  }

  // 空值统一为 null，日期统一为 ISO 字符串，数字和布尔值保持原类型。
  function convertCell(value) {
    if (value === undefined || value === null) return null;
    if (typeof value === "string" && value.trim() === "") return null;
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? null : value.toISOString();
    }
    return value;
  }

  // 用于预览区域的数据类型统计。
  function classifyCell(value) {
    if (value === undefined || value === null) return "empty";
    if (typeof value === "string" && value.trim() === "") return "empty";
    if (value instanceof Date) return "date";
    if (typeof value === "number") return "number";
    if (typeof value === "boolean") return "boolean";
    return "text";
  }

  // 校验手动映射：字段不能缺列，同一 Excel 列也不能分配给多个字段。
  function resolveFieldMapping(sourceHeaders, fields, columnIndexByFieldId) {
    const headers = sourceHeaders || [];
    const mapping = columnIndexByFieldId || {};
    const usedColumns = new Map();
    const duplicateColumns = new Set();

    const fieldMappings = (fields || []).map((field) => {
      const rawIndex = mapping[field.id];
      const numericIndex = rawIndex === "" || rawIndex === undefined || rawIndex === null
        ? NaN
        : Number(rawIndex);
      const headerIndex = Number.isInteger(numericIndex) &&
        numericIndex >= 0 &&
        numericIndex < headers.length
        ? numericIndex
        : null;

      if (headerIndex !== null) {
        if (usedColumns.has(headerIndex)) duplicateColumns.add(headerIndex);
        else usedColumns.set(headerIndex, field.id);
      }

      return {
        fieldId: field.id,
        field,
        headerIndex,
        header: headerIndex === null ? null : headers[headerIndex],
      };
    });

    const missingFields = fieldMappings.filter((item) => item.headerIndex === null);
    return {
      fieldMappings,
      matchedCount: fieldMappings.length - missingFields.length,
      missingFields,
      duplicateColumns: Array.from(duplicateColumns),
      valid: missingFields.length === 0 && duplicateColumns.size === 0,
    };
  }

  // 按字段映射把源表整行转换为固定字段对象，未选择列时该字段为空。
  function buildMappedRows(fieldMappings, rows) {
    return (rows || []).map((row, rowIndex) => {
      const record = { _rowIndex: rowIndex };
      (fieldMappings || []).forEach((mapping) => {
        record[mapping.field.key] = mapping.headerIndex === null
          ? null
          : convertCell(row[mapping.headerIndex]);
      });
      return record;
    });
  }

  function isBlank(value) {
    return value === undefined || value === null || String(value).trim() === "";
  }

  function normalizedText(value) {
    return isBlank(value) ? null : String(value).trim();
  }

  // 完整检查当前 BOM；空行自动忽略，问题行使用实际 Excel 行号报告。
  function validateImportRows(rows, dataStartRow) {
    const validRows = [];
    const errors = [];
    const seenPartNos = new Map();
    const firstDataRow = Number(dataStartRow) || 2;
    const lengthRules = [
      ["partNo", "零件编号", 200],
      ["partName", "零件名称", 150],
      ["partMaterial", "材质", 300],
      ["partSpecs", "实际规格", 300],
      ["partRemark", "备注", 200],
      ["partProcessReq", "加工要求", 200],
    ];

    (rows || []).forEach((sourceRow) => {
      const row = { ...sourceRow };
      const excelRow = firstDataRow + Number(row._rowIndex || 0);
      const values = Object.keys(row).filter((key) => key !== "_rowIndex");
      if (values.every((key) => isBlank(row[key]))) return;

      row.partNo = normalizedText(row.partNo);
      row.partName = normalizedText(row.partName);
      row.partMaterial = normalizedText(row.partMaterial);
      row.partSpecs = normalizedText(row.partSpecs);
      row.partRemark = normalizedText(row.partRemark);
      row.partProcessReq = normalizedText(row.partProcessReq);
      row.partNum = Number(row.partNum);

      const rowErrors = [];
      if (!row.partNo) rowErrors.push("零件编号为空");
      if (!row.partName) rowErrors.push("零件名称为空");
      if (!Number.isInteger(row.partNum) || row.partNum < 1) {
        rowErrors.push("数量必须是大于 0 的整数");
      }

      lengthRules.forEach(([key, label, maxLength]) => {
        if (row[key] && row[key].length > maxLength) {
          rowErrors.push(`${label}不能超过 ${maxLength} 个字符`);
        }
      });

      if (row.partNo && seenPartNos.has(row.partNo)) {
        rowErrors.push(
          `零件编号 ${row.partNo} 与第 ${seenPartNos.get(row.partNo)} 行重复`
        );
      }

      if (rowErrors.length > 0) {
        errors.push(`第 ${excelRow} 行：${rowErrors.join("；")}`);
        return;
      }
      seenPartNos.set(row.partNo, excelRow);
      validRows.push(row);
    });

    return { rows: validRows, errors };
  }

  return {
    colLetter,
    flattenFields,
    flattenTemplate,
    normalizeHeader,
    autoMatchHeaders,
    resolveFieldMapping,
    convertCell,
    classifyCell,
    buildMappedRows,
    validateImportRows,
  };
});
