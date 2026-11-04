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

  // 自动匹配只建立一对一关系，后续手工映射会在下一阶段补充。
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

  return {
    colLetter,
    flattenFields,
    flattenTemplate,
    normalizeHeader,
    autoMatchHeaders,
    convertCell,
    classifyCell,
  };
});
