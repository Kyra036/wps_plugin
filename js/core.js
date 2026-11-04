// 纯逻辑层：模板字段解析、排序、字段标签和嵌套 JSON 导出。
// 模板来自 data/mock.json，结构可以是 groups[].fields[]，也兼容旧的 fields[]。
// visible=false 的组或字段不会进入字段列表，也不会参与上传。
(function (root, factory) {
  const Core = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = Core;
  root.Core = Core;
})(typeof self !== "undefined" ? self : globalThis, function () {
  "use strict";

  const RESULT_VERSION = "2.0";

  function colLetter(index) {
    // 0 基列索引 → 列字母（0 → A）
    let s = "";
    let n = index + 1;
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // 递归展开任意层级的字段树，返回按模板声明顺序排列的叶子字段。
  // 只有叶子字段对应工作表列；分组路径用于显示分组和生成嵌套 JSON。
  function flattenFields(fields, parentKeys, parentLabels, includeHidden) {
    const keys = parentKeys || [];
    const labels = parentLabels || [];
    const out = [];
    for (const field of fields || []) {
      if (!field || typeof field.key !== "string" || !field.key) continue;
      const visible = field.visible !== false;
      if (!visible && !includeHidden) continue;
      const path = keys.concat(field.key);
      const labelPath = labels.concat(field.label || field.key);
      if (Array.isArray(field.children) && field.children.length > 0) {
        out.push(...flattenFields(field.children, path, labelPath, includeHidden));
      } else {
        out.push({
          id: path.join("."),
          key: field.key,
          label: labelPath[labelPath.length - 1],
          path,
          labelPath,
          groupLabels: labelPath.slice(0, -1),
          visible,
        });
      }
    }
    return out;
  }

  /**
   * 展开模板分组。
   * groups[].visible=false 时整组跳过；字段级 visible=false 由 flattenFields 处理。
   * sourceIndex 保留字段在完整模板中的原始列位置，隐藏字段不会导致后续列错位。
   */
  function flattenTemplate(template) {
    const groups = Array.isArray(template.groups)
      ? template.groups
      : [{ key: "", label: "", fields: template.fields || [] }];
    const sourceLeaves = [];
    for (const group of groups) {
      if (!group || group.visible === false) continue;
      sourceLeaves.push(...flattenFields(
        group.fields,
        group.key ? [group.key] : [],
        group.label ? [group.label] : [],
        true
      ));
    }
    sourceLeaves.forEach((leaf, index) => {
      leaf.sourceIndex = index;
    });
    return sourceLeaves.filter((leaf) => leaf.visible);
  }

  function leafMap(leaves) {
    const map = new Map();
    for (const leaf of leaves) map.set(leaf.id, leaf);
    return map;
  }

  function defaultOrder(leaves) {
    return leaves.map((leaf) => leaf.id);
  }

  // 字段集合发生变化时重置顺序，避免新增字段被追加到末尾后与源列错位。
  function normalizeOrder(order, leaves) {
    const ids = defaultOrder(leaves);
    const allowed = new Set(ids);
    const kept = [];
    const seen = new Set();
    for (const id of Array.isArray(order) ? order : []) {
      if (allowed.has(id) && !seen.has(id)) {
        seen.add(id);
        kept.push(id);
      }
    }
    return kept.length === ids.length ? kept : ids;
  }

  function moveItem(list, from, to) {
    if (from === to || from < 0 || to < 0 || from >= list.length || to >= list.length) {
      return list.slice();
    }
    const next = list.slice();
    const [item] = next.splice(from, 1);
    next.splice(to, 0, item);
    return next;
  }

  function orderedLeaves(order, leaves) {
    const map = leafMap(leaves);
    return order.map((id) => map.get(id)).filter(Boolean);
  }

  // 按当前字段顺序输出模板叶子字段的显示名，用于字段预览。
  function headerLabels(order, leaves) {
    return orderedLeaves(order, leaves).map((leaf) => leaf.label);
  }

  function convertCell(value) {
    // 所有导出数据在这里统一做空值和日期转换，避免各适配器产生不同格式。
    if (value === undefined || value === null) return null;
    if (typeof value === "string" && value.trim() === "") return null;
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? null : value.toISOString();
    }
    return value;
  }

  function classifyCell(value) {
    // 只用于预览统计，不改变实际单元格值。
    if (value === undefined || value === null) return "empty";
    if (typeof value === "string" && value.trim() === "") return "empty";
    if (value instanceof Date) return "date";
    if (typeof value === "number") return "number";
    if (typeof value === "boolean") return "boolean";
    return "text";
  }

  function setPath(target, path, value) {
    let node = target;
    for (let i = 0; i < path.length - 1; i++) {
      const key = path[i];
      if (!node[key] || typeof node[key] !== "object" || Array.isArray(node[key])) {
        node[key] = {};
      }
      node = node[key];
    }
    node[path[path.length - 1]] = value;
  }

  // 行数据按当前列顺序读取，再按字段路径还原成多层对象。
  function buildRecords(rows, leaves) {
    return rows.map((row) => {
      const record = {};
      leaves.forEach((leaf, index) => {
        const sourceIndex = Number.isInteger(leaf.sourceIndex) ? leaf.sourceIndex : index;
        setPath(record, leaf.path, convertCell(row[sourceIndex]));
      });
      return record;
    });
  }

  function buildExport(template, sheetName, order, leaves, rows) {
    const current = orderedLeaves(order, leaves);
    return {
      version: RESULT_VERSION,
      template: { id: template.id, name: template.name },
      sheet: sheetName,
      fields: current.map((leaf, index) => ({
        path: leaf.path.join("."),
        label: leaf.label,
        // 使用源字段位置，隐藏字段后列字母不会重新压缩。
        column: colLetter(Number.isInteger(leaf.sourceIndex) ? leaf.sourceIndex : index),
      })),
      records: buildRecords(rows, current),
    };
  }

  return {
    RESULT_VERSION,
    colLetter,
    flattenFields,
    flattenTemplate,
    defaultOrder,
    normalizeOrder,
    moveItem,
    orderedLeaves,
    headerLabels,
    convertCell,
    classifyCell,
    buildRecords,
    buildExport,
  };
});
