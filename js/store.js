// 本地持久化：保存字段顺序和用户自定义模板，不恢复上次选中的模板。
// 内置模板来自 data/templates.json；自定义模板只保存在当前浏览器本地。
(function (root) {
  "use strict";

  const KEY = "wps-field-mapping/v2";

  function load() {
    try {
      const raw = localStorage.getItem(KEY);
      if (!raw) return { orders: {}, customTemplates: [] };
      const data = JSON.parse(raw);
      return {
        orders: data.orders && typeof data.orders === "object" ? data.orders : {},
        customTemplates: Array.isArray(data.customTemplates) ? data.customTemplates : [],
      };
    } catch (_) {
      return { orders: {}, customTemplates: [] };
    }
  }

  function save(state) {
    try {
      localStorage.setItem(
        KEY,
        JSON.stringify({
          orders: state.orders || {},
          customTemplates: state.customTemplates || [],
        })
      );
      return true;
    } catch (_) {
      // 浏览器禁用存储或空间不足时由调用方决定是否提示用户。
      return false;
    }
  }

  root.Store = { load, save };
})(typeof self !== "undefined" ? self : globalThis);
