// 本地持久化：只保存当前模板选择和每个模板下的字段顺序。
// 模板定义本身始终来自 data/mock.json，插件不写入、不修改。
(function (root) {
  "use strict";

  const KEY = "wps-field-mapping/v2";

  function load() {
    try {
      const raw = localStorage.getItem(KEY);
      if (!raw) return { templateId: "", orders: {} };
      const data = JSON.parse(raw);
      return {
        templateId: typeof data.templateId === "string" ? data.templateId : "",
        orders: data.orders && typeof data.orders === "object" ? data.orders : {},
      };
    } catch (_) {
      return { templateId: "", orders: {} };
    }
  }

  function save(state) {
    localStorage.setItem(
      KEY,
      JSON.stringify({
        templateId: state.templateId || "",
        orders: state.orders || {},
      })
    );
  }

  root.Store = { load, save };
})(typeof self !== "undefined" ? self : globalThis);
