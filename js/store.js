// 本地持久化：只保存每个模板下的字段顺序，不恢复上次选中的模板。
// 模板定义本身始终来自 data/mock.json，插件不写入、不修改。
(function (root) {
  "use strict";

  const KEY = "wps-field-mapping/v2";

  function load() {
    try {
      const raw = localStorage.getItem(KEY);
      if (!raw) return { orders: {} };
      const data = JSON.parse(raw);
      return {
        orders: data.orders && typeof data.orders === "object" ? data.orders : {},
      };
    } catch (_) {
      return { orders: {} };
    }
  }

  function save(state) {
    localStorage.setItem(
      KEY,
      JSON.stringify({
        orders: state.orders || {},
      })
    );
  }

  root.Store = { load, save };
})(typeof self !== "undefined" ? self : globalThis);
