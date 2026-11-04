// 启动装配：加载内置模板库和 Mock 工作表，探测运行环境后初始化任务窗格。
(function () {
  "use strict";

  function isWpsEnv() {
    return typeof window.wps !== "undefined" && typeof window.wps.EtApplication === "function";
  }

  // 所有静态数据都通过 HTTP 加载，避免 file:// 环境下 fetch 被浏览器拦截。
  async function loadJson(path) {
    const res = await fetch(path);
    if (!res.ok) throw new Error(`${path}: HTTP ${res.status}`);
    return res.json();
  }

  async function boot() {
    const wpsMode = isWpsEnv();
    let templateLibrary;
    let mock;
    try {
      // 模板库与示例 Sheet 分离；真实 WPS 环境不需要加载 Mock 工作表。
      [templateLibrary, mock] = await Promise.all([
        loadJson("data/templates.json"),
        wpsMode ? Promise.resolve(null) : loadJson("data/mock.json"),
      ]);
    } catch (err) {
      const box = document.getElementById("error-summary");
      box.classList.remove("hidden");
      box.textContent =
        `加载模板库或示例工作表失败：${err.message}（请通过 HTTP 服务访问页面）`;
      return;
    }

    const templates = templateLibrary && Array.isArray(templateLibrary.templates)
      ? templateLibrary.templates
      : [];
    if (templates.length === 0) {
      const box = document.getElementById("error-summary");
      box.classList.remove("hidden");
      box.textContent = "data/templates.json 中没有可用模板（templates 为空）";
      return;
    }

    const adapter = wpsMode
      ? new Adapters.WpsAdapter()
      : new Adapters.MockAdapter((mock && mock.mockSheet) || { name: "Sheet1", rows: [] });
    UI.init(templates, adapter);
  }

  document.addEventListener("DOMContentLoaded", boot);
})();
