// 启动装配：加载 mock 模板与模拟工作表，探测运行环境，初始化任务窗格。
(function () {
  "use strict";

  function isWpsEnv() {
    return typeof window.wps !== "undefined" && typeof window.wps.EtApplication === "function";
  }

  async function boot() {
    let mock;
    try {
      const res = await fetch("data/mock.json");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      mock = await res.json();
    } catch (err) {
      const box = document.getElementById("error-summary");
      box.classList.remove("hidden");
      box.textContent =
        `加载 data/mock.json 失败：${err.message}（请通过 http 服务访问页面，而不是 file:// 直接打开）`;
      return;
    }

    const templates = Array.isArray(mock.templates) ? mock.templates : [];
    if (templates.length === 0) {
      const box = document.getElementById("error-summary");
      box.classList.remove("hidden");
      box.textContent = "data/mock.json 中没有可用模板（templates 为空）";
      return;
    }

    const adapter = isWpsEnv()
      ? new Adapters.WpsAdapter()
      : new Adapters.MockAdapter(mock.mockSheet || { name: "Sheet1", rows: [] });
    UI.init(templates, adapter);
  }

  document.addEventListener("DOMContentLoaded", boot);
})();
