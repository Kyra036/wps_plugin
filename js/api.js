// 平台接口层：统一处理地址、登录令牌、模具查询和只读零件信息。
(function (root) {
  "use strict";

  const STORAGE_KEY = "wps_bom_platform_config_v1";
  const DEFAULT_BASE_URL = "http://47.99.212.121:9001";
  const DEFAULT_CLIENT_ID = "e5cd7e4891bf95d1d19206ce24a0550e";
  const REQUEST_TIMEOUT = 30000;
  const DEMO_MOLDS = [
    {
      moldId: "2114535348499886082",
      moldNo: "M-WPS-DEMO-001",
      moldName: "BOM导入演示模具",
    },
  ];
  const DEMO_PART_COUNT = 2;

  function normalizeBaseUrl(value) {
    const text = String(value || "").trim().replace(/\/+$/, "");
    if (!text) return DEFAULT_BASE_URL;
    // 兼容上一版已经保存到本地的默认地址，避免用户重复修改连接配置。
    if (/^http:\/\/47\.99\.212\.121:9001\/(prod-api|dev-api)$/i.test(text)) {
      return DEFAULT_BASE_URL;
    }
    return text;
  }

  function normalizeToken(value) {
    return String(value || "").trim().replace(/^Bearer\s+/i, "");
  }

  // 默认使用演示数据；后续接好自动登录后可通过 ?mode=platform 切回真实接口。
  function isDemoMode() {
    try {
      return new URLSearchParams(location.search).get("mode") !== "platform";
    } catch (_) {
      return true;
    }
  }

  // localStorage 在极少数 WPS 内嵌环境可能不可用，因此保留内存配置兜底。
  let memoryConfig = null;

  function loadConfig() {
    if (memoryConfig) return { ...memoryConfig };
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
      if (saved && typeof saved === "object") {
        memoryConfig = {
          baseUrl: normalizeBaseUrl(saved.baseUrl),
          token: normalizeToken(saved.token),
        };
        return { ...memoryConfig };
      }
    } catch (_) {
      // 配置损坏时回退到默认值，避免整个插件无法打开。
    }
    memoryConfig = { baseUrl: DEFAULT_BASE_URL, token: "" };
    return { ...memoryConfig };
  }

  function saveConfig(config) {
    memoryConfig = {
      baseUrl: normalizeBaseUrl(config && config.baseUrl),
      token: normalizeToken(config && config.token),
    };
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(memoryConfig));
    } catch (_) {
      // 存储失败时仍可支持当前窗口继续使用。
    }
    return { ...memoryConfig };
  }

  function requestHeaders(config, jsonBody) {
    const headers = {
      clientid: DEFAULT_CLIENT_ID,
      Authorization: `Bearer ${config.token}`,
    };
    if (jsonBody) headers["Content-Type"] = "application/json;charset=utf-8";
    return headers;
  }

  function responseMessage(payload, status) {
    if (payload && typeof payload === "object") {
      return payload.msg || payload.message || `接口请求失败（HTTP ${status}）`;
    }
    return `接口请求失败（HTTP ${status}）`;
  }

  // 所有接口共用同一个请求入口，便于统一处理超时、登录失效和业务错误码。
  async function request(path, options) {
    const config = loadConfig();
    if (!config.token) throw new Error("请先填写平台访问令牌");

    const requestOptions = options || {};
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT);
    try {
      const response = await fetch(`${config.baseUrl}${path}`, {
        method: requestOptions.method || "GET",
        headers: requestHeaders(config, !!requestOptions.json),
        body: requestOptions.json
          ? JSON.stringify(requestOptions.json)
          : requestOptions.body,
        signal: controller.signal,
        mode: "cors",
      });

      const text = await response.text();
      let payload = null;
      try {
        payload = text ? JSON.parse(text) : null;
      } catch (_) {
        throw new Error(`平台返回的不是 JSON（HTTP ${response.status}）`);
      }

      if (!response.ok) {
        throw new Error(responseMessage(payload, response.status));
      }
      if (payload && typeof payload.code !== "undefined" && Number(payload.code) !== 200) {
        throw new Error(responseMessage(payload, response.status));
      }
      return payload;
    } catch (error) {
      if (error && error.name === "AbortError") {
        throw new Error("平台请求超时，请检查网络或平台地址");
      }
      if (error instanceof TypeError) {
        throw new Error("无法连接平台，请检查平台地址、网络和跨域配置");
      }
      throw error;
    } finally {
      clearTimeout(timer);
    }
  }

  // 模具查询同时支持编号和名称，用户不知道内部 moldId 也能找到目标模具。
  async function searchMolds(keyword) {
    const text = String(keyword || "").trim().toLowerCase();
    if (isDemoMode()) {
      return DEMO_MOLDS.filter((mold) =>
        !text ||
        String(mold.moldNo).toLowerCase().includes(text) ||
        String(mold.moldName).toLowerCase().includes(text)
      );
    }

    const query = encodeURIComponent(text);
    const path = `/tenant/v2/bom/getBomMoldList${query ? `?moldNameOrNo=${query}` : ""}`;
    const payload = await request(path);
    return Array.isArray(payload && payload.data) ? payload.data : [];
  }

  // 只读取分页总数，不加载零件明细；当前步骤不会修改目标模具。
  async function countParts(moldId) {
    if (isDemoMode()) {
      return String(moldId) === DEMO_MOLDS[0].moldId ? DEMO_PART_COUNT : 0;
    }

    if (moldId === undefined || moldId === null || String(moldId).trim() === "") {
      throw new Error("缺少目标模具 ID");
    }

    const payload = await request("/workOrder/advanced/list", {
      method: "POST",
      json: {
        moldId,
        queryRequest: {
          pageNum: 1,
          pageSize: 1,
          sort: [{ field: "createTime", order: "asc" }],
          filter: [],
          search: { fields: ["partName", "partNo"], keyword: "" },
        },
      },
    });
    const data = payload && payload.data;
    const total = Number(data && data.total);
    if (Number.isFinite(total)) return Math.max(0, total);
    if (Array.isArray(data && data.records)) return data.records.length;
    return 0;
  }

  root.BomApi = {
    DEFAULT_BASE_URL,
    loadConfig,
    saveConfig,
    normalizeBaseUrl,
    normalizeToken,
    isDemoMode,
    searchMolds,
    countParts,
  };
})(typeof self !== "undefined" ? self : globalThis);
