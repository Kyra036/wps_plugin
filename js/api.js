// 平台接口层：统一处理地址、登录令牌、模具查询、零件查重和 BOM 上传。
(function (root) {
  "use strict";

  const STORAGE_KEY = "wps_bom_platform_config_v1";
  const DEFAULT_BASE_URL = "http://47.99.212.121:9001/prod-api";
  const DEFAULT_CLIENT_ID = "e5cd7e4891bf95d1d19206ce24a0550e";
  const REQUEST_TIMEOUT = 30000;

  function normalizeBaseUrl(value) {
    const text = String(value || "").trim().replace(/\/+$/, "");
    if (!text) return DEFAULT_BASE_URL;
    if (/\/(prod-api|dev-api)$/.test(text)) return text;
    return `${text}/prod-api`;
  }

  function normalizeToken(value) {
    return String(value || "").trim().replace(/^Bearer\s+/i, "");
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
    const query = encodeURIComponent(String(keyword || "").trim());
    const path = `/tenant/v2/bom/getBomMoldList${query ? `?moldNameOrNo=${query}` : ""}`;
    const payload = await request(path);
    return Array.isArray(payload && payload.data) ? payload.data : [];
  }

  root.BomApi = {
    DEFAULT_BASE_URL,
    loadConfig,
    saveConfig,
    normalizeBaseUrl,
    normalizeToken,
    searchMolds,
  };
})(typeof self !== "undefined" ? self : globalThis);
