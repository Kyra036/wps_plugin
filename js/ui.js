// 任务窗格：读取当前 BOM 工作表，匹配字段，并选择本次导入的目标模具。
(function (root) {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const PREVIEW_ROW_LIMIT = 20;

  const UI = {
    templates: [],
    fields: [],
    currentTemplateId: "",
    adapter: null,
    headerInfo: null,
    previewRows: [],
    totalDataRows: 0,
    previewRequestId: 0,
    molds: [],
    selectedMoldId: "",
    connection: { baseUrl: "", token: "" },
    busy: false,
  };

  function esc(value) {
    return String(value).replace(/[&<>"']/g, (character) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;",
    }[character]));
  }

  let toastTimer = null;
  function toast(message) {
    const element = $("toast");
    element.textContent = message;
    element.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => element.classList.remove("show"), 2500);
  }

  // 错误区同时支持单条消息和多条行级校验错误。
  function showError(message) {
    const box = $("error-summary");
    const messages = Array.isArray(message) ? message : [message];
    if (messages.length > 1) {
      box.innerHTML = `<ul>${messages.map((item) => `<li>${esc(item)}</li>`).join("")}</ul>`;
    } else {
      box.textContent = messages[0] || "发生未知错误";
    }
    box.classList.remove("hidden");
    box.focus();
  }

  function clearErrors() {
    const box = $("error-summary");
    box.textContent = "";
    box.classList.add("hidden");
  }

  function currentTemplate() {
    return UI.templates.find((template) => template.id === UI.currentTemplateId) || null;
  }

  function selectedMold() {
    return UI.molds.find((mold) => String(mold.moldId) === String(UI.selectedMoldId)) || null;
  }

  function formatValue(value) {
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? "" : value.toISOString().slice(0, 10);
    }
    return value === null || value === undefined ? "" : String(value);
  }

  // 读取当前活动 Sheet 的范围信息，帮助用户确认没有读错工作表。
  function renderSheetInfo() {
    const element = $("sheet-info");
    try {
      const info = UI.adapter.getSheetInfo();
      if (!info || !info.rowCount || !info.columnCount) {
        element.textContent = "当前工作表为空";
        return;
      }
      const startColumn = Core.colLetter(info.startColumn - 1);
      const endColumn = Core.colLetter(info.startColumn + info.columnCount - 2);
      const startRow = info.startRow;
      const endRow = info.startRow + info.rowCount - 1;
      element.innerHTML =
        `<span class="label">Sheet</span><span class="value">${esc(info.name)}</span>` +
        `<span class="label">使用范围</span><span class="value">${startColumn}${startRow}:${endColumn}${endRow}</span>` +
        `<span class="label">区域大小</span><span class="value">${info.rowCount} 行 × ${info.columnCount} 列</span>`;
    } catch (error) {
      element.textContent = `读取工作表信息失败：${error.message}`;
    }
  }

  // 源表头只显示列字母和表头名称。
  function renderSourceHeaders() {
    const list = $("source-headers");
    const count = $("header-count");
    const headers = UI.headerInfo && Array.isArray(UI.headerInfo.headers)
      ? UI.headerInfo.headers
      : [];
    count.textContent = ` ${headers.length} 列`;
    list.innerHTML = "";

    if (headers.length === 0) {
      list.className = "dim";
      list.textContent = "当前工作表没有可读取的表头";
      return;
    }

    list.className = "";
    headers.forEach((header, index) => {
      const chip = document.createElement("span");
      const label = header.fullLabel || header.label || "";
      chip.className = "header-chip" + (label ? "" : " empty");
      chip.title = label;

      const column = document.createElement("span");
      column.className = "col";
      column.textContent = header.column || Core.colLetter(index);
      chip.append(column, document.createTextNode(label || "（空表头）"));
      list.appendChild(chip);
    });
  }

  // 模板选择只保留当前内置 BOM 模板，默认为“请选择模板”。
  function renderTemplateBar() {
    const select = $("tpl-select");
    const placeholder = '<option value="" disabled' +
      (UI.currentTemplateId ? "" : " selected") + ">请选择 BOM 模板</option>";
    select.innerHTML = placeholder + UI.templates
      .map((template) =>
        `<option value="${esc(template.id)}"${template.id === UI.currentTemplateId ? " selected" : ""}>` +
        `${esc(template.name)}</option>`
      )
      .join("");
    select.value = UI.currentTemplateId;
    select.disabled = UI.busy;
    $("tpl-hint").textContent = currentTemplate()
      ? "自动匹配当前源表头与 BOM 字段"
      : "选择模板后开始自动匹配";
  }

  // 自动匹配结果只读展示，真正上传时按这组匹配关系生成平台文件。
  function renderMapping() {
    const section = $("mapping-section");
    const list = $("mapping-list");
    const stats = $("mapping-stats");
    const hint = $("mapping-hint");
    const template = currentTemplate();
    list.innerHTML = "";

    if (!template) {
      section.classList.add("hidden");
      return;
    }

    const result = Core.autoMatchHeaders(
      UI.headerInfo ? UI.headerInfo.headers : [],
      UI.fields
    );
    section.classList.remove("hidden");
    stats.textContent = `${result.matchedCount}/${result.matches.length} 列已匹配`;

    result.matches.forEach((match) => {
      const row = document.createElement("div");
      row.className = "mapping-row" + (match.field ? "" : " unmatched");
      row.title = match.reason;

      const source = document.createElement("span");
      source.className = "mapping-source";
      const column = document.createElement("span");
      column.className = "col";
      column.textContent = match.header.column || Core.colLetter(match.headerIndex);
      source.append(column, document.createTextNode(
        match.header.fullLabel || match.header.label || ""
      ));

      const arrow = document.createElement("span");
      arrow.className = "mapping-arrow";
      arrow.textContent = "→";

      const target = document.createElement("span");
      target.className = "mapping-target";
      if (match.field) {
        const group = document.createElement("span");
        group.className = "group";
        group.textContent = match.field.groupLabels.length
          ? `${match.field.groupLabels.join(" / ")} / `
          : "";
        target.append(group, document.createTextNode(match.field.label));
      } else {
        target.textContent = "未匹配";
      }

      row.append(source, arrow, target);
      list.appendChild(row);
    });

    const unmatchedSourceCount = result.matches.length - result.matchedCount;
    const messages = [];
    if (unmatchedSourceCount > 0) messages.push(`${unmatchedSourceCount} 个源表头未匹配`);
    if (result.unmatchedTemplateFields.length > 0) {
      messages.push(`${result.unmatchedTemplateFields.length} 个 BOM 字段没有源列`);
    }
    hint.textContent = messages.length ? messages.join("；") : "全部源表头已自动匹配";
  }

  // 只预览前 20 行，避免大表格一次性渲染过多 DOM。
  function renderPreview() {
    const panel = $("preview-panel");
    const stats = $("preview-stats");
    const headers = UI.headerInfo ? UI.headerInfo.headers : [];
    const rows = UI.previewRows;
    const columnCount = Math.max(headers.length, ...rows.map((row) => row.length), 0);
    stats.textContent = ` 共 ${UI.totalDataRows} 行，显示前 ${rows.length} 行`;
    panel.className = "";
    panel.innerHTML = "";

    if (headers.length === 0) {
      panel.className = "dim";
      panel.textContent = "没有可预览的数据";
      return;
    }

    const typeCounts = { text: 0, number: 0, date: 0, boolean: 0, empty: 0 };
    rows.forEach((row) => {
      for (let column = 0; column < columnCount; column++) {
        typeCounts[Core.classifyCell(row[column])]++;
      }
    });

    const summary = document.createElement("div");
    summary.className = "preview-summary";
    summary.textContent =
      `文本 ${typeCounts.text} · 数字 ${typeCounts.number} · 日期 ${typeCounts.date} · 空值 ${typeCounts.empty}`;
    panel.appendChild(summary);

    const table = document.createElement("table");
    table.className = "preview";
    const columnHeader = document.createElement("tr");
    columnHeader.appendChild(document.createElement("th"));
    for (let column = 0; column < columnCount; column++) {
      const th = document.createElement("th");
      th.textContent = Core.colLetter(column);
      columnHeader.appendChild(th);
    }
    table.appendChild(columnHeader);

    const headerRow = document.createElement("tr");
    const headerNumber = document.createElement("td");
    headerNumber.className = "row-no";
    headerNumber.textContent = UI.headerInfo.dataStartRow - 1;
    headerRow.appendChild(headerNumber);
    for (let column = 0; column < columnCount; column++) {
      const cell = document.createElement("th");
      cell.textContent = headers[column] ? (headers[column].fullLabel || headers[column].label) : "";
      headerRow.appendChild(cell);
    }
    table.appendChild(headerRow);

    rows.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      const numberCell = document.createElement("td");
      numberCell.className = "row-no";
      numberCell.textContent = UI.headerInfo.dataStartRow + rowIndex;
      tr.appendChild(numberCell);
      for (let column = 0; column < columnCount; column++) {
        const cell = document.createElement("td");
        cell.textContent = formatValue(row[column]);
        tr.appendChild(cell);
      }
      table.appendChild(tr);
    });

    panel.appendChild(table);
  }

  // 首次打开若已有令牌，自动收起连接区，让主流程保持简洁。
  function renderConnection() {
    const config = BomApi.loadConfig();
    UI.connection = config;
    $("api-base").value = config.baseUrl;
    $("api-token").value = config.token;
    const connected = !!config.token;
    $("connection-status").textContent = connected ? "已配置" : "未连接";
    $("connection-hint").textContent = connected
      ? "平台连接已保存在当前 WPS 本地"
      : "令牌只保存在当前 WPS 本地";
    if (connected) $("connection-panel").open = false;
  }

  function saveConnection() {
    const config = BomApi.saveConfig({
      baseUrl: $("api-base").value,
      token: $("api-token").value,
    });
    UI.connection = config;
    $("api-base").value = config.baseUrl;
    $("api-token").value = config.token;
    $("connection-status").textContent = config.token ? "已配置" : "未连接";
    $("connection-panel").open = !config.token;
    clearErrors();
    toast(config.token ? "平台连接已保存" : "请填写访问令牌");
    renderImportState();
  }

  function renderMoldOptions() {
    const select = $("mold-select");
    select.innerHTML = "";
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = UI.molds.length ? "请选择目标模具" : "没有匹配的模具";
    select.appendChild(placeholder);

    UI.molds.forEach((mold) => {
      const option = document.createElement("option");
      option.value = String(mold.moldId);
      option.textContent = [mold.moldNo, mold.moldName]
        .filter((value) => value !== undefined && value !== null && String(value).trim())
        .join(" / ");
      select.appendChild(option);
    });
    select.disabled = UI.molds.length === 0 || UI.busy;
    select.value = UI.selectedMoldId;
  }

  // 当前步骤只验证连接和目标模具选择，上传入口留到下一步接入。
  function renderImportState() {
    const button = $("btn-one-click");
    const hint = $("import-hint");
    const template = currentTemplate();
    const mapping = template
      ? Core.autoMatchHeaders(UI.headerInfo ? UI.headerInfo.headers : [], UI.fields)
      : null;
    const missingFields = mapping ? mapping.unmatchedTemplateFields.length : UI.fields.length;
    const mold = selectedMold();

    let message = "";
    if (!template) message = "请先选择 BOM 模板";
    else if (missingFields > 0) message = `仍有 ${missingFields} 个 BOM 字段没有源列`;
    else if (!UI.connection.token) message = "请先填写并保存平台访问令牌";
    else if (!mold) message = "请先查询并选择目标模具";
    else message = `已选择目标模具：${mold.moldNo || mold.moldName || "未命名模具"}`;

    button.disabled = true;
    button.textContent = "一键导入 BOM";
    hint.textContent = message;
  }

  function renderAll() {
    renderSheetInfo();
    renderSourceHeaders();
    renderTemplateBar();
    renderMoldOptions();
    renderMapping();
    renderPreview();
    renderImportState();
  }

  // 从适配器读取表头和前 20 行，浏览器和 WPS 共用同一套展示逻辑。
  async function loadPreviewData() {
    const requestId = ++UI.previewRequestId;
    try {
      clearErrors();
      const headerInfo = UI.adapter.getHeaderInfo();
      const columnCount = headerInfo.headers ? headerInfo.headers.length : 0;
      const rows = columnCount > 0
        ? await UI.adapter.readRows(columnCount, PREVIEW_ROW_LIMIT)
        : [];
      if (requestId !== UI.previewRequestId) return;

      const info = UI.adapter.getSheetInfo();
      UI.headerInfo = headerInfo;
      UI.previewRows = rows;
      UI.totalDataRows = Math.max(0, info.rowCount - (headerInfo.headerRows || 1));
      renderAll();
    } catch (error) {
      if (requestId !== UI.previewRequestId) return;
      showError(`读取当前工作表失败：${error.message}`);
      renderSheetInfo();
    }
  }

  function activateTemplate(templateId) {
    const template = UI.templates.find((item) => item.id === templateId);
    if (!template) return;
    UI.currentTemplateId = template.id;
    UI.fields = Core.flattenTemplate(template);
    clearErrors();
    renderAll();
    toast("已加载 BOM 模板");
  }

  // 查询模具时把按钮暂时置为忙碌，避免用户连续点击造成重复请求。
  async function searchMolds() {
    const keyword = $("mold-keyword").value.trim();
    if (!UI.connection.token) {
      $("connection-panel").open = true;
      showError("请先填写并保存平台访问令牌");
      return;
    }

    const button = $("btn-search-mold");
    UI.busy = true;
    UI.selectedMoldId = "";
    button.disabled = true;
    button.textContent = "查询中";
    $("mold-select").disabled = true;
    clearErrors();
    renderImportState();
    try {
      UI.molds = await BomApi.searchMolds(keyword);
      if (UI.molds.length === 1) UI.selectedMoldId = String(UI.molds[0].moldId);
      renderMoldOptions();
      if (UI.molds.length === 0) {
        showError("没有找到匹配的模具，请检查模具编号或名称");
      } else if (UI.molds.length === 1) {
        toast("已选择唯一匹配的模具");
      } else {
        toast(`找到 ${UI.molds.length} 个模具，请选择`);
      }
    } catch (error) {
      UI.molds = [];
      renderMoldOptions();
      showError(error.message);
    } finally {
      UI.busy = false;
      button.disabled = false;
      button.textContent = "查询";
      renderAll();
    }
  }

  function init(templates, adapter) {
    UI.templates = templates;
    UI.adapter = adapter;

    $("tpl-select").addEventListener("change", (event) => {
      if (event.target.value) activateTemplate(event.target.value);
    });
    $("btn-save-connection").addEventListener("click", saveConnection);
    $("btn-search-mold").addEventListener("click", searchMolds);
    $("mold-keyword").addEventListener("keydown", (event) => {
      if (event.key === "Enter") searchMolds();
    });
    $("mold-select").addEventListener("change", (event) => {
      UI.selectedMoldId = event.target.value;
      renderImportState();
    });
    if (adapter.onDataChange) {
      adapter.onDataChange(() => loadPreviewData());
    }

    $("env-tag").textContent = adapter.isMock ? "浏览器 Mock 模式" : "WPS 环境";
    renderConnection();
    renderTemplateBar();
    loadPreviewData();
  }

  root.UI = { init };
})(typeof self !== "undefined" ? self : globalThis);
