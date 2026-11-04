// 任务窗格：只读当前 BOM 工作表，自动匹配表头并预览数据。
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

  function showError(message) {
    const box = $("error-summary");
    box.textContent = message;
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

  function formatValue(value) {
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? "" : value.toISOString().slice(0, 10);
    }
    return value === null || value === undefined ? "" : String(value);
  }

  // 读取当前活动 Sheet 的元信息。
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

  // 模板选择只保留当前内置 BOM 模板，不再提供自定义模板管理。
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

  // 自动匹配结果只读展示；下一阶段再增加手工选择目标字段。
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
    summary.textContent = `文本 ${typeCounts.text} · 数字 ${typeCounts.number} · 日期 ${typeCounts.date} · 空值 ${typeCounts.empty}`;
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

  // 当前阶段只保留上传入口，真实接口将在下一阶段接入。
  function renderImportState() {
    const button = $("btn-one-click");
    const hasTemplate = !!currentTemplate();
    button.disabled = true;
    $("import-hint").textContent = hasTemplate
      ? "BOM 字段已匹配，上传接口将在下一步接入"
      : "请先选择 BOM 模板";
  }

  function renderAll() {
    renderSheetInfo();
    renderSourceHeaders();
    renderTemplateBar();
    renderMapping();
    renderPreview();
    renderImportState();
  }

  // 从适配器读取表头和前 20 行，浏览器和 WPS 使用同一套展示逻辑。
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
    renderAll();
    toast("已加载 BOM 模板");
  }

  function init(templates, adapter) {
    UI.templates = templates;
    UI.adapter = adapter;

    $("tpl-select").addEventListener("change", (event) => {
      if (event.target.value) activateTemplate(event.target.value);
    });
    $("btn-one-click").addEventListener("click", () => {
      toast("上传接口将在下一步接入");
    });

    if (adapter.onDataChange) {
      adapter.onDataChange(() => loadPreviewData());
    }

    $("env-tag").textContent = adapter.isMock ? "浏览器 Mock 模式" : "WPS 环境";
    renderTemplateBar();
    loadPreviewData();
  }

  root.UI = { init };
})(typeof self !== "undefined" ? self : globalThis);
