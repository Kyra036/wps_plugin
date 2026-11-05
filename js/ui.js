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
    headerRow: null,
    previewRows: [],
    totalDataRows: 0,
    previewRequestId: 0,
    molds: [],
    selectedMoldId: "",
    existingParts: [],
    existingPartsLoaded: false,
    existingPartsLoading: false,
    existingPartsRequestId: 0,
    connection: { token: "" },
    busy: false,
    validationBusy: false,
    generateBusy: false,
    generatedFile: null,
    generatedRows: [],
    importBusy: false,
    uploadCompleted: false,
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

  // BOM 检查结果只保留在页面中，不会发送到平台。
  function setValidationStatus(message, type) {
    const element = $("validation-status");
    element.textContent = message || "";
    element.className = message ? (type || "") : "hidden";
  }

  function setGenerateStatus(message, type) {
    const element = $("generate-status");
    element.textContent = message || "";
    element.className = message ? (type || "") : "hidden";
  }

  function setImportStatus(message, type) {
    const element = $("import-status");
    element.textContent = message || "";
    element.className = message ? (type || "") : "hidden";
  }

  function resetGeneratedFile() {
    UI.generatedFile = null;
    UI.generatedRows = [];
    UI.uploadCompleted = false;
    setGenerateStatus("");
    setImportStatus("");
    renderGeneratedPreview();
  }

  // 页面预览严格模拟临时 XLSX：第 5 行表头，之后逐行展示 BOM 数据。
  function renderGeneratedPreview() {
    const panel = $("generated-preview");
    const table = $("generated-preview-table");
    table.innerHTML = "";
    if (!UI.generatedFile) {
      panel.classList.add("hidden");
      return;
    }

    const headers = ["Excel 行", "序号", "名称", "材料", "规格", "数量", "备注", "加工要求"];
    const keys = [
      "partNo",
      "partName",
      "partMaterial",
      "partSpecs",
      "partNum",
      "partRemark",
      "partProcessReq",
    ];
    const headerRow = document.createElement("tr");
    headers.forEach((label) => {
      const cell = document.createElement("th");
      cell.textContent = label;
      headerRow.appendChild(cell);
    });
    table.appendChild(headerRow);

    const platform = currentTemplate() && currentTemplate().platform;
    const excelHeaderRow = Number(platform && platform.headerRow) || 5;
    const headerContent = document.createElement("tr");
    const headerNumber = document.createElement("td");
    headerNumber.className = "row-no";
    headerNumber.textContent = excelHeaderRow;
    headerContent.appendChild(headerNumber);
    ["序号", "名称", "材料", "规格（宽*长*高）（MM）", "数量", "备注", "加工要求"]
      .forEach((value) => {
        const cell = document.createElement("td");
        cell.textContent = value;
        headerContent.appendChild(cell);
      });
    table.appendChild(headerContent);

    UI.generatedRows.forEach((row, index) => {
      const tr = document.createElement("tr");
      const rowNumber = document.createElement("td");
      rowNumber.className = "row-no";
      rowNumber.textContent = excelHeaderRow + index + 1;
      tr.appendChild(rowNumber);
      keys.forEach((key) => {
        const cell = document.createElement("td");
        cell.textContent = row[key] === undefined || row[key] === null ? "" : String(row[key]);
        tr.appendChild(cell);
      });
      table.appendChild(tr);
    });
    panel.classList.remove("hidden");
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

  function renderHeaderRowControl() {
    const input = $("source-header-row");
    if (document.activeElement !== input && UI.headerRow) {
      input.value = UI.headerRow;
    }
    $("source-header-row-hint").textContent = UI.headerRow
      ? `当前按第 ${UI.headerRow} 行作为源表头，数据从第 ${UI.headerRow + 1} 行开始`
      : "请输入当前工作表的表头所在行";
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

  // 开发阶段只暴露令牌；真实接口地址固定使用 HTTPS 正式网关。
  function renderConnection() {
    const config = BomApi.loadConfig();
    UI.connection = { token: config.token };
    $("dev-token").value = config.token;
    $("connection-status").textContent = config.token ? "已配置" : "未配置";
    $("connection-hint").textContent = config.token
      ? "令牌已保存在当前浏览器本地"
      : "令牌只保存在当前浏览器本地";
    if (config.token) $("connection-panel").open = false;
  }

  function saveToken() {
    const config = BomApi.saveConfig({
      baseUrl: BomApi.DEFAULT_BASE_URL,
      token: $("dev-token").value,
    });
    UI.connection = { token: config.token };
    $("dev-token").value = config.token;
    $("connection-status").textContent = config.token ? "已配置" : "未配置";
    $("connection-hint").textContent = config.token
      ? "令牌已保存在当前浏览器本地"
      : "令牌只保存在当前浏览器本地";
    $("connection-panel").open = !config.token;
    clearErrors();
    toast(config.token ? "开发令牌已保存" : "请填写平台访问令牌");
    renderImportState();
  }

  // 目标模具下方只展示零件总数；读取失败时显示错误，不影响当前工作表预览。
  function renderMoldPartStatus() {
    const element = $("mold-part-status");
    const mold = selectedMold();
    if (!mold) {
      element.textContent = "选择模具后读取已有零件数量";
      return;
    }
    if (UI.existingPartsLoading) {
      element.textContent = "正在读取目标模具已有零件...";
      return;
    }
    if (!UI.existingPartsLoaded) {
      element.textContent = "尚未读取已有零件";
      return;
    }
    if (UI.existingParts.length > 0) {
      element.textContent = `目标模具已有 ${UI.existingParts.length} 个零件`;
      return;
    }
    element.textContent = "目标模具没有已有零件";
  }

  // 只展示前 20 条已有零件，完整数组保留在内存中供后续功能使用。
  function renderExistingPartsPreview() {
    const element = $("existing-parts-preview");
    element.innerHTML = "";
    if (!selectedMold() || UI.existingPartsLoading || UI.existingParts.length === 0) {
      element.classList.add("hidden");
      return;
    }

    element.classList.remove("hidden");
    UI.existingParts.slice(0, 20).forEach((part) => {
      const row = document.createElement("div");
      row.className = "existing-part-item";
      const partNo = document.createElement("span");
      partNo.className = "part-no";
      partNo.textContent = part.partNo || "（无编号）";
      partNo.title = part.partNo || "";
      const partName = document.createElement("span");
      partName.textContent = part.partName || "（无名称）";
      partName.title = part.partName || "";
      row.append(partNo, partName);
      element.appendChild(row);
    });

    if (UI.existingParts.length > 20) {
      const more = document.createElement("div");
      more.className = "existing-part-more";
      more.textContent = `还有 ${UI.existingParts.length - 20} 个零件未显示`;
      element.appendChild(more);
    }
  }

  // 只调用查询接口读取零件，本函数不包含新增、更新或删除操作。
  async function loadExistingParts(moldId) {
    const requestId = ++UI.existingPartsRequestId;
    UI.existingPartsLoading = true;
    UI.existingParts = [];
    UI.existingPartsLoaded = false;
    renderMoldPartStatus();
    renderExistingPartsPreview();
    try {
      const parts = await BomApi.listParts(moldId);
      if (requestId !== UI.existingPartsRequestId) return;
      UI.existingParts = parts;
      UI.existingPartsLoaded = true;
      return parts;
    } catch (error) {
      if (requestId !== UI.existingPartsRequestId) return;
      showError(`读取已有零件失败：${error.message}`);
      return null;
    } finally {
      if (requestId === UI.existingPartsRequestId) {
        UI.existingPartsLoading = false;
        renderMoldPartStatus();
        renderExistingPartsPreview();
        renderImportState();
      }
    }
  }

  // 所有选择模具的入口都走这里，确保加载状态和请求序号同时更新。
  function selectMold(moldId) {
    UI.selectedMoldId = moldId || "";
    UI.existingParts = [];
    UI.existingPartsLoaded = false;
    UI.existingPartsLoading = false;
    UI.existingPartsRequestId += 1;
    UI.uploadCompleted = false;
    setImportStatus("");
    renderMoldOptions();
    renderMoldPartStatus();
    renderExistingPartsPreview();
    renderImportState();
    if (UI.selectedMoldId) loadExistingParts(UI.selectedMoldId);
  }

  // 令牌和目标模具准备好后即可一键生成并上传临时文件。
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
    else if (!UI.connection.token) message = "请先在开发登录中配置平台令牌";
    else if (!mold) message = "请先查询并选择目标模具";
    else if (UI.generateBusy) message = "正在生成新的临时 XLSX...";
    else if (UI.uploadCompleted) message = "本次临时文件已上传，请重新生成后再上传";
    else if (UI.generatedFile) {
      message = `准备上传到：${mold.moldNo || mold.moldName || "未命名模具"}（不判断重复零件）`;
    } else {
      message = `点击后将生成临时 XLSX 并上传到：${mold.moldNo || mold.moldName || "未命名模具"}`;
    }

    button.disabled =
      !UI.connection.token ||
      !mold ||
      UI.generateBusy ||
      UI.uploadCompleted ||
      UI.importBusy;
    button.textContent = UI.importBusy ? "正在上传..." : "一键导入 BOM";
    hint.textContent = message;
  }

  // 完整检查只依赖模板和当前工作表，不要求配置令牌或选择目标模具。
  function renderValidationState() {
    const button = $("btn-validate-bom");
    const template = currentTemplate();
    const mapping = template
      ? Core.autoMatchHeaders(UI.headerInfo ? UI.headerInfo.headers : [], UI.fields)
      : null;
    const ready = !!template &&
      mapping.unmatchedTemplateFields.length === 0 &&
      UI.totalDataRows > 0;
    button.disabled = !ready || UI.validationBusy;
    button.textContent = UI.validationBusy ? "正在检查..." : "检查完整 BOM";
  }

  // 临时文件只依赖当前模板和源表数据，不要求令牌或目标模具。
  function renderGenerateState() {
    const button = $("btn-generate-xlsx");
    const template = currentTemplate();
    const mapping = template
      ? Core.autoMatchHeaders(UI.headerInfo ? UI.headerInfo.headers : [], UI.fields)
      : null;
    const ready = !!template &&
      mapping.unmatchedTemplateFields.length === 0 &&
      UI.totalDataRows > 0;
    button.disabled = !ready || UI.generateBusy || UI.importBusy;
    button.textContent = UI.generateBusy
      ? "正在生成..."
      : "生成临时 XLSX（暂不上传）";
  }

  function renderAll() {
    renderSheetInfo();
    renderSourceHeaders();
    renderHeaderRowControl();
    renderTemplateBar();
    renderMoldOptions();
    renderMoldPartStatus();
    renderExistingPartsPreview();
    renderMapping();
    renderPreview();
    renderValidationState();
    renderGenerateState();
    renderImportState();
  }

  // 从适配器读取表头和前 20 行，浏览器和 WPS 共用同一套展示逻辑。
  async function loadPreviewData(requestedHeaderRow) {
    const requestId = ++UI.previewRequestId;
    try {
      clearErrors();
      const info = UI.adapter.getSheetInfo();
      const defaultHeaderRow = UI.adapter.getDefaultHeaderRow
        ? UI.adapter.getDefaultHeaderRow()
        : (Number(info.startRow) || 1);
      const candidate = Number(requestedHeaderRow);
      const headerRow = Number.isInteger(candidate) && candidate > 0
        ? candidate
        : (UI.headerRow || defaultHeaderRow);
      const headerInfo = UI.adapter.getHeaderInfo(headerRow);
      const columnCount = headerInfo.headers ? headerInfo.headers.length : 0;
      const rows = columnCount > 0
        ? await UI.adapter.readRows(columnCount, PREVIEW_ROW_LIMIT, headerRow)
        : [];
      if (requestId !== UI.previewRequestId) return;

      const lastSheetRow = info.startRow + info.rowCount - 1;
      UI.headerInfo = headerInfo;
      UI.headerRow = headerInfo.headerRow || headerRow;
      UI.previewRows = rows;
      UI.totalDataRows = Math.max(0, lastSheetRow - headerInfo.dataStartRow + 1);
      setValidationStatus("");
      resetGeneratedFile();
      renderAll();
    } catch (error) {
      if (requestId !== UI.previewRequestId) return;
      showError(`读取当前工作表失败：${error.message}`);
      renderSheetInfo();
    }
  }

  // 用户确认表头行后重新读取列名、预览和校验状态。
  async function applyHeaderRow() {
    const input = $("source-header-row");
    const row = Number(input.value);
    const info = UI.adapter.getSheetInfo();
    const firstRow = Number(info.startRow) || 1;
    const lastRow = firstRow + Number(info.rowCount) - 1;
    if (!Number.isInteger(row) || row < firstRow || row > lastRow) {
      showError(`表头行必须是 ${firstRow} 到 ${lastRow} 之间的整数`);
      return;
    }
    setValidationStatus("");
    await loadPreviewData(row);
    toast(`已按第 ${row} 行读取表头`);
  }

  function activateTemplate(templateId) {
    const template = UI.templates.find((item) => item.id === templateId);
    if (!template) return;
    UI.currentTemplateId = template.id;
    UI.fields = Core.flattenTemplate(template);
    clearErrors();
    setValidationStatus("");
    resetGeneratedFile();
    renderAll();
    toast("已加载 BOM 模板");
  }

  // 读取全部源行并完成校验；该函数不调用任何平台写接口。
  async function readAndValidateBom() {
    const mapping = Core.autoMatchHeaders(
      UI.headerInfo ? UI.headerInfo.headers : [],
      UI.fields
    );
    if (mapping.unmatchedTemplateFields.length > 0) {
      throw new Error(`仍有 ${mapping.unmatchedTemplateFields.length} 个 BOM 字段没有源列`);
    }
    const sourceRows = await UI.adapter.readRows(
      UI.headerInfo.headers.length,
      Number.MAX_SAFE_INTEGER,
      UI.headerRow
    );
    const mappedRows = Core.buildMappedRows(mapping.matches, sourceRows, UI.fields);
    const validation = Core.validateImportRows(mappedRows, UI.headerInfo.dataStartRow);
    return { sourceRows, mappedRows, validation };
  }

  async function checkBom() {
    if (UI.validationBusy) return;
    clearErrors();
    setValidationStatus("正在读取并检查全部 BOM 行...");
    UI.validationBusy = true;
    renderValidationState();
    try {
      const result = await readAndValidateBom();
      if (result.validation.errors.length > 0) {
        showError(result.validation.errors);
        setValidationStatus(
          `检查完成：${result.validation.rows.length} 行有效，${result.validation.errors.length} 行有错误`,
          "error"
        );
      } else {
        setValidationStatus(`检查通过：${result.validation.rows.length} 行可导入`, "ok");
        toast("BOM 检查通过");
      }
    } catch (error) {
      showError(`BOM 检查失败：${error.message}`);
      setValidationStatus("检查失败，请根据错误提示修改后重试", "error");
    } finally {
      UI.validationBusy = false;
      renderValidationState();
    }
  }

  // 生成操作独立成函数，供“预览文件”和“一键导入”共同复用。
  async function buildTemporaryFile() {
    const result = await readAndValidateBom();
    if (result.validation.errors.length > 0) {
      const error = new Error("BOM 数据校验失败");
      error.validationErrors = result.validation.errors;
      throw error;
    }

    const template = currentTemplate();
    const fileName = `BOM-${Date.now()}.xlsx`;
    const file = XlsxBuilder.buildBomFile(
      result.validation.rows,
      fileName,
      template.platform
    );
    return { file, rows: result.validation.rows.slice() };
  }

  function applyTemporaryFile(generated) {
    UI.generatedFile = generated.file;
    UI.generatedRows = generated.rows;
    renderGeneratedPreview();
  }

  // 单独生成时展示文件大小，不上传。
  async function generateTemporaryFile() {
    if (UI.generateBusy || UI.importBusy) return;
    clearErrors();
    setValidationStatus("");
    setImportStatus("");
    UI.uploadCompleted = false;
    setGenerateStatus("正在读取并校验 BOM...");
    UI.generateBusy = true;
    renderGenerateState();
    renderImportState();
    try {
      const generated = await buildTemporaryFile();
      applyTemporaryFile(generated);
      const sizeKb = Math.max(1, Math.ceil(generated.file.size / 1024));
      setGenerateStatus(
        `临时 XLSX 生成成功：${generated.rows.length} 行数据，约 ${sizeKb} KB，尚未上传`,
        "ok"
      );
      renderImportState();
      toast("临时 XLSX 已生成");
    } catch (error) {
      UI.generatedFile = null;
      UI.generatedRows = [];
      if (error.validationErrors) {
        showError(error.validationErrors);
      } else {
        showError(`临时 XLSX 生成失败：${error.message}`);
      }
      setGenerateStatus("生成失败，请根据错误提示修改后重试", "error");
    } finally {
      UI.generateBusy = false;
      renderGenerateState();
      renderImportState();
    }
  }

  // 直接上传已生成的内存 XLSX；当前版本不判断平台已有重复零件。
  async function importBom() {
    if (UI.importBusy) return;
    const template = currentTemplate();
    const mold = selectedMold();
    if (!template || !mold || !UI.connection.token) {
      showError("请先选择目标模具并配置平台令牌");
      return;
    }

    clearErrors();
    setImportStatus("正在上传临时 XLSX 到平台...");
    UI.importBusy = true;
    renderImportState();
    try {
      const beforeCount = UI.existingPartsLoaded ? UI.existingParts.length : null;
      // 未单独生成时，自动完成校验和临时文件生成。
      if (!UI.generatedFile) {
        setImportStatus("正在生成临时 XLSX...");
        const generated = await buildTemporaryFile();
        applyTemporaryFile(generated);
      }
      setImportStatus("正在上传临时 XLSX 到平台...");
      const result = await BomApi.uploadBom(
        UI.generatedFile,
        mold.moldId,
        template.platform.excelParseConfigId
      );
      const count = Array.isArray(result && result.data) ? result.data.length : 0;
      UI.uploadCompleted = true;
      setImportStatus(`上传成功：平台返回 ${count} 个新零件 ID，正在刷新零件清单...`, "ok");
      toast("BOM 上传成功");
      await loadExistingParts(mold.moldId);
      if (UI.existingPartsLoaded) {
        const afterCount = UI.existingParts.length;
        const changeText = beforeCount === null
          ? `当前共 ${afterCount} 个零件`
          : `平台零件数 ${beforeCount} → ${afterCount}`;
        setImportStatus(
          `上传完成：接口返回 ${count} 个新零件 ID；${changeText}`,
          "ok"
        );
      } else {
        setImportStatus(
          `上传接口已成功，但零件清单刷新失败；请稍后手动刷新确认`,
          "error"
        );
      }
    } catch (error) {
      UI.uploadCompleted = false;
      if (error.validationErrors) {
        showError(error.validationErrors);
      } else {
        showError(`BOM 上传失败：${error.message}`);
      }
      setImportStatus("上传失败，请根据错误提示处理后重试", "error");
    } finally {
      UI.importBusy = false;
      renderImportState();
    }
  }

  // 只有用户点击下载时才创建本地链接，生成和上传流程不会自动下载文件。
  function downloadGeneratedFile() {
    if (!UI.generatedFile) {
      showError("请先生成临时 XLSX");
      return;
    }
    const url = URL.createObjectURL(UI.generatedFile);
    const link = document.createElement("a");
    link.href = url;
    link.download = UI.generatedFile.name || "wps-bom-import.xlsx";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // 查询模具时把按钮暂时置为忙碌，避免用户连续点击造成重复请求。
  async function searchMolds() {
    const keyword = $("mold-keyword").value.trim();
    if (!UI.connection.token) {
      $("connection-panel").open = true;
      showError("请先在开发登录中配置平台令牌");
      return;
    }

    const button = $("btn-search-mold");
    let autoSelectMoldId = "";
    UI.busy = true;
    UI.selectedMoldId = "";
    UI.existingParts = [];
    UI.existingPartsLoaded = false;
    UI.existingPartsLoading = false;
    UI.existingPartsRequestId += 1;
    button.disabled = true;
    button.textContent = "查询中";
    $("mold-select").disabled = true;
    clearErrors();
    renderImportState();
    try {
      UI.molds = await BomApi.searchMolds(keyword);
      if (UI.molds.length === 1) autoSelectMoldId = String(UI.molds[0].moldId);
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
    if (autoSelectMoldId) selectMold(autoSelectMoldId);
  }

  function init(templates, adapter) {
    UI.templates = templates;
    UI.adapter = adapter;

    $("tpl-select").addEventListener("change", (event) => {
      if (event.target.value) activateTemplate(event.target.value);
    });
    $("btn-save-token").addEventListener("click", saveToken);
    $("btn-apply-header-row").addEventListener("click", applyHeaderRow);
    $("source-header-row").addEventListener("keydown", (event) => {
      if (event.key === "Enter") applyHeaderRow();
    });
    $("btn-validate-bom").addEventListener("click", checkBom);
    $("btn-generate-xlsx").addEventListener("click", generateTemporaryFile);
    $("btn-download-xlsx").addEventListener("click", downloadGeneratedFile);
    $("btn-one-click").addEventListener("click", importBom);
    $("btn-search-mold").addEventListener("click", searchMolds);
    $("mold-keyword").addEventListener("keydown", (event) => {
      if (event.key === "Enter") searchMolds();
    });
    $("mold-select").addEventListener("change", (event) => {
      selectMold(event.target.value);
    });
    if (adapter.onDataChange) {
      adapter.onDataChange(() => loadPreviewData(UI.headerRow));
    }

    $("env-tag").textContent = adapter.isMock ? "浏览器 Mock 模式" : "WPS 环境";
    renderConnection();
    renderTemplateBar();
    loadPreviewData();
  }

  root.UI = { init };
})(typeof self !== "undefined" ? self : globalThis);
