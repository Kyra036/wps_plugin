// 任务窗格：读取当前 BOM 工作表，匹配字段，并选择本次导入的目标模具。
(function (root) {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const PREVIEW_ROW_LIMIT = 20;

  const UI = {
    templates: [],
    fields: [],
    mappingByField: {},
    currentTemplateId: "",
    adapter: null,
    headerInfo: null,
    headerRow: null,
    previewRows: [],
    totalDataRows: 0,
    previewRequestId: 0,
    molds: [],
    selectedMoldId: "",
    selectedMoldData: null,
    moldSearching: false,
    moldSearchError: "",
    moldSearchTimer: null,
    moldSearchRequestId: 0,
    moldDropdownOpen: false,
    moldActiveIndex: -1,
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

  // 转义写入 innerHTML 的文本，避免表格内容被当作 HTML 执行。
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
  // 显示短暂的轻提示，不阻塞当前操作。
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

  // 清空顶部错误摘要。
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

  // 显示临时 XLSX 生成状态，type 可为空、ok 或 error。
  function setGenerateStatus(message, type) {
    const element = $("generate-status");
    element.textContent = message || "";
    element.className = message ? (type || "") : "hidden";
  }

  // 显示真实上传状态和上传后的逐行核对结果。
  function setImportStatus(message, type) {
    const element = $("import-status");
    element.textContent = message || "";
    element.className = message ? (type || "") : "hidden";
  }

  // 清空内存中的临时文件、预览和上传状态，源数据变化时调用。
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

  // 返回当前选中的本地 BOM 模板。
  function currentTemplate() {
    return UI.templates.find((template) => template.id === UI.currentTemplateId) || null;
  }

  // 返回当前选中的目标模具；优先使用完整对象，避免搜索结果刷新后丢失。
  function selectedMold() {
    if (UI.selectedMoldData) return UI.selectedMoldData;
    return UI.molds.find((mold) => String(mold.moldId) === String(UI.selectedMoldId)) || null;
  }

  // 当前映射完全以 UI.mappingByField 为准，自动匹配只负责初始化它。
  function currentFieldMapping() {
    return Core.resolveFieldMapping(
      UI.headerInfo ? UI.headerInfo.headers : [],
      UI.fields,
      UI.mappingByField
    );
  }

  // 根据表头名称生成初始字段映射，写入 UI.mappingByField。
  function applyAutoMapping() {
    const result = Core.autoMatchHeaders(
      UI.headerInfo ? UI.headerInfo.headers : [],
      UI.fields
    );
    const nextMapping = {};
    result.matches.forEach((match) => {
      if (match.fieldId && Number.isInteger(match.headerIndex)) {
        nextMapping[match.fieldId] = match.headerIndex;
      }
    });
    UI.mappingByField = nextMapping;
  }

  // 手动选择时保持一对一，若该列已分配给其它字段则自动移除旧选择。
  function updateManualMapping(fieldId, value) {
    clearErrors();
    setValidationStatus("");
    resetGeneratedFile();
    if (value === "") {
      delete UI.mappingByField[fieldId];
    } else {
      const sourceIndex = Number(value);
      Object.keys(UI.mappingByField).forEach((otherFieldId) => {
        if (
          otherFieldId !== fieldId &&
          Number(UI.mappingByField[otherFieldId]) === sourceIndex
        ) {
          delete UI.mappingByField[otherFieldId];
        }
      });
      UI.mappingByField[fieldId] = sourceIndex;
    }
    renderMapping();
    renderValidationState();
    renderGenerateState();
    renderImportState();
  }

  // 用户点击自动匹配时，重置手动选择并重新执行表头匹配。
  function useAutomaticMapping() {
    if (UI.busy || UI.validationBusy || UI.generateBusy || UI.importBusy) return;
    clearErrors();
    setValidationStatus("");
    resetGeneratedFile();
    applyAutoMapping();
    renderMapping();
    renderValidationState();
    renderGenerateState();
    renderImportState();
    toast("已重新按表头名称自动匹配");
  }

  // 将日期和普通单元格值转换为适合页面展示的字符串。
  function formatValue(value) {
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? "" : value.toISOString().slice(0, 10);
    }
    return value === null || value === undefined ? "" : String(value);
  }

  // 显示当前活动 Sheet 名称，帮助用户确认没有读错工作表。
  function renderSheetInfo() {
    const element = $("sheet-info");
    try {
      const info = UI.adapter.getSheetInfo();
      if (!info || !info.rowCount || !info.columnCount) {
        element.textContent = "当前工作表为空";
        return;
      }
      element.innerHTML =
        `<span class="label">Sheet</span><span class="value">${esc(info.name)}</span>`;
    } catch (error) {
      element.textContent = `读取工作表信息失败：${error.message}`;
    }
  }

  // 同步表头行输入框，避免用户正在输入时被渲染覆盖。
  function renderHeaderRowControl() {
    const input = $("source-header-row");
    if (document.activeElement !== input && UI.headerRow) {
      input.value = UI.headerRow;
    }
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
      ? "自动匹配后，可在字段映射中手动选择 Excel 列"
      : "选择模板后开始自动匹配";
  }

  // 左边固定平台字段，右边手动选择 Excel 列；自动匹配只是初始选择。
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

    const headers = UI.headerInfo ? UI.headerInfo.headers : [];
    const result = currentFieldMapping();
    section.classList.remove("hidden");
    $("btn-auto-map").disabled =
      UI.busy || UI.validationBusy || UI.generateBusy || UI.importBusy;
    stats.textContent = `${result.assignedCount}/${UI.fields.length} 个字段已指定`;

    result.fieldMappings.forEach((mapping) => {
      const row = document.createElement("div");
      const target = document.createElement("span");
      target.className = "mapping-target";
      if (mapping.field) {
        const group = document.createElement("span");
        group.className = "group";
        group.textContent = mapping.field.groupLabels.length
          ? `${mapping.field.groupLabels.join(" / ")} / `
          : "";
        const requiredText = mapping.field.required ? "" : "（可选）";
        target.append(group, document.createTextNode(`${mapping.field.label}${requiredText}`));
      } else {
        target.textContent = "未命名字段";
      }

      const select = document.createElement("select");
      select.className = "mapping-select";
      select.dataset.fieldId = mapping.fieldId;
      select.disabled =
        UI.busy || UI.validationBusy || UI.generateBusy || UI.importBusy;
      const emptyOption = document.createElement("option");
      emptyOption.value = "";
      emptyOption.textContent = mapping.field.required
        ? "请选择 Excel 列"
        : "不选择，字段留空";
      select.appendChild(emptyOption);

      headers.forEach((header, headerIndex) => {
        const option = document.createElement("option");
        const column = header.column || Core.colLetter(headerIndex);
        const label = header.fullLabel || header.label || "（空表头）";
        option.value = String(headerIndex);
        option.textContent = `${column}  ${label}`;
        select.appendChild(option);
      });

      select.value = mapping.headerIndex === null ? "" : String(mapping.headerIndex);
      select.addEventListener("change", (event) => {
        updateManualMapping(mapping.fieldId, event.target.value);
      });

      const duplicate = mapping.headerIndex !== null &&
        result.duplicateColumns.includes(mapping.headerIndex);
      if (mapping.headerIndex === null || duplicate) row.className = "mapping-row unmatched";
      else row.className = "mapping-row";

      row.append(target, select);
      list.appendChild(row);
    });

    const messages = [];
    if (result.duplicateColumns.length > 0) {
      messages.push("同一 Excel 列不能指定给多个字段");
    }
    if (result.missingFields.length > 0) {
      messages.push(`还有 ${result.missingFields.length} 个必填字段未选择 Excel 列`);
    }
    if (result.valid && result.unmappedFields.length > 0) {
      const optionalNames = result.unmappedFields
        .map((item) => item.field.label)
        .join("、");
      messages.push(`可选项未指定：${optionalNames}，将留空`);
      if (result.unmappedFields.some((item) => item.field.key === "partNum")) {
        messages.push("数量未指定时按 1 导入");
      }
    }
    hint.textContent = messages.length
      ? messages.join("；")
      : "必填和可选字段均已指定，可手动调整右侧列";
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

  // 组合模具编号和名称，作为选中后输入框的展示文本。
  function moldDisplayName(mold) {
    return [mold && mold.moldNo, mold && mold.moldName]
      .filter((value) => value !== undefined && value !== null && String(value).trim())
      .join(" / ");
  }

  // 按零件编号统计平台已有数量，用于上传后识别被后端跳过的行。
  function countPartNumbers(parts) {
    const counts = new Map();
    (parts || []).forEach((part) => {
      const partNo = part && part.partNo !== undefined && part.partNo !== null
        ? String(part.partNo).trim()
        : "";
      if (partNo) counts.set(partNo, (counts.get(partNo) || 0) + 1);
    });
    return counts;
  }

  // 兼容平台模板 importSchema 是对象或 JSON 字符串两种返回形式。
  function parseImportSchema(rawSchema) {
    if (!rawSchema) return null;
    if (typeof rawSchema === "string") {
      try {
        return JSON.parse(rawSchema);
      } catch (_) {
        return null;
      }
    }
    return rawSchema;
  }

  // 固定结束行早于当前文件最后一行时，后部数据会被平台解析器忽略。
  async function validateTemplateDataRange(template, lastDataRow) {
    let payload;
    try {
      payload = await BomApi.getParseConfig(template.platform.excelParseConfigId);
    } catch (error) {
      return { checked: false, message: error.message };
    }
    const schema = parseImportSchema(payload && payload.data && payload.data.importSchema);
    const sheets = schema && Array.isArray(schema.sheets) ? schema.sheets : [];
    const fixedEnds = [];

    sheets.forEach((sheet) => {
      const end = sheet && sheet.data && sheet.data.end;
      const strategy = String(end && end.strategy || "").toLowerCase();
      if (strategy === "fixed" && Number.isInteger(Number(end.row))) {
        fixedEnds.push(Number(end.row));
      }
    });

    const blockingEnd = fixedEnds.find((row) => row < lastDataRow);
    if (blockingEnd !== undefined) {
      throw new Error(
        `平台导入模板的数据结束行固定为第 ${blockingEnd} 行，` +
        `但当前临时文件数据到第 ${lastDataRow} 行。` +
        "请把模板数据结束策略改为 lastRow，或缩短本次导入范围。"
      );
    }
    return { checked: true };
  }

  // 打开模具候选下拉并刷新结果状态。
  function openMoldDropdown() {
    UI.moldDropdownOpen = true;
    renderMoldOptions();
  }

  // 关闭模具候选下拉并清除键盘高亮位置。
  function closeMoldDropdown() {
    UI.moldDropdownOpen = false;
    UI.moldActiveIndex = -1;
    renderMoldOptions();
  }

  // 组合框结果区只展示搜索状态和候选模具，不修改目标模具本身。
  function renderMoldOptions() {
    const input = $("mold-keyword");
    const toggle = $("btn-toggle-molds");
    const results = $("mold-results");
    results.innerHTML = "";
    input.setAttribute("aria-expanded", String(UI.moldDropdownOpen));
    toggle.disabled = UI.moldSearching;

    if (!UI.moldDropdownOpen) {
      results.classList.add("hidden");
      return;
    }
    results.classList.remove("hidden");

    if (UI.moldSearching) {
      const state = document.createElement("div");
      state.className = "mold-result-state";
      state.textContent = "正在查询模具...";
      results.appendChild(state);
      return;
    }
    if (UI.moldSearchError) {
      const state = document.createElement("div");
      state.className = "mold-result-state";
      state.textContent = UI.moldSearchError;
      results.appendChild(state);
      return;
    }
    if (UI.molds.length === 0) {
      const state = document.createElement("div");
      state.className = "mold-result-state";
      state.textContent = "输入编号或名称搜索模具";
      results.appendChild(state);
      return;
    }

    UI.molds.forEach((mold, index) => {
      const item = document.createElement("div");
      item.className = "mold-result-item" + (index === UI.moldActiveIndex ? " active" : "");
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", String(index === UI.moldActiveIndex));
      const no = document.createElement("span");
      no.className = "mold-result-no";
      no.textContent = mold.moldNo || "未编号模具";
      const name = document.createElement("span");
      name.className = "mold-result-name";
      name.textContent = mold.moldName || "未命名模具";
      item.append(no, name);
      item.addEventListener("mousedown", (event) => event.preventDefault());
      item.addEventListener("click", () => selectMold(mold));
      results.appendChild(item);
    });
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

  // 保存开发令牌到本地，并刷新依赖连接状态的按钮。
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
  function selectMold(mold) {
    const selected = typeof mold === "object" && mold
      ? mold
      : UI.molds.find((item) => String(item.moldId) === String(mold));
    if (!selected) return;
    UI.selectedMoldId = String(selected.moldId);
    UI.selectedMoldData = selected;
    $("mold-keyword").value = moldDisplayName(selected);
    UI.moldDropdownOpen = false;
    UI.moldActiveIndex = -1;
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
    const mapping = template ? currentFieldMapping() : null;
    const missingFields = mapping ? mapping.missingFields.length : UI.fields.length;
    const mold = selectedMold();

    let message = "";
    if (!template) message = "请先选择 BOM 模板";
    else if (missingFields > 0) message = `仍有 ${missingFields} 个字段未选择 Excel 列`;
    else if (mapping.duplicateColumns.length > 0) message = "同一 Excel 列不能指定给多个字段";
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
    const mapping = template ? currentFieldMapping() : null;
    const ready = !!template &&
      mapping.valid &&
      UI.totalDataRows > 0;
    button.disabled = !ready || UI.validationBusy;
    button.textContent = UI.validationBusy ? "正在检查..." : "检查完整 BOM";
  }

  // 临时文件只依赖当前模板和源表数据，不要求令牌或目标模具。
  function renderGenerateState() {
    const button = $("btn-generate-xlsx");
    const template = currentTemplate();
    const mapping = template ? currentFieldMapping() : null;
    const ready = !!template &&
      mapping.valid &&
      UI.totalDataRows > 0;
    button.disabled = !ready || UI.generateBusy || UI.importBusy;
    button.textContent = UI.generateBusy
      ? "正在生成..."
      : "生成临时 XLSX（暂不上传）";
  }

  // 统一刷新工作表和全部业务区域的展示状态。
  function renderAll() {
    renderSheetInfo();
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
      const dataRange = columnCount > 0
        ? UI.adapter.getDataRange(headerInfo, columnCount)
        : { dataRowCount: 0, lastDataRow: headerInfo.dataStartRow - 1 };
      const rows = columnCount > 0
        ? await UI.adapter.readRows(columnCount, PREVIEW_ROW_LIMIT, headerRow)
        : [];
      if (requestId !== UI.previewRequestId) return;

      UI.headerInfo = headerInfo;
      UI.headerRow = headerInfo.headerRow || headerRow;
      UI.previewRows = rows;
      UI.totalDataRows = dataRange.dataRowCount;
      applyAutoMapping();
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

  // 加载模板字段，并按当前表头自动初始化列映射。
  function activateTemplate(templateId) {
    const template = UI.templates.find((item) => item.id === templateId);
    if (!template) return;
    UI.currentTemplateId = template.id;
    UI.fields = Core.flattenTemplate(template);
    applyAutoMapping();
    clearErrors();
    setValidationStatus("");
    resetGeneratedFile();
    renderAll();
    toast("已加载 BOM 模板");
  }

  // 读取全部源行并完成校验；该函数不调用任何平台写接口。
  async function readAndValidateBom() {
    const mapping = currentFieldMapping();
    if (mapping.missingFields.length > 0) {
      throw new Error(`仍有 ${mapping.missingFields.length} 个字段未选择 Excel 列`);
    }
    if (mapping.duplicateColumns.length > 0) {
      throw new Error("同一 Excel 列不能指定给多个字段");
    }
    const sourceRows = await UI.adapter.readRows(
      UI.headerInfo.headers.length,
      Number.MAX_SAFE_INTEGER,
      UI.headerRow
    );
    const mappedRows = Core.buildMappedRows(mapping.fieldMappings, sourceRows);
    const validation = Core.validateImportRows(mappedRows, UI.headerInfo.dataStartRow);
    return { sourceRows, mappedRows, validation };
  }

  // 执行只读 BOM 全量校验，并在页面显示有效行和错误行。
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
        setValidationStatus(`检查通过：${result.validation.rows.length} 行通过前端校验`, "ok");
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

  // 保存生成结果到内存，并刷新临时 XLSX 预览。
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

  // 一键导入入口：按需生成临时 XLSX、检查模板范围、上传并核对结果。
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
      if (!UI.existingPartsLoaded) {
        setImportStatus("正在读取上传前的平台零件清单...");
        await loadExistingParts(mold.moldId);
        if (!UI.existingPartsLoaded) {
          throw new Error("无法读取上传前的平台零件清单");
        }
      }
      const beforePartNos = countPartNumbers(UI.existingParts);
      const beforeCount = UI.existingParts.length;
      // 未单独生成时，自动完成校验和临时文件生成。
      if (!UI.generatedFile) {
        setImportStatus("正在生成临时 XLSX...");
        const generated = await buildTemporaryFile();
        applyTemporaryFile(generated);
      }
      setImportStatus("正在检查平台导入模板的数据范围...");
      const lastDataRow = Number(template.platform.headerRow || 5) + UI.generatedRows.length;
      const rangeCheck = await validateTemplateDataRange(template, lastDataRow);
      if (!rangeCheck.checked) {
        setImportStatus(`模板范围检查跳过：${rangeCheck.message}；继续上传...`);
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
        const afterPartNos = countPartNumbers(UI.existingParts);
        const submittedPartNos = UI.generatedRows
          .map((row) => String(row.partNo || "").trim())
          .filter(Boolean);
        const notAdded = submittedPartNos.filter(
          (partNo) => (afterPartNos.get(partNo) || 0) <= (beforePartNos.get(partNo) || 0)
        );
        const existingDuplicates = notAdded.filter(
          (partNo) => (beforePartNos.get(partNo) || 0) > 0
        );
        const backendSkipped = notAdded.filter(
          (partNo) => (beforePartNos.get(partNo) || 0) === 0
        );
        const changeText = beforeCount === null
          ? `当前共 ${afterCount} 个零件`
          : `平台零件数 ${beforeCount} → ${afterCount}`;
        const lines = [`上传完成：接口返回 ${count} 个新零件 ID；${changeText}`];
        if (existingDuplicates.length > 0) {
          lines.push(`未新增（平台已有同编号）：${existingDuplicates.join("、")}`);
        }
        if (backendSkipped.length > 0) {
          lines.push(`未入库（后端未返回 ID）：${backendSkipped.join("、")}`);
        }
        setImportStatus(lines.join("\n"), notAdded.length > 0 ? "error" : "ok");
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

  // 用户修改搜索内容时立即取消旧选择，但不打断后续输入。
  function clearSelectedMoldForTyping() {
    if (!UI.selectedMoldId && !UI.selectedMoldData) return;
    UI.selectedMoldId = "";
    UI.selectedMoldData = null;
    UI.existingParts = [];
    UI.existingPartsLoaded = false;
    UI.existingPartsLoading = false;
    UI.existingPartsRequestId += 1;
    UI.uploadCompleted = false;
    setImportStatus("");
    renderMoldPartStatus();
    renderExistingPartsPreview();
    renderImportState();
  }

  // 对模具搜索做 350ms 防抖，避免每次按键都请求平台。
  function scheduleMoldSearch() {
    clearTimeout(UI.moldSearchTimer);
    UI.moldSearchTimer = setTimeout(() => searchMolds(), 350);
  }

  // 远程搜索组合框：输入停止 350ms 后查询，空输入返回平台第一页模具。
  async function searchMolds() {
    clearTimeout(UI.moldSearchTimer);
    const input = $("mold-keyword");
    const keyword = input.value.trim();
    UI.moldDropdownOpen = true;
    UI.moldSearchError = "";
    UI.moldActiveIndex = -1;

    if (!UI.connection.token) {
      UI.molds = [];
      UI.moldSearching = false;
      UI.moldSearchError = "请先在开发登录中配置平台令牌";
      renderMoldOptions();
      return;
    }

    const requestId = ++UI.moldSearchRequestId;
    UI.moldSearching = true;
    renderMoldOptions();
    try {
      const molds = await BomApi.searchMolds(keyword);
      if (requestId !== UI.moldSearchRequestId) return;
      UI.molds = molds;
      UI.moldActiveIndex = molds.length > 0 ? 0 : -1;
      UI.moldSearchError = molds.length > 0
        ? ""
        : "没有找到匹配的模具";
    } catch (error) {
      if (requestId !== UI.moldSearchRequestId) return;
      UI.molds = [];
      UI.moldSearchError = error.message;
    } finally {
      if (requestId === UI.moldSearchRequestId) {
        UI.moldSearching = false;
        renderMoldOptions();
      }
    }
  }

  // 初始化页面模板、数据适配器、事件监听和首屏工作表预览。
  function init(templates, adapter) {
    UI.templates = templates;
    UI.adapter = adapter;

    $("btn-reload-page").addEventListener("click", () => window.location.reload());
    $("tpl-select").addEventListener("change", (event) => {
      if (event.target.value) activateTemplate(event.target.value);
    });
    $("btn-auto-map").addEventListener("click", useAutomaticMapping);
    $("btn-save-token").addEventListener("click", saveToken);
    $("btn-apply-header-row").addEventListener("click", applyHeaderRow);
    $("source-header-row").addEventListener("keydown", (event) => {
      if (event.key === "Enter") applyHeaderRow();
    });
    $("btn-validate-bom").addEventListener("click", checkBom);
    $("btn-generate-xlsx").addEventListener("click", generateTemporaryFile);
    $("btn-download-xlsx").addEventListener("click", downloadGeneratedFile);
    $("btn-one-click").addEventListener("click", importBom);
    $("btn-toggle-molds").addEventListener("click", () => {
      if (UI.moldDropdownOpen) {
        closeMoldDropdown();
      } else {
        openMoldDropdown();
        if (!UI.moldSearching && UI.molds.length === 0) searchMolds();
      }
    });
    $("mold-keyword").addEventListener("focus", () => {
      openMoldDropdown();
      if (UI.selectedMoldData && UI.molds.length === 0) {
        UI.molds = [UI.selectedMoldData];
        UI.moldActiveIndex = 0;
        renderMoldOptions();
      } else if (UI.molds.length === 0 && !UI.moldSearching) {
        searchMolds();
      }
    });
    $("mold-keyword").addEventListener("input", () => {
      clearSelectedMoldForTyping();
      UI.moldSearchRequestId += 1;
      UI.moldSearching = false;
      UI.molds = [];
      UI.moldDropdownOpen = true;
      UI.moldActiveIndex = -1;
      UI.moldSearchError = "";
      renderMoldOptions();
      scheduleMoldSearch();
    });
    $("mold-keyword").addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        closeMoldDropdown();
        return;
      }
      if (event.key === "ArrowDown" || event.key === "ArrowUp") {
        event.preventDefault();
        if (UI.moldSearching) return;
        if (!UI.moldDropdownOpen) openMoldDropdown();
        if (UI.molds.length === 0) {
          searchMolds();
          return;
        }
        const direction = event.key === "ArrowDown" ? 1 : -1;
        UI.moldActiveIndex = (UI.moldActiveIndex + direction + UI.molds.length) % UI.molds.length;
        renderMoldOptions();
        return;
      }
      if (event.key === "Enter" && UI.moldDropdownOpen) {
        event.preventDefault();
        if (UI.moldSearching) return;
        if (UI.molds.length > 0) {
          const index = UI.moldActiveIndex >= 0 ? UI.moldActiveIndex : 0;
          selectMold(UI.molds[index]);
        } else {
          searchMolds();
        }
      }
    });
    document.addEventListener("mousedown", (event) => {
      if (!$("mold-combobox").contains(event.target)) closeMoldDropdown();
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
