// 任务窗格：选择模板 → 展示全部字段 → 拖动排序并实时预览第 1 行插件表头 → 上传导出 JSON。
(function (root) {
  "use strict";

  const $ = (id) => document.getElementById(id);

  const UI = {
    templates: [],
    fields: [], // 当前模板递归展开后的叶子字段
    currentTemplateId: null,
    order: [], // 叶子字段 id 的当前顺序，决定列位置
    orders: {}, // 每个模板各自保存的顺序
    adapter: null,
    drag: null, // { fromIndex, snapshot, order, lastTarget }
    result: null,
    stale: false,
    busy: false,
    lastLabelCount: 0,
    writeChain: Promise.resolve(),
  };

  function esc(value) {
    return String(value).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c]));
  }

  let toastTimer = null;
  function toast(message) {
    const el = $("toast");
    el.textContent = message;
    el.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove("show"), 2500);
  }

  function showErrors(messages) {
    const box = $("error-summary");
    box.innerHTML = "<strong>操作未完成：</strong><ul>" +
      messages.map((message) => `<li>${esc(message)}</li>`).join("") + "</ul>";
    box.classList.remove("hidden");
    box.focus();
  }

  function hideErrors() {
    $("error-summary").classList.add("hidden");
    $("error-summary").innerHTML = "";
  }

  function currentTemplate() {
    return UI.templates.find((t) => t.id === UI.currentTemplateId) || null;
  }

  function labelsFor(order) {
    return Core.headerLabels(order || UI.order, UI.fields);
  }

  function persist() {
    if (UI.currentTemplateId) UI.orders[UI.currentTemplateId] = UI.order.slice();
    Store.save({ templateId: UI.currentTemplateId || "", orders: UI.orders });
  }

  function safeFileName(text) {
    return String(text).replace(/[\\/:*?"<>|]/g, "_");
  }

  // ---------- 插件表头第 1 行 ----------

  // 所有表头写入串行执行，避免拖拽过程中快速换位时后写覆盖先写。
  function writeHeader(order) {
    const labels = labelsFor(order);
    const previousCount = UI.lastLabelCount;
    UI.lastLabelCount = labels.length;
    UI.writeChain = UI.writeChain
      .then(() => UI.adapter.writePluginHeader(labels, previousCount))
      .then(() => {
        renderSheetInfo();
        renderSheetPanel();
      })
      .catch((err) => {
        showErrors([`插件表头写入失败：${err.message}`]);
      });
    return UI.writeChain;
  }

  // ---------- 模板 ----------

  async function activateTemplate(templateId, announce) {
    const template = UI.templates.find((t) => t.id === templateId);
    if (!template) return;

    const previousCount = UI.lastLabelCount;
    UI.currentTemplateId = template.id;
    UI.fields = Core.flattenFields(template.fields);
    UI.order = Core.normalizeOrder(UI.orders[template.id], UI.fields);
    UI.lastLabelCount = UI.order.length;
    UI.result = null;
    UI.stale = false;
    persist();
    renderAll();

    const labels = labelsFor();
    try {
      UI.writeChain = UI.writeChain
        .then(async () => {
          const info = await UI.adapter.ensurePluginHeader(labels);
          const clearTo = Math.max(previousCount, (info && info.headerWidth) || 0);
          await UI.adapter.writePluginHeader(labels, clearTo);
          return info;
        })
        .then((info) => {
          renderSheetInfo();
          renderSheetPanel();
          if (announce) {
            toast(info && info.created ? "已在第 1 行创建插件表头" : "已切换模板，插件表头已更新");
          }
        })
        .catch((err) => {
          showErrors([`插件表头创建失败：${err.message}`]);
        });
      await UI.writeChain;
    } catch (_) {
      /* 错误已在上方显示 */
    }
  }

  // ---------- 渲染：模板栏 ----------

  function renderTemplateBar() {
    const select = $("tpl-select");
    const placeholder = '<option value="" disabled' +
      (UI.currentTemplateId ? "" : " selected") + ">请选择模板</option>";
    select.innerHTML = placeholder + UI.templates
      .map((t) => `<option value="${esc(t.id)}"${t.id === UI.currentTemplateId ? " selected" : ""}>${esc(t.name)}</option>`)
      .join("");
    select.value = UI.currentTemplateId || "";
    select.disabled = UI.busy || !!UI.drag;
    $("tpl-hint").textContent = UI.currentTemplateId
      ? "模板来自 data/mock.json，插件内不可编辑"
      : "选择模板后会在工作表第 1 行创建插件表头";
  }

  // ---------- 渲染：字段列表 ----------

  function renderFieldList() {
    const list = $("field-list");
    const empty = $("field-empty");
    const template = currentTemplate();

    $("field-count").textContent = template ? `${UI.fields.length} 个字段` : "";
    if (!template) {
      empty.textContent = "请选择模板";
      empty.classList.remove("hidden");
      list.innerHTML = "";
      updateControls();
      return;
    }
    if (UI.fields.length === 0) {
      empty.textContent = "模板中没有可用字段";
      empty.classList.remove("hidden");
      list.innerHTML = "";
      updateControls();
      return;
    }
    empty.classList.add("hidden");
    list.innerHTML = "";

    const map = new Map(UI.fields.map((field) => [field.id, field]));
    UI.order.forEach((id, index) => {
      const field = map.get(id);
      if (!field) return;
      const row = document.createElement("div");
      row.className = "field-row" + (UI.drag && UI.drag.fromIndex === index ? " dragging" : "");
      row.draggable = true;
      row.dataset.index = String(index);
      row.title = field.path.join(".");
      const group = field.groupLabels.length ? field.groupLabels.join(" / ") : "";
      row.innerHTML =
        `<span class="handle" aria-hidden="true">≡</span>` +
        `<div class="field-body">` +
        `  <div class="field-title"><span>${esc(field.label)}</span><span class="mono">${esc(field.path.join("."))}</span></div>` +
        (group ? `  <div class="field-path">${esc(group)}</div>` : "") +
        `</div>` +
        `<span class="field-col">${Core.colLetter(index)}</span>`;

      row.addEventListener("dragstart", (ev) => onDragStart(ev, index));
      row.addEventListener("dragover", (ev) => onDragOver(ev, index));
      row.addEventListener("drop", onDrop);
      row.addEventListener("dragend", onDragEnd);
      row.addEventListener("mouseenter", async () => {
        row.classList.add("hover");
        await UI.adapter.highlightColumn(index);
        renderSheetPanel();
      });
      row.addEventListener("mouseleave", async () => {
        row.classList.remove("hover");
        await UI.adapter.highlightColumn(null);
        renderSheetPanel();
      });

      list.appendChild(row);
    });
    updateControls();
  }

  function updateControls() {
    const hasTemplate = !!currentTemplate();
    const canUpload = hasTemplate && UI.fields.length > 0 && !UI.busy && !UI.drag;
    $("btn-upload").disabled = !canUpload;
    $("tpl-select").disabled = UI.busy || !!UI.drag;
    $("upload-hint").textContent = hasTemplate
      ? "拖动字段调整列顺序；上传会按当前顺序导出原始表头下方（第 3 行起）的数据"
      : "选择模板后即可拖动排序并上传";
  }

  // ---------- 拖动排序：每次位置变化写一次表头预览 ----------

  function onDragStart(ev, index) {
    UI.drag = {
      fromIndex: index,
      snapshot: UI.order.slice(),
      order: UI.order.slice(),
      lastTarget: index,
    };
    ev.dataTransfer.effectAllowed = "move";
    ev.dataTransfer.setData("text/plain", String(index));
    ev.currentTarget.classList.add("dragging");
    updateControls();
  }

  function dragInsertIndex(overIndex, ev) {
    const rect = ev.currentTarget.getBoundingClientRect();
    return ev.clientY > rect.top + rect.height / 2 ? overIndex + 1 : overIndex;
  }

  function onDragOver(ev, overIndex) {
    if (!UI.drag) return;
    ev.preventDefault();
    ev.dataTransfer.dropEffect = "move";
    const to = dragInsertIndex(overIndex, ev);
    const from = UI.drag.fromIndex;
    const target = to > from ? to - 1 : to;
    if (target === UI.drag.lastTarget) return;

    UI.drag.lastTarget = target;
    UI.drag.order = Core.moveItem(UI.drag.snapshot, from, target);
    markInsertion(overIndex, ev);
    writeHeader(UI.drag.order); // 位置一变化就预览，不等待松手
  }

  function markInsertion(overIndex, ev) {
    document.querySelectorAll(".field-row").forEach((row) => {
      row.classList.remove("insert-before", "insert-after");
    });
    const rect = ev.currentTarget.getBoundingClientRect();
    ev.currentTarget.classList.add(ev.clientY > rect.top + rect.height / 2 ? "insert-after" : "insert-before");
  }

  function onDrop(ev) {
    if (!UI.drag) return;
    ev.preventDefault();
    UI.order = UI.drag.order.slice();
    endDrag();
    persist();
    renderFieldList();
    writeHeader(UI.order);
    markResultStale();
  }

  function cancelDrag() {
    if (!UI.drag) return;
    const snapshot = UI.drag.snapshot;
    endDrag();
    UI.order = snapshot;
    renderFieldList();
    writeHeader(snapshot);
    toast("已取消拖拽，恢复原顺序");
  }

  function endDrag() {
    UI.drag = null;
    document.querySelectorAll(".field-row").forEach((row) => {
      row.classList.remove("dragging", "insert-before", "insert-after");
    });
    updateControls();
  }

  function onDragEnd() {
    if (UI.drag) cancelDrag();
  }

  function markResultStale() {
    if (!UI.result) return;
    UI.stale = true;
    renderResult();
  }

  // ---------- 渲染：模拟工作表预览 ----------

  function renderSheetInfo() {
    const el = $("sheet-info");
    if (!el) return;
    if (!UI.adapter || typeof UI.adapter.getSheetInfo !== "function") {
      el.textContent = "当前环境不支持读取工作表信息";
      return;
    }

    let info;
    try {
      info = UI.adapter.getSheetInfo();
    } catch (err) {
      el.textContent = `读取工作表信息失败：${err.message}`;
      return;
    }
    if (!info) {
      el.textContent = "无法读取当前工作表";
      return;
    }

    const range = info.rowCount > 0 && info.columnCount > 0
      ? `${Core.colLetter(info.startColumn - 1)}${info.startRow}:` +
        `${Core.colLetter(info.startColumn + info.columnCount - 2)}` +
        `${info.startRow + info.rowCount - 1}`
      : "空工作表";
    const size = info.rowCount > 0 && info.columnCount > 0
      ? `${info.rowCount} 行 × ${info.columnCount} 列`
      : "无数据";

    el.innerHTML =
      `<span class="label">Sheet</span><span class="value">${esc(info.name)}</span>` +
      `<span class="label">使用范围</span><span class="value">${esc(range)}</span>` +
      `<span class="label">区域大小</span><span class="value">${size}</span>`;
  }

  function renderSheetPanel() {
    const panel = $("sheet-panel");
    const view = UI.adapter.getSheetView ? UI.adapter.getSheetView() : null;
    if (!view) {
      panel.innerHTML = currentTemplate()
        ? '<div class="dim">插件表头已写入当前工作表第 1 行。</div>'
        : '<div class="dim">选择模板后会写入工作表第 1 行。</div>';
      return;
    }
    const headers = view.headers || [];
    const colCount = Math.max(
      view.pluginLabels.length,
      headers.length,
      ...view.rows.map((row) => row.length),
      0
    );
    const hasPlugin = view.hasPluginHeader;
    let html = `<div class="sheet-name">模拟工作表：${esc(view.name)}` +
      (hasPlugin ? "（第 1 行插件表头 / 第 2 行原始表头）" : "（第 1 行原始表头）") +
      "</div>";
    html += '<table class="sheet"><tr><th></th>';
    for (let c = 0; c < colCount; c++) {
      html += `<th class="col-letter">${Core.colLetter(c)}</th>`;
    }
    html += "</tr>";

    let rowNo = 1;
    if (hasPlugin) {
      html += `<tr><td class="row-no">${rowNo++}</td>`;
      for (let c = 0; c < colCount; c++) {
        const label = view.pluginLabels[c] || "";
        const cls = label ? "" : " empty";
        const hover = c === view.highlightCol ? " hover" : "";
        html += `<td class="plugin-header${cls}${hover}">${esc(label)}</td>`;
      }
      html += "</tr>";
    }

    html += `<tr><td class="row-no">${rowNo++}</td>`;
    for (let c = 0; c < colCount; c++) {
      const value = headers[c];
      const hover = c === view.highlightCol ? " hover" : "";
      html += `<td class="src-header${hover}">${value === undefined || value === null ? "" : esc(value)}</td>`;
    }
    html += "</tr>";

    view.rows.forEach((row) => {
      html += `<tr><td class="row-no">${rowNo++}</td>`;
      for (let c = 0; c < colCount; c++) {
        const value = row[c];
        const hover = c === view.highlightCol ? " hover" : "";
        html += `<td class="${hover}">${value === null || value === undefined ? "" : esc(value)}</td>`;
      }
      html += "</tr>";
    });
    html += "</table>";
    panel.innerHTML = html;
  }

  // ---------- 上传导出 ----------

  async function upload() {
    hideErrors();
    const template = currentTemplate();
    if (!template || UI.fields.length === 0) return;

    UI.busy = true;
    updateControls();
    try {
      const leaves = Core.orderedLeaves(UI.order, UI.fields);
      const labels = labelsFor();
      await UI.writeChain;
      const info = await UI.adapter.ensurePluginHeader(labels);
      const clearTo = Math.max(UI.lastLabelCount, (info && info.headerWidth) || 0);
      await UI.adapter.writePluginHeader(labels, clearTo);
      UI.lastLabelCount = labels.length;

      const rows = await UI.adapter.readRows(leaves.length);
      const json = Core.buildExport(template, UI.adapter.getSheetName(), UI.order, UI.fields, rows);
      UI.result = json;
      UI.stale = false;
      renderResult();

      const filename = `${safeFileName(UI.adapter.getSheetName())}-${safeFileName(template.name)}.json`;
      await UI.adapter.uploadJson(json, filename);
      toast(`已导出 ${rows.length} 行数据：${filename}`);
      renderSheetInfo();
      renderSheetPanel();
    } catch (err) {
      showErrors([`上传失败：${err.message}`]);
    } finally {
      UI.busy = false;
      updateControls();
    }
  }

  function renderResult() {
    const section = $("result-section");
    if (!UI.result) {
      section.classList.add("hidden");
      return;
    }
    section.classList.remove("hidden");
    const json = UI.result;
    const stale = UI.stale ? " · 字段顺序已变化，请重新上传" : "";
    $("res-stats").textContent = `${json.sheet} · ${json.fields.length} 个字段 · ${json.records.length} 行${stale}`;
    $("res-json").textContent = JSON.stringify(json, null, 2);
  }

  async function copyResult() {
    if (!UI.result) return;
    const text = JSON.stringify(UI.result, null, 2);
    try {
      await navigator.clipboard.writeText(text);
    } catch (_) {
      const area = document.createElement("textarea");
      area.value = text;
      document.body.appendChild(area);
      area.select();
      document.execCommand("copy");
      area.remove();
    }
    toast("已复制 JSON");
  }

  async function exportResult() {
    if (!UI.result) return;
    const template = currentTemplate();
    const filename = `${safeFileName(UI.adapter.getSheetName())}-${safeFileName(template ? template.name : "mapping")}.json`;
    await UI.adapter.uploadJson(UI.result, filename);
    toast(`已导出：${filename}`);
  }

  // ---------- 初始化 ----------

  function renderAll() {
    renderSheetInfo();
    renderTemplateBar();
    renderFieldList();
    renderSheetPanel();
    updateControls();
    renderResult();
  }

  function init(templates, adapter) {
    UI.templates = templates;
    UI.adapter = adapter;
    const saved = Store.load();
    UI.orders = saved.orders || {};

    $("tpl-select").addEventListener("change", (ev) => {
      if (ev.target.value) activateTemplate(ev.target.value, true);
    });
    $("btn-upload").addEventListener("click", upload);
    $("btn-copy").addEventListener("click", copyResult);
    $("btn-export").addEventListener("click", exportResult);

    document.addEventListener("keydown", (ev) => {
      if (ev.key === "Escape" && UI.drag) cancelDrag();
    });
    window.addEventListener("blur", () => {
      if (UI.drag) cancelDrag();
    });

    if (adapter.onDataChange) {
      adapter.onDataChange(() => {
        renderSheetInfo();
        if (UI.result) {
          UI.stale = true;
          renderResult();
        }
      });
    }

    $("env-tag").textContent = adapter.isMock ? "浏览器 mock 模式" : "WPS 环境";
    renderAll();

    const initial = templates.some((t) => t.id === saved.templateId) ? saved.templateId : "";
    if (initial) activateTemplate(initial, false);
  }

  root.UI = { init };
})(typeof self !== "undefined" ? self : globalThis);
