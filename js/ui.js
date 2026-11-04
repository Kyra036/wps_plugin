// 任务窗格：只读当前工作表 → 展示源表头与字段 → 映射排序 → 上传导出 JSON。
(function (root) {
  "use strict";

  const $ = (id) => document.getElementById(id);
  // 预览限制只影响页面渲染，上传时仍读取全部数据行。
  const PREVIEW_ROW_LIMIT = 20;

  const UI = {
    templates: [],
    customTemplates: [], // 用户保存的模板，与内置模板分开维护
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
    editingFieldId: null,
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
    return Store.save({ orders: UI.orders, customTemplates: UI.customTemplates });
  }

  function safeFileName(text) {
    return String(text).replace(/[\\/:*?"<>|]/g, "_");
  }

  // ---------- 字段顺序预览 ----------

  function writeHeader(order) {
    // 拖拽字段后只刷新本地预览；客户源工作表始终保持只读。
    UI.lastLabelCount = labelsFor(order).length;
    renderSheetInfo();
    renderSourceHeaders();
    renderSheetPanel();
    return Promise.resolve();
  }

  // ---------- 模板 ----------

  // 模板选中后展开分组字段，并恢复该模板上次保存的字段顺序。
  async function activateTemplate(templateId, announce) {
    const template = UI.templates.find((t) => t.id === templateId);
    if (!template) return;

    UI.currentTemplateId = template.id;
    UI.fields = Core.flattenTemplate(template);
    UI.order = Core.normalizeOrder(UI.orders[template.id], UI.fields);
    UI.lastLabelCount = UI.order.length;
    UI.result = null;
    UI.stale = false;
    persist();
    renderAll();

    renderSheetInfo();
    renderSourceHeaders();
    renderSheetPanel();
    if (announce) {
      toast("已切换模板");
    }
  }

  // ---------- 渲染：模板栏 ----------

  function renderTemplateBar() {
    const select = $("tpl-select");
    // 按 builtIn 将模板分成系统模板和用户模板，后续“我的模板”直接进入第二组。
    const builtInTemplates = UI.templates.filter((template) => template.builtIn === true);
    const customTemplates = UI.templates.filter((template) => template.builtIn !== true);
    const templateOption = (template) =>
      `<option value="${esc(template.id)}"${template.id === UI.currentTemplateId ? " selected" : ""}>` +
      `${esc(template.name)}</option>`;
    const placeholder = '<option value="" disabled' +
      (UI.currentTemplateId ? "" : " selected") + ">请选择模板</option>";
    const customOptions = customTemplates.length
      ? customTemplates.map(templateOption).join("")
      : '<option value="" disabled>暂无自定义模板</option>';
    select.innerHTML = placeholder +
      `<optgroup label="内置模板">${builtInTemplates.map(templateOption).join("")}</optgroup>` +
      `<optgroup label="我的模板">${customOptions}</optgroup>`;
    select.value = UI.currentTemplateId || "";
    select.disabled = UI.busy || !!UI.drag;
    $("template-library-meta").textContent =
      `模板库：内置 ${builtInTemplates.length} 个 · 我的模板 ${customTemplates.length} 个`;
    const activeTemplate = currentTemplate();
    $("template-save").classList.toggle("hidden", !activeTemplate);
    $("btn-save-template").disabled = !activeTemplate || UI.busy || !!UI.drag;
    $("tpl-hint").textContent = activeTemplate
      ? `${activeTemplate.builtIn === true ? "内置模板" : "我的模板"}已加载，源工作表不会被修改`
      : "选择 BOM 模板后开始匹配源表头";
    const clearButton = $("btn-clear-template");
    clearButton.classList.toggle("hidden", !UI.currentTemplateId);
    clearButton.disabled = UI.busy || !!UI.drag;
    const applyButton = $("btn-apply-template");
    applyButton.classList.toggle("hidden", !UI.currentTemplateId);
    applyButton.disabled = !UI.currentTemplateId || UI.busy || !!UI.drag;
    renderCurrentTemplate();
    renderTemplateGroups();
    renderTemplateConfig();
  }

  // 清除当前选择只影响任务窗格状态，不修改本地字段顺序或模板定义。
  function clearCurrentTemplate() {
    UI.currentTemplateId = null;
    UI.fields = [];
    UI.order = [];
    UI.result = null;
    UI.stale = false;
    renderAll();
    toast("已清除当前模板");
  }

  // 使用 JSON 深拷贝隔离运行时修改，避免“我的模板”和原模板共享 groups 对象。
  function cloneTemplate(template) {
    return JSON.parse(JSON.stringify(template));
  }

  function createCustomTemplateId() {
    const suffix = Math.random().toString(36).slice(2, 8);
    return `tpl_custom_${Date.now().toString(36)}_${suffix}`;
  }

  // 保存当前分组、字段、显隐和自定义字段，并立即切换到新模板。
  function saveCurrentAsCustomTemplate() {
    const source = currentTemplate();
    const input = $("custom-template-name");
    const name = input.value.trim();
    if (!source) {
      toast("请先选择模板");
      return;
    }
    if (!name) {
      input.focus();
      toast("请输入模板名称");
      return;
    }
    if (UI.templates.some((template) => template.name.toLowerCase() === name.toLowerCase())) {
      input.focus();
      toast("模板名称已存在");
      return;
    }

    const template = cloneTemplate(source);
    template.id = createCustomTemplateId();
    template.name = name;
    template.description = `基于“${source.name}”保存的自定义模板`;
    template.builtIn = false;
    template.version = 1;
    template.sourceTemplateId = source.id;

    const previousTemplateId = UI.currentTemplateId;
    UI.templates.push(template);
    UI.customTemplates.push(template);
    UI.orders[template.id] = UI.order.slice();
    UI.currentTemplateId = template.id;
    if (!persist()) {
      // 存储失败时回滚内存数据，避免界面显示一个刷新后就消失的模板。
      UI.templates.pop();
      UI.customTemplates.pop();
      delete UI.orders[template.id];
      UI.currentTemplateId = previousTemplateId;
      showErrors(["保存到我的模板失败：浏览器本地存储不可用或空间不足"]);
      return;
    }

    input.value = "";
    renderAll();
    toast("已保存到我的模板");
  }

  // 应用模板会写入两行表头；Mock 清空示例数据，WPS 只允许空白工作表。
  async function applyCurrentTemplateToSheet() {
    const template = currentTemplate();
    if (!template) return;
    const layout = Core.buildTemplateLayout(template);
    if (layout.columns.length === 0) {
      showErrors(["模板中没有可应用的可见字段"]);
      return;
    }

    const message = UI.adapter.isMock
      ? "浏览器 Mock 将清空当前示例数据并生成模板表头，刷新页面可恢复示例数据。继续吗？"
      : "模板只会写入空白工作表。当前工作表如果有内容将拒绝操作。继续吗？";
    if (!window.confirm(message)) return;

    UI.busy = true;
    updateControls();
    try {
      const result = await UI.adapter.applyTemplateLayout(layout);
      UI.editingFieldId = null;
      UI.result = null;
      UI.stale = false;
      renderAll();
      toast(`已写入 ${result.rowCount} 行模板表头`);
    } catch (err) {
      showErrors([`应用模板失败：${err.message}`]);
    } finally {
      UI.busy = false;
      updateControls();
    }
  }

  // 展示当前模板的来源、版本和有效字段数量，帮助确认选中的是哪套配置。
  function renderCurrentTemplate() {
    const box = $("current-template");
    const template = currentTemplate();
    box.innerHTML = "";
    if (!template) {
      box.className = "hidden";
      return;
    }

    const visibleGroups = Array.isArray(template.groups)
      ? template.groups.filter((group) => group.visible !== false)
      : [];
    const fields = Core.flattenTemplate(template);
    const source = template.builtIn === true ? "内置模板" : "自定义模板";
    const version = Number.isFinite(Number(template.version)) ? Number(template.version) : 1;

    const head = document.createElement("div");
    head.className = "template-current-head";
    const label = document.createElement("span");
    label.className = "dim";
    label.textContent = "当前模板";
    const name = document.createElement("strong");
    name.textContent = template.name || template.id;
    const badge = document.createElement("span");
    badge.className = "template-badge";
    badge.textContent = source;
    head.append(label, name, badge);

    const meta = document.createElement("div");
    meta.id = "template-meta";
    meta.textContent = `版本 ${version} · ${visibleGroups.length} 个分组 · ${fields.length} 个字段`;

    box.className = "";
    box.append(head, meta);
  }

  function renderTemplateGroups() {
    const list = $("template-groups");
    const template = currentTemplate();
    list.innerHTML = "";
    if (!template) return;

    const groups = Array.isArray(template.groups)
      ? template.groups.filter((group) => group.visible !== false)
      : [];
    // 当前只渲染可见分组标签；字段显隐编辑和“我的模板”在后续步骤实现。
    groups.forEach((group) => {
      const chip = document.createElement("span");
      chip.className = "group-chip";
      chip.textContent = group.label || group.key;
      list.appendChild(chip);
    });
  }

  // 模板配置区展示完整分组和字段状态，并通过复选框临时调整字段显隐。
  function renderTemplateConfig() {
    const panel = $("template-config");
    const list = $("template-config-groups");
    const template = currentTemplate();
    // 重新渲染时保留用户已展开的分组；切换模板时恢复为默认展开第一组。
    const openGroupKeys = new Set(
      Array.from(list.querySelectorAll("details.config-group[open]"))
        .map((element) => element.dataset.groupKey)
        .filter(Boolean)
    );
    const sameTemplate = list.dataset.templateId === (template && template.id);
    list.innerHTML = "";
    panel.classList.toggle("hidden", !template);
    if (!template) {
      delete list.dataset.templateId;
      return;
    }
    list.dataset.templateId = template.id;

    const groups = Array.isArray(template.groups) ? template.groups : [];
    groups.forEach((group, groupIndex) => {
      const groupBox = document.createElement("details");
      groupBox.className = "config-group";
      groupBox.dataset.groupKey = group.key;
      // 默认展开第一个分组，其他分组保持折叠，避免窄任务窗格一次显示过长。
      groupBox.open = sameTemplate ? openGroupKeys.has(group.key) : groupIndex === 0;

      const fields = Array.isArray(group.fields) ? group.fields : [];
      const visibleCount = fields.filter((field) => field.visible !== false).length;
      const groupState = group.visible === false ? "分组隐藏" : `${visibleCount}/${fields.length} 个字段可见`;
      const summary = document.createElement("summary");
      summary.textContent = `${group.label || group.key} · ${groupState}`;
      groupBox.appendChild(summary);

      fields.forEach((field, fieldIndex) => {
        const row = document.createElement("div");
        row.className = "config-field";

        const toggle = document.createElement("input");
        toggle.type = "checkbox";
        toggle.checked = field.visible !== false;
        toggle.disabled = group.visible === false;
        toggle.dataset.groupIndex = String(groupIndex);
        toggle.dataset.fieldIndex = String(fieldIndex);
        toggle.setAttribute("aria-label", `显示字段 ${field.label || field.key}`);

        const state = document.createElement("span");
        const hidden = field.visible === false;
        state.className = "visibility-state" + (hidden ? " hidden-field" : "");
        state.textContent = hidden ? "隐藏" : "显示";

        const title = document.createElement("span");
        title.className = "config-field-title";
        const fieldId = `${group.key}.${field.key}`;
        if (field.builtIn === false && UI.editingFieldId === fieldId) {
          const input = document.createElement("input");
          input.className = "field-name-input";
          input.value = field.label || field.key;
          input.setAttribute("aria-label", "自定义字段名称");
          let cancelled = false;
          input.addEventListener("keydown", (ev) => {
            if (ev.key === "Enter") input.blur();
            if (ev.key === "Escape") {
              cancelled = true;
              UI.editingFieldId = null;
              renderTemplateConfig();
            }
          });
          input.addEventListener("blur", () => {
            if (!cancelled) commitCustomFieldRename(groupIndex, fieldIndex, input.value);
          });
          title.appendChild(input);
          requestAnimationFrame(() => input.focus());
        } else {
          const name = document.createElement("span");
          name.className = "field-name";
          name.textContent = field.label || field.key;
          title.appendChild(name);
          if (field.builtIn === false) {
            const tag = document.createElement("span");
            tag.className = "custom-tag";
            tag.textContent = "自定义";
            title.appendChild(tag);
          }
        }

        const key = document.createElement("span");
        key.className = "field-key";
        key.textContent = `${group.key}.${field.key}`;

        row.append(toggle, title, state, key);
        if (field.builtIn === false) {
          const actions = document.createElement("span");
          actions.className = "custom-field-actions";
          const rename = document.createElement("button");
          rename.type = "button";
          rename.textContent = "重命名";
          rename.addEventListener("click", () => beginCustomFieldRename(fieldId));
          const remove = document.createElement("button");
          remove.type = "button";
          remove.className = "danger";
          remove.textContent = "删除";
          remove.addEventListener("click", () => deleteCustomField(groupIndex, fieldIndex));
          actions.append(rename, remove);
          row.appendChild(actions);
        }
        groupBox.appendChild(row);
      });

      const addButton = document.createElement("button");
      addButton.type = "button";
      addButton.className = "add-custom-field";
      addButton.textContent = "+ 添加自定义字段";
      addButton.disabled = group.visible === false;
      addButton.addEventListener("click", () => addCustomField(groupIndex));
      groupBox.appendChild(addButton);
      list.appendChild(groupBox);
    });
  }

  // 显隐、添加、改名和删除都会改变有效字段集合，因此统一重建列表和顺序。
  function refreshTemplateFields(template) {
    UI.fields = Core.flattenTemplate(template);
    UI.order = Core.normalizeOrder(UI.orders[template.id], UI.fields);
    if (UI.result) UI.stale = true;
  }

  // 字段显隐是当前会话配置；P12 才会把调整保存为“我的模板”。
  function updateFieldVisibility(groupIndex, fieldIndex, visible) {
    const template = currentTemplate();
    const group = template && template.groups && template.groups[groupIndex];
    const field = group && group.fields && group.fields[fieldIndex];
    if (!field) return;

    const allFields = template.groups.flatMap((item) => item.fields || []);
    const visibleCount = allFields.filter((item) => item.visible !== false).length;
    const nextVisibleCount = visibleCount - (field.visible === false ? 0 : 1) + (visible ? 1 : 0);
    if (nextVisibleCount === 0) {
      toast("至少保留一个可见字段");
      renderTemplateConfig();
      return;
    }

    field.visible = visible;
    refreshTemplateFields(template);
    renderAll();
    toast(`${field.label || field.key}已${visible ? "显示" : "隐藏"}`);
  }

  // 自定义字段使用独立 key；只改显示名称不会影响字段身份。
  function addCustomField(groupIndex) {
    const template = currentTemplate();
    const group = template && template.groups && template.groups[groupIndex];
    if (!group || group.visible === false) return;

    const keys = new Set(
      template.groups.flatMap((item) => (item.fields || []).map((field) => field.key))
    );
    let number = 1;
    while (keys.has(`custom${number}`)) number++;
    const field = {
      key: `custom${number}`,
      label: `自定义字段${number}`,
      visible: true,
      builtIn: false,
      aliases: [],
    };
    group.fields = Array.isArray(group.fields) ? group.fields : [];
    group.fields.push(field);
    UI.editingFieldId = `${group.key}.${field.key}`;
    refreshTemplateFields(template);
    renderAll();
    toast(`已添加${field.label}`);
  }

  function beginCustomFieldRename(fieldId) {
    UI.editingFieldId = fieldId;
    renderTemplateConfig();
  }

  function commitCustomFieldRename(groupIndex, fieldIndex, value) {
    const template = currentTemplate();
    const field = template && template.groups && template.groups[groupIndex] &&
      template.groups[groupIndex].fields && template.groups[groupIndex].fields[fieldIndex];
    if (!field || field.builtIn !== false) return;

    const label = String(value || "").trim();
    UI.editingFieldId = null;
    if (!label) {
      renderTemplateConfig();
      toast("自定义字段名称不能为空");
      return;
    }
    field.label = label;
    refreshTemplateFields(template);
    renderAll();
    toast("字段名称已更新");
  }

  function deleteCustomField(groupIndex, fieldIndex) {
    const template = currentTemplate();
    const group = template && template.groups && template.groups[groupIndex];
    const field = group && group.fields && group.fields[fieldIndex];
    if (!field || field.builtIn !== false) return;

    const visibleCount = template.groups
      .flatMap((item) => item.fields || [])
      .filter((item) => item.visible !== false).length;
    if (field.visible !== false && visibleCount <= 1) {
      toast("至少保留一个可见字段");
      return;
    }
    if (!window.confirm(`确定删除自定义字段“${field.label || field.key}”吗？`)) return;

    group.fields.splice(fieldIndex, 1);
    UI.editingFieldId = null;
    refreshTemplateFields(template);
    renderAll();
    toast("自定义字段已删除");
  }

  // 自动映射只读展示：字段名/别名匹配模板，P16 再允许人工修正。
  function renderHeaderMapping() {
    const section = $("mapping-section");
    const list = $("mapping-list");
    const stats = $("mapping-stats");
    const hint = $("mapping-hint");
    const template = currentTemplate();
    list.innerHTML = "";
    if (!template) {
      section.classList.add("hidden");
      stats.textContent = "";
      hint.textContent = "";
      return;
    }

    let headerInfo;
    try {
      headerInfo = UI.adapter.getHeaderInfo();
    } catch (err) {
      section.classList.remove("hidden");
      stats.textContent = "";
      hint.textContent = `读取表头映射失败：${err.message}`;
      return;
    }
    const sourceHeaders = headerInfo.headers || [];
    const result = Core.autoMatchHeaders(sourceHeaders, UI.fields);
    section.classList.remove("hidden");
    stats.textContent = `${result.matchedCount}/${sourceHeaders.length} 列已匹配`;

    result.matches.forEach((match) => {
      const row = document.createElement("div");
      row.className = "mapping-row" + (match.field ? "" : " unmatched");
      row.title = match.reason;

      const source = document.createElement("span");
      source.className = "mapping-source";
      const column = document.createElement("span");
      column.className = "col";
      column.textContent = match.header.column || Core.colLetter(match.headerIndex);
      source.append(column, document.createTextNode(match.header.fullLabel || match.header.label || ""));

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

    const unmatchedSourceCount = sourceHeaders.length - result.matchedCount;
    const mismatchMessages = [];
    if (unmatchedSourceCount > 0) mismatchMessages.push(`${unmatchedSourceCount} 个源表头未匹配`);
    if (result.unmatchedTemplateFields.length > 0) {
      mismatchMessages.push(`${result.unmatchedTemplateFields.length} 个模板字段未找到源列`);
    }
    hint.textContent = mismatchMessages.length
      ? mismatchMessages.join("；")
      : "所有源表头均已自动匹配";
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
      const sourceColumn = Number.isInteger(field.sourceIndex)
        ? Core.colLetter(field.sourceIndex)
        : "新";
      row.innerHTML =
        `<span class="handle" aria-hidden="true">≡</span>` +
        `<div class="field-body">` +
        `  <div class="field-title"><span>${esc(field.label)}</span><span class="mono">${esc(field.path.join("."))}</span></div>` +
        (group ? `  <div class="field-path">${esc(group)}</div>` : "") +
        `</div>` +
        // 隐藏字段后仍显示字段真实对应的源列字母，不能按当前可见序号重新编号。
        `<span class="field-col">${sourceColumn}</span>`;

      row.addEventListener("dragstart", (ev) => onDragStart(ev, index));
      row.addEventListener("dragover", (ev) => onDragOver(ev, index));
      row.addEventListener("drop", onDrop);
      row.addEventListener("dragend", onDragEnd);
      row.addEventListener("mouseenter", async () => {
        row.classList.add("hover");
        // 高亮使用源列位置，避免隐藏字段后悬停到错误的列。
        await UI.adapter.highlightColumn(Number.isInteger(field.sourceIndex) ? field.sourceIndex : null);
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
    const hasUnmappedField = UI.fields.some((field) => field.builtIn === false);
    const canUpload = hasTemplate && UI.fields.length > 0 && !hasUnmappedField && !UI.busy && !UI.drag;
    $("btn-upload").disabled = !canUpload;
    $("tpl-select").disabled = UI.busy || !!UI.drag;
    // 拖拽或上传期间禁止清空模板，避免界面状态与进行中的操作冲突。
    $("btn-clear-template").disabled = UI.busy || !!UI.drag;
    $("btn-save-template").disabled = !hasTemplate || UI.busy || !!UI.drag;
    $("btn-apply-template").disabled = !hasTemplate || UI.busy || !!UI.drag;
    $("upload-hint").textContent = !hasTemplate
      ? "选择模板后即可拖动排序并上传"
      : hasUnmappedField
        ? "存在尚未映射的自定义字段，请先删除或完成后续映射配置"
        : "上传会按当前字段顺序读取源表头下方的数据，不修改工作表";
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

    // 适配器返回 1 基行列号，Core.colLetter 接收 0 基索引。
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

  function renderSourceHeaders() {
    const list = $("source-headers");
    const count = $("header-count");
    if (!list || !count) return;

    let headerInfo;
    try {
      if (UI.adapter && typeof UI.adapter.getHeaderInfo === "function") {
        headerInfo = UI.adapter.getHeaderInfo();
      } else {
        const fallback = UI.adapter && typeof UI.adapter.getSourceHeaders === "function"
          ? UI.adapter.getSourceHeaders()
          : [];
        headerInfo = {
          mode: "single",
          headers: fallback.map((label, index) => ({
            fullLabel: label,
            columnIndex: index,
            column: Core.colLetter(index),
          })),
        };
      }
    } catch (err) {
      count.textContent = "";
      list.className = "error-text";
      list.textContent = `读取源表头失败：${err.message}`;
      return;
    }

    const headers = Array.isArray(headerInfo.headers) ? headerInfo.headers : [];
    const headerMode = headerInfo.mode === "grouped" ? "两行分组表头" : "单行表头";
    count.textContent = `${headers.length} 列 · ${headerMode}`;
    list.className = "";
    list.innerHTML = "";
    if (headers.length === 0) {
      list.className = "dim";
      list.textContent = "当前工作表没有可读取的表头";
      return;
    }

    // 按源列顺序生成表头标签，空表头也保留列位置。
    // 两行表头使用“分组 / 字段”作为完整显示名，并保留实际列字母。
    headers.forEach((header, index) => {
      const chip = document.createElement("span");
      const fullLabel = header.fullLabel || header.label || "";
      chip.className = "header-chip" + (fullLabel ? "" : " empty");
      chip.title = header.path ? header.path.join(" / ") : fullLabel;

      const column = document.createElement("span");
      column.className = "col";
      column.textContent = header.column || Core.colLetter(index);

      chip.append(column, document.createTextNode(fullLabel || "（空表头）"));
      list.appendChild(chip);
    });
  }

  function formatPreviewValue(value) {
    // 日期仅用于预览展示；实际导出仍由 Core.convertCell 转为 ISO 字符串。
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? "" : value.toISOString().slice(0, 10);
    }
    return value === null || value === undefined ? "" : String(value);
  }

  function renderSheetPanel() {
    const panel = $("sheet-panel");
    const view = UI.adapter.getSheetView ? UI.adapter.getSheetView() : null;
    if (!view) {
      panel.innerHTML = currentTemplate()
        ? '<div class="dim">当前工作表已读取，源表头见上方。</div>'
        : '<div class="dim">选择模板后会读取当前工作表。</div>';
      return;
    }
    const headers = view.headers || [];
    const groupHeaders = Array.isArray(view.groupHeaders) ? view.groupHeaders : [];
    const previewRows = view.rows.slice(0, PREVIEW_ROW_LIMIT);
    const colCount = Math.max(
      headers.length,
      ...previewRows.map((row) => row.length),
      0
    );
    const rowSummary = groupHeaders.length && view.rows.length === 0
      ? "模板表头已生成，暂无业务数据"
      : view.rows.length > PREVIEW_ROW_LIMIT
        ? `显示前 ${PREVIEW_ROW_LIMIT} 行，共 ${view.rows.length} 行`
        : `显示 ${view.rows.length} 行`;
    const typeCounts = { text: 0, number: 0, date: 0, boolean: 0, empty: 0 };
    // 统计当前预览区域，帮助确认空值、数字和日期没有被错误转换。
    previewRows.forEach((row) => {
      for (let c = 0; c < colCount; c++) typeCounts[Core.classifyCell(row[c])]++;
    });
    const typeLabels = { text: "文本", number: "数字", date: "日期", boolean: "布尔", empty: "空值" };
    const typeParts = Object.keys(typeLabels)
      .filter((type) => typeCounts[type] > 0)
      .map((type) => `${typeLabels[type]} <strong>${typeCounts[type]}</strong>`);
    const layoutLabel = groupHeaders.length
      ? "第 1 行分组表头 / 第 2 行字段表头"
      : "第 1 行源表头";
    let html = `<div class="sheet-name">模拟工作表：${esc(view.name)}` +
      `（${layoutLabel}） · ${rowSummary}</div>`;
    if (typeParts.length) {
      html += `<div class="type-summary">类型：${typeParts.join(" · ")}</div>`;
    }
    html += '<table class="sheet"><tr><th></th>';
    for (let c = 0; c < colCount; c++) {
      html += `<th class="col-letter">${Core.colLetter(c)}</th>`;
    }
    html += "</tr>";

    let rowNo = 1;
    if (groupHeaders.length) {
      html += `<tr><td class="row-no">${rowNo++}</td>`;
      let columnIndex = 0;
      while (columnIndex < colCount) {
        const group = groupHeaders.find((item) =>
          item.startIndex === columnIndex && item.endIndex >= item.startIndex
        );
        if (group) {
          const span = group.endIndex - group.startIndex + 1;
          html += `<td class="group-header" colspan="${span}">${esc(group.label || "")}</td>`;
          columnIndex = group.endIndex + 1;
        } else {
          html += '<td class="group-header"></td>';
          columnIndex++;
        }
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

    previewRows.forEach((row) => {
      html += `<tr><td class="row-no">${rowNo++}</td>`;
      for (let c = 0; c < colCount; c++) {
        const value = row[c];
        const hover = c === view.highlightCol ? " hover" : "";
        html += `<td class="${hover}">${esc(formatPreviewValue(value))}</td>`;
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
      await UI.writeChain;

      // 读取完整源列范围；buildRecords 再按每个字段的 sourceIndex 取值。
      const sourceColumnCount = leaves.reduce(
        (max, leaf) => Math.max(max, (Number.isInteger(leaf.sourceIndex) ? leaf.sourceIndex : 0) + 1),
        leaves.length
      );
      const rows = await UI.adapter.readRows(sourceColumnCount);
      const json = Core.buildExport(template, UI.adapter.getSheetName(), UI.order, UI.fields, rows);
      UI.result = json;
      UI.stale = false;
      renderResult();

      const filename = `${safeFileName(UI.adapter.getSheetName())}-${safeFileName(template.name)}.json`;
      await UI.adapter.uploadJson(json, filename);
      toast(`已导出 ${rows.length} 行数据：${filename}`);
      renderSheetInfo();
      renderSourceHeaders();
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
    const stale = UI.stale ? " · 字段配置或顺序已变化，请重新上传" : "";
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
    renderSourceHeaders();
    renderTemplateBar();
    renderHeaderMapping();
    renderFieldList();
    renderSheetPanel();
    updateControls();
    renderResult();
  }

  // 本地模板可能来自旧版本或被手工修改，加载前做最小结构检查。
  function isValidStoredTemplate(template) {
    return !!template &&
      typeof template.id === "string" &&
      typeof template.name === "string" &&
      Array.isArray(template.groups);
  }

  function init(templates, adapter) {
    UI.adapter = adapter;
    const saved = Store.load();
    // 只恢复字段顺序；上次选择的模板不恢复，因此初始状态始终是“请选择模板”。
    UI.orders = saved.orders || {};
    const templateIds = new Set(templates.map((template) => template.id));
    UI.customTemplates = [];
    saved.customTemplates.forEach((template) => {
      if (isValidStoredTemplate(template) && !templateIds.has(template.id)) {
        templateIds.add(template.id);
        UI.customTemplates.push({ ...template, builtIn: false });
      }
    });
    UI.templates = templates.concat(UI.customTemplates);

    $("tpl-select").addEventListener("change", (ev) => {
      if (ev.target.value) activateTemplate(ev.target.value, true);
    });
    $("btn-upload").addEventListener("click", upload);
    $("btn-copy").addEventListener("click", copyResult);
    $("btn-export").addEventListener("click", exportResult);
    $("btn-clear-template").addEventListener("click", clearCurrentTemplate);
    $("btn-apply-template").addEventListener("click", applyCurrentTemplateToSheet);
    $("btn-save-template").addEventListener("click", saveCurrentAsCustomTemplate);
    $("custom-template-name").addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") saveCurrentAsCustomTemplate();
    });
    $("template-config-groups").addEventListener("change", (ev) => {
      if (!(ev.target instanceof HTMLInputElement) || ev.target.type !== "checkbox") return;
      updateFieldVisibility(
        Number(ev.target.dataset.groupIndex),
        Number(ev.target.dataset.fieldIndex),
        ev.target.checked
      );
    });

    document.addEventListener("keydown", (ev) => {
      if (ev.key === "Escape" && UI.drag) cancelDrag();
    });
    window.addEventListener("blur", () => {
      if (UI.drag) cancelDrag();
    });

    if (adapter.onDataChange) {
      adapter.onDataChange(() => {
        renderSheetInfo();
        renderSourceHeaders();
        renderHeaderMapping();
        if (UI.result) {
          UI.stale = true;
          renderResult();
        }
      });
    }

    $("env-tag").textContent = adapter.isMock ? "浏览器 mock 模式" : "WPS 环境";
    renderAll();
  }

  root.UI = { init };
})(typeof self !== "undefined" ? self : globalThis);
