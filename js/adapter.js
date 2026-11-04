// 工作表适配层：统一 WPS 真实环境与浏览器 Mock 环境，单表流程只读源工作表。
(function (root) {
  "use strict";

  // 旧版本曾在第 1 行插入插件表头；兼容这类工作簿时，源表头位于第 2 行。
  const LEGACY_SOURCE_HEADER_ROW = 2;

  // JSON 只能保存字符串，Mock 日期列在这里恢复为 Date，模拟 WPS 单元格的真实类型。
  function parseMockDate(value) {
    if (value instanceof Date) return value;
    if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
      const date = new Date(`${value}T00:00:00.000Z`);
      if (!Number.isNaN(date.getTime())) return date;
    }
    return value;
  }

  // WPS Range 使用 A/B/C 地址；该函数把 0 基索引转换成列字母。
  function columnLetter(index) {
    let value = "";
    let number = index + 1;
    while (number > 0) {
      const remainder = (number - 1) % 26;
      value = String.fromCharCode(65 + remainder) + value;
      number = Math.floor((number - 1) / 26);
    }
    return value;
  }

  function downloadJson(json, filename) {
    const blob = new Blob([JSON.stringify(json, null, 2)], { type: "application/json;charset=utf-8" });
    const anchor = document.createElement("a");
    const url = URL.createObjectURL(blob);
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  // ---------- Mock 适配器：内存模拟工作表 ----------
  class MockAdapter {
    constructor(mockSheet) {
      this.isMock = true;
      this.name = mockSheet.name;
      this.headers = (mockSheet.headers || []).slice();
      this.groupHeaders = [];
      // dateColumns 使用表头名称声明，避免把列字母或位置写死在 Mock 数据里。
      const dateColumns = new Set(mockSheet.dateColumns || []);
      const dateIndexes = new Set(
        this.headers
          .map((header, index) => dateColumns.has(header) ? index : -1)
          .filter((index) => index >= 0)
      );
      this.rows = mockSheet.rows.map((row) =>
        row.map((value, index) => dateIndexes.has(index) ? parseMockDate(value) : value)
      );
      this.highlightCol = null;
    }

    async readRows(colCount) {
      return this.rows.map((row) => {
        const out = [];
        for (let i = 0; i < colCount; i++) out.push(row[i] === undefined ? null : row[i]);
        return out;
      });
    }

    getSheetName() {
      return this.name;
    }

    // 返回统一的表头结构；Mock 应用模板后自动进入 grouped 模式。
    getHeaderInfo() {
      const grouped = this.groupHeaders.length > 0;
      const headers = this.headers.map((label, index) => {
        const group = grouped
          ? this.groupHeaders.find((item) => index >= item.startIndex && index <= item.endIndex)
          : null;
        const groupLabel = group ? String(group.label || "").trim() : "";
        const fieldLabel = String(label || "").trim();
        return {
          label: fieldLabel,
          groupLabel,
          fullLabel: groupLabel ? `${groupLabel} / ${fieldLabel}` : fieldLabel,
          path: groupLabel ? [groupLabel, fieldLabel] : [fieldLabel],
          columnIndex: index,
          column: columnLetter(index),
        };
      });
      return {
        mode: grouped ? "grouped" : "single",
        headerRows: grouped ? 2 : 1,
        dataStartRow: grouped ? 3 : 2,
        startColumn: 1,
        headers,
      };
    }

    getSourceHeaders() {
      // 返回副本，防止 UI 修改数组时污染 Mock 工作表。
      return this.getHeaderInfo().headers.map((header) => header.fullLabel);
    }

    getSheetInfo() {
      return {
        name: this.name,
        startRow: 1,
        startColumn: 1,
        rowCount: (this.groupHeaders.length ? 1 : 0) + 1 + this.rows.length,
        columnCount: Math.max(
          this.headers.length,
          ...this.rows.map((row) => row.length),
          0
        ),
      };
    }

    async highlightColumn(col) {
      this.highlightCol = col;
    }

    onDataChange() {}

    async uploadJson(json, filename) {
      downloadJson(json, filename);
    }

    // Mock 应用模板时清空示例数据，模拟向空白工作表创建两行表头。
    async applyTemplateLayout(layout) {
      this.headers = layout.columns.map((column) => column.label);
      this.groupHeaders = layout.groups.map((group) => ({ ...group }));
      this.rows = [];
      this.highlightCol = null;
      return { sheetName: this.name, rowCount: 2, columnCount: this.headers.length };
    }

    // 模拟工作表视图，供任务窗格预览源表头和业务数据。
    getSheetView() {
      return {
        name: this.name,
        headers: this.headers.slice(),
        groupHeaders: this.groupHeaders.map((group) => ({ ...group })),
        rows: this.rows.map((row) => row.slice()),
        highlightCol: this.highlightCol,
      };
    }
  }

  // ---------- WPS 适配器：wps.EtApplication() ----------
  class WpsAdapter {
    constructor() {
      this.isMock = false;
      this.app = window.wps.EtApplication();
    }

    _sheet() {
      return this.app.ActiveSheet;
    }

    _sheetStorageKey(prefix, sheet) {
      let book = "";
      try {
        book = this.app.ActiveWorkbook.Name;
      } catch (_) {
        /* 工作簿名不可用时退化为按工作表名记录 */
      }
      return `${prefix}|${book}|${sheet.Name}`;
    }

    _headerKey(sheet) {
      return this._sheetStorageKey("field_mapping_header", sheet);
    }

    _templateLayoutKey(sheet) {
      return this._sheetStorageKey("field_mapping_template_layout", sheet);
    }

    _hasHeaderMark(sheet) {
      // 只用于识别旧版本修改过的工作簿，新流程不会再写入这个标记。
      const key = this._headerKey(sheet);
      try {
        if (localStorage.getItem(key) === "1") return true;
      } catch (_) {
        /* 任务窗格存储不可用时继续查 WPS 存储 */
      }
      try {
        return this.app.PluginStorage.getItem(key) === "1";
      } catch (_) {
        return false;
      }
    }

    _hasTemplateLayoutMark(sheet) {
      const key = this._templateLayoutKey(sheet);
      try {
        return localStorage.getItem(key) === "1" ||
          this.app.PluginStorage.getItem(key) === "1";
      } catch (_) {
        return false;
      }
    }

    _markTemplateLayout(sheet) {
      const key = this._templateLayoutKey(sheet);
      try {
        localStorage.setItem(key, "1");
      } catch (_) {
        // 任务窗格存储不可用时忽略。
      }
      try {
        this.app.PluginStorage.setItem(key, "1");
      } catch (_) {
        // WPS 存储不可用时不影响已经写入的表头。
      }
    }

    _usedRange(sheet) {
      return typeof sheet.UsedRange === "function" ? sheet.UsedRange() : sheet.UsedRange;
    }

    // 只允许写入空白工作表，避免“应用模板”覆盖客户已有数据。
    _isBlankSheet(sheet) {
      const used = this._usedRange(sheet);
      if (!used) return true;
      const rowCount = Number(used.Rows && used.Rows.Count) || 0;
      const columnCount = Number(used.Columns && used.Columns.Count) || 0;
      for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < columnCount; c++) {
          const value = sheet.Cells.Item(used.Row + r, used.Column + c).Value2;
          if (value !== null && value !== undefined && String(value).trim() !== "") return false;
        }
      }
      return true;
    }

    _readRowValues(sheet, row, startColumn, columnCount) {
      const values = [];
      for (let offset = 0; offset < columnCount; offset++) {
        const value = sheet.Cells.Item(row, startColumn + offset).Value2;
        values.push(value === null || value === undefined ? "" : String(value).trim());
      }
      return values;
    }

    // 自动区分单行表头和应用模板后生成的两行分组表头。
    getHeaderInfo() {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) {
        return { mode: "empty", headerRows: 0, dataStartRow: 1, startColumn: 1, headers: [] };
      }

      const startRow = Number(used.Row) || 1;
      const startColumn = Number(used.Column) || 1;
      const columnCount = Number(used.Columns && used.Columns.Count) || 0;
      const grouped = this._hasTemplateLayoutMark(sheet);
      const legacyHeader = !grouped && this._hasHeaderMark(sheet);
      const fieldRow = grouped ? startRow + 1 : legacyHeader ? LEGACY_SOURCE_HEADER_ROW : startRow;
      const fieldValues = this._readRowValues(sheet, fieldRow, startColumn, columnCount);
      let groupValues = [];
      if (grouped) {
        groupValues = this._readRowValues(sheet, startRow, startColumn, columnCount);
      }

      let currentGroup = "";
      const headers = fieldValues.map((fieldLabel, index) => {
        if (grouped && groupValues[index]) currentGroup = groupValues[index];
        const groupLabel = grouped ? currentGroup : "";
        return {
          label: fieldLabel,
          groupLabel,
          fullLabel: groupLabel ? `${groupLabel} / ${fieldLabel}` : fieldLabel,
          path: groupLabel ? [groupLabel, fieldLabel] : [fieldLabel],
          columnIndex: index,
          column: columnLetter(startColumn - 1 + index),
        };
      });
      return {
        mode: grouped ? "grouped" : "single",
        headerRows: grouped ? 2 : 1,
        dataStartRow: fieldRow + 1,
        startColumn,
        headers,
      };
    }

    // 第 1 行写入并合并分组表头，第 2 行写入字段表头。
    async applyTemplateLayout(layout) {
      const sheet = this._sheet();
      if (!this._isBlankSheet(sheet)) {
        throw new Error("当前工作表已有内容，请先新建或切换到空白工作表");
      }

      layout.groups.forEach((group) => {
        const start = columnLetter(group.startIndex);
        const end = columnLetter(group.endIndex);
        const range = sheet.Range(`${start}1:${end}1`);
        try {
          range.Merge();
        } catch (_) {
          try {
            range.Merge = true;
          } catch (_) {
            // 不支持合并时仍保留该组起始单元格的文字。
          }
        }
        range.Value2 = group.label;
        try {
          range.Interior.Color = 0xfeeadb; // 浅蓝色，BGR 顺序
          range.Font.Bold = true;
          range.HorizontalAlignment = -4108; // xlCenter
        } catch (_) {
          // 分组样式失败不影响字段表头。
        }
      });

      layout.columns.forEach((column, index) => {
        const cell = sheet.Cells.Item(2, index + 1);
        cell.Value2 = column.label;
        try {
          cell.Interior.Color = 0xf8fafc;
          cell.Font.Bold = true;
          cell.HorizontalAlignment = -4131; // xlLeft
          sheet.Columns.Item(index + 1).ColumnWidth = column.width;
        } catch (_) {
          // 单个列宽或样式失败不影响其他字段。
        }
      });
      this._markTemplateLayout(sheet);
      return { sheetName: sheet.Name, rowCount: 2, columnCount: layout.columns.length };
    }

    async readRows(colCount) {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) return [];
      // 表头识别逻辑集中到 getHeaderInfo，单行和两行布局使用不同数据起始行。
      const dataStartRow = this.getHeaderInfo().dataStartRow;
      const lastRow = used.Row + used.Rows.Count - 1;
      if (lastRow < dataStartRow) return [];
      const rows = [];
      for (let r = dataStartRow; r <= lastRow; r++) {
        const row = [];
        for (let c = 0; c < colCount; c++) {
          let value = sheet.Cells.Item(r, c + 1).Value2;
          row.push(value === undefined ? null : value);
        }
        rows.push(row);
      }
      return rows;
    }

    getSheetName() {
      return this._sheet().Name;
    }

    getSheetInfo() {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) {
        return { name: sheet.Name, startRow: 0, startColumn: 0, rowCount: 0, columnCount: 0 };
      }
      // WPS Range 的行列号从 1 开始，UI 会转换成 A1:H9 形式。
      return {
        name: sheet.Name,
        startRow: Number(used.Row) || 1,
        startColumn: Number(used.Column) || 1,
        rowCount: Number(used.Rows && used.Rows.Count) || 0,
        columnCount: Number(used.Columns && used.Columns.Count) || 0,
      };
    }

    getSourceHeaders() {
      return this.getHeaderInfo().headers.map((header) => header.fullLabel);
    }

    async highlightColumn() {
      // WPS 侧不做悬停高亮，避免在拖拽预览时干扰真实单元格
    }

    onDataChange(cb) {
      try {
        window.wps.ApiEvent.AddApiEventListener(this.app, "SheetChange", () => cb());
      } catch (_) {
        /* 事件不可用时忽略 */
      }
    }

    async uploadJson(json, filename) {
      downloadJson(json, filename);
    }

    getSheetView() {
      return null;
    }
  }

  root.Adapters = { MockAdapter, WpsAdapter };
})(typeof self !== "undefined" ? self : globalThis);
