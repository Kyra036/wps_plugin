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

    getSourceHeaders() {
      // 返回副本，防止 UI 修改数组时污染 Mock 工作表。
      return this.headers.map((value) => value === null || value === undefined ? "" : String(value));
    }

    getSheetInfo() {
      return {
        name: this.name,
        startRow: 1,
        startColumn: 1,
        rowCount: 1 + this.rows.length,
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

    // 模拟工作表视图，供任务窗格预览源表头和业务数据。
    getSheetView() {
      return {
        name: this.name,
        headers: this.headers.slice(),
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

    _headerKey(sheet) {
      let book = "";
      try {
        book = this.app.ActiveWorkbook.Name;
      } catch (_) {
        /* 工作簿名不可用时退化为按工作表名记录 */
      }
      return `field_mapping_header|${book}|${sheet.Name}`;
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

    _usedRange(sheet) {
      return typeof sheet.UsedRange === "function" ? sheet.UsedRange() : sheet.UsedRange;
    }

    async readRows(colCount) {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) return [];
      // 新工作簿从 UsedRange 首行读取表头；旧工作簿则跳过历史插件表头。
      const sourceHeaderRow = this._hasHeaderMark(sheet)
        ? LEGACY_SOURCE_HEADER_ROW
        : (Number(used.Row) || 1);
      const dataStartRow = sourceHeaderRow + 1;
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
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) return [];

      // 兼容旧工作簿：历史插件表头存在时，原表头在第 2 行。
      const row = this._hasHeaderMark(sheet) ? LEGACY_SOURCE_HEADER_ROW : (Number(used.Row) || 1);
      const startColumn = Number(used.Column) || 1;
      const columnCount = Number(used.Columns && used.Columns.Count) || 0;
      const headers = [];
      for (let offset = 0; offset < columnCount; offset++) {
        const value = sheet.Cells.Item(row, startColumn + offset).Value2;
        headers.push(value === null || value === undefined ? "" : String(value).trim());
      }
      return headers;
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
