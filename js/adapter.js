// 工作表适配层：统一 WPS 真实环境与浏览器 Mock 环境。
// 约定：插件表头固定写入工作表第 1 行；原表头在第 2 行；业务数据从第 3 行开始。
(function (root) {
  "use strict";

  const HEADER_ROW = 1;
  const SOURCE_HEADER_ROW = 2;
  const DATA_START_ROW = SOURCE_HEADER_ROW + 1;

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
      this.rows = mockSheet.rows.map((row) => row.slice());
      this.hasPluginHeader = false;
      this.pluginLabels = [];
      this.highlightCol = null;
    }

    async ensurePluginHeader() {
      if (this.hasPluginHeader) {
        return { created: false, headerWidth: this.pluginLabels.length };
      }
      this.hasPluginHeader = true;
      return { created: true, headerWidth: 0 };
    }

    async writePluginHeader(labels) {
      this.pluginLabels = labels.slice();
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

    getSheetInfo() {
      return {
        name: this.name,
        startRow: 1,
        startColumn: 1,
        rowCount: 1 + this.rows.length + (this.hasPluginHeader ? 1 : 0),
        columnCount: Math.max(
          this.headers.length,
          this.pluginLabels.length,
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

    // 模拟工作表视图，供任务窗格内预览插件表头、原始表头和业务数据
    getSheetView() {
      return {
        name: this.name,
        hasPluginHeader: this.hasPluginHeader,
        pluginLabels: this.pluginLabels.slice(),
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

    _markHeader(sheet) {
      const key = this._headerKey(sheet);
      try {
        localStorage.setItem(key, "1");
      } catch (_) {
        /* 忽略任务窗格存储失败 */
      }
      try {
        this.app.PluginStorage.setItem(key, "1");
      } catch (_) {
        /* 忽略 WPS 存储失败 */
      }
    }

    _usedRange(sheet) {
      return typeof sheet.UsedRange === "function" ? sheet.UsedRange() : sheet.UsedRange;
    }

    // WPS JS API 中 Rows 是集合对象（Rows.Item(1)），不是可调用函数。
    // 不同版本对象模型略有差异，这里按 Rows 集合 → EntireRow → Range 逐级降级。
    _insertRow(sheet, row) {
      const attempts = [];
      const rows = sheet.Rows;
      if (typeof rows === "function") {
        attempts.push(() => rows.call(sheet, row).Insert());
      } else if (rows && typeof rows.Item === "function") {
        attempts.push(() => rows.Item(row).Insert());
      }

      const cell = sheet.Cells && sheet.Cells.Item ? sheet.Cells.Item(row, 1) : null;
      if (cell) {
        attempts.push(() => {
          const entire = typeof cell.EntireRow === "function" ? cell.EntireRow() : cell.EntireRow;
          entire.Insert();
        });
      }

      if (typeof sheet.Range === "function") {
        attempts.push(() => {
          const range = sheet.Range(`A${row}:A${row}`);
          const entire = typeof range.EntireRow === "function" ? range.EntireRow() : range.EntireRow;
          entire.Insert();
        });
      }

      let lastError = null;
      for (const attempt of attempts) {
        try {
          attempt();
          return;
        } catch (err) {
          lastError = err;
        }
      }
      throw lastError || new Error("当前 WPS 版本不支持插入整行");
    }

    _headerWidth(sheet) {
      try {
        const used = this._usedRange(sheet);
        if (!used) return 0;
        const lastCol = used.Column + used.Columns.Count - 1;
        let width = 0;
        for (let c = 1; c <= Math.min(lastCol, 200); c++) {
          const value = sheet.Cells.Item(HEADER_ROW, c).Value2;
          if (value !== undefined && value !== null && String(value) !== "") width = c;
        }
        return width;
      } catch (_) {
        return 0;
      }
    }

    async ensurePluginHeader() {
      const sheet = this._sheet();
      if (this._hasHeaderMark(sheet)) {
        return { created: false, headerWidth: this._headerWidth(sheet) };
      }
      this._insertRow(sheet, HEADER_ROW);
      this._adjustFreezePanes();
      this._markHeader(sheet);
      return { created: true, headerWidth: 0 };
    }

    _adjustFreezePanes() {
      try {
        const win = this.app.ActiveWindow;
        if (win.FreezePanes) {
          const split = win.SplitRow + 1;
          win.FreezePanes = false;
          win.SplitRow = split;
          win.FreezePanes = true;
        }
      } catch (_) {
        /* 冻结行调整失败不影响主流程 */
      }
    }

    async writePluginHeader(labels, previousCount) {
      const sheet = this._sheet();
      const count = Math.max(labels.length, previousCount || 0);
      for (let i = 0; i < count; i++) {
        const cell = sheet.Cells.Item(HEADER_ROW, i + 1);
        if (i < labels.length) {
          cell.Value2 = labels[i];
          try {
            cell.Interior.Color = 0xfeeadb; // #DBEAFE（BGR）
            cell.Font.Color = 0xaf401e; // #1E40AF（BGR）
            cell.Font.Bold = true;
          } catch (_) {
            /* 样式写入失败不影响表头内容 */
          }
        } else {
          cell.Value2 = "";
          try {
            cell.Interior.ColorIndex = -4142; // xlNone
            cell.Font.ColorIndex = -4105; // xlAutomatic
            cell.Font.Bold = false;
          } catch (_) {
            /* 样式写入失败不影响表头内容 */
          }
        }
      }
    }

    async readRows(colCount) {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) return [];
      const lastRow = used.Row + used.Rows.Count - 1;
      if (lastRow < DATA_START_ROW) return [];
      const rows = [];
      for (let r = DATA_START_ROW; r <= lastRow; r++) {
        const row = [];
        for (let c = 0; c < colCount; c++) {
          let value = sheet.Cells.Item(r, c + 1).Value2;
          if (value instanceof Date) value = value.toISOString();
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
      return {
        name: sheet.Name,
        startRow: Number(used.Row) || 1,
        startColumn: Number(used.Column) || 1,
        rowCount: Number(used.Rows && used.Rows.Count) || 0,
        columnCount: Number(used.Columns && used.Columns.Count) || 0,
      };
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
