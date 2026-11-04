// 工作表适配层：统一浏览器 Mock 和真实 WPS 表格，只负责读取数据。
(function (root) {
  "use strict";

  // 工具函数：WPS Range 使用 A/B/C 地址，数据层使用 0 基列索引。
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

  // 浏览器 Mock：用 JSON 数据模拟一个只读工作表。
  class MockAdapter {
    constructor(mockSheet) {
      this.isMock = true;
      this.name = mockSheet.name || "Sheet1";
      this.headers = (mockSheet.headers || []).slice();
      this.rows = (mockSheet.rows || []).map((row) => row.slice());
    }

    async readRows(colCount, maxRows) {
      const limit = Number.isInteger(maxRows) ? Math.max(0, maxRows) : this.rows.length;
      return this.rows.slice(0, limit).map((row) => {
        const values = [];
        for (let index = 0; index < colCount; index++) {
          values.push(row[index] === undefined ? null : row[index]);
        }
        return values;
      });
    }

    getSheetName() {
      return this.name;
    }

    getHeaderInfo() {
      const headers = this.headers.map((label, index) => {
        const labelText = label === null || label === undefined ? "" : String(label).trim();
        return {
          label: labelText,
          fullLabel: labelText,
          groupLabel: "",
          columnIndex: index,
          column: columnLetter(index),
        };
      });
      return { mode: "single", headerRows: 1, dataStartRow: 2, headers };
    }

    getSourceHeaders() {
      return this.getHeaderInfo().headers.map((header) => header.fullLabel);
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

    onDataChange() {}
  }

  // WPS 表格：通过 wps.EtApplication 读取当前活动工作表。
  class WpsAdapter {
    constructor() {
      this.isMock = false;
      this.app = window.wps.EtApplication();
    }

    _sheet() {
      return this.app.ActiveSheet;
    }

    _usedRange(sheet) {
      return typeof sheet.UsedRange === "function" ? sheet.UsedRange() : sheet.UsedRange;
    }

    _readCell(sheet, row, column) {
      const value = sheet.Cells.Item(row, column).Value2;
      return value === undefined ? null : value;
    }

    getHeaderInfo() {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) {
        return { mode: "empty", headerRows: 0, dataStartRow: 1, headers: [] };
      }

      const headerRow = Number(used.Row) || 1;
      const startColumn = Number(used.Column) || 1;
      const columnCount = Number(used.Columns && used.Columns.Count) || 0;
      const headers = [];
      for (let offset = 0; offset < columnCount; offset++) {
        const value = this._readCell(sheet, headerRow, startColumn + offset);
        const label = value === null ? "" : String(value).trim();
        headers.push({
          label,
          fullLabel: label,
          groupLabel: "",
          columnIndex: offset,
          column: columnLetter(startColumn - 1 + offset),
        });
      }
      return { mode: "single", headerRows: 1, dataStartRow: headerRow + 1, headers };
    }

    async readRows(colCount, maxRows) {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) return [];

      const headerInfo = this.getHeaderInfo();
      const lastRow = used.Row + used.Rows.Count - 1;
      const limit = Number.isInteger(maxRows) ? Math.max(0, maxRows) : Infinity;
      const dataRows = [];
      for (let row = headerInfo.dataStartRow; row <= lastRow && dataRows.length < limit; row++) {
        const values = [];
        for (let column = 0; column < colCount; column++) {
          let value = this._readCell(sheet, row, column + 1);
          if (value instanceof Date) value = value.toISOString();
          values.push(value);
        }
        dataRows.push(values);
      }
      return dataRows;
    }

    getSheetName() {
      return this._sheet().Name;
    }

    getSourceHeaders() {
      return this.getHeaderInfo().headers.map((header) => header.fullLabel);
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

    onDataChange(callback) {
      try {
        window.wps.ApiEvent.AddApiEventListener(this.app, "SheetChange", () => callback());
      } catch (_) {
        // WPS 版本不支持事件时保持基础读取能力。
      }
    }
  }

  root.Adapters = { MockAdapter, WpsAdapter };
})(typeof self !== "undefined" ? self : globalThis);
