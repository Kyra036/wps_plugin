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
    // 初始化 Mock 工作表名称、表头、数据行和默认表头位置。
    constructor(mockSheet) {
      this.isMock = true;
      this.name = mockSheet.name || "Sheet1";
      this.headerRow = Number(mockSheet.headerRow) || 1;
      this.headers = (mockSheet.headers || []).slice();
      this.rows = (mockSheet.rows || []).map((row) => row.slice());
    }

    // 按列数和最大行数返回 Mock 数据，用于浏览器预览和上传前校验。
    async readRows(colCount, maxRows, _headerRow) {
      const limit = Number.isInteger(maxRows) ? Math.max(0, maxRows) : this.rows.length;
      return this.rows.slice(0, limit).map((row) => {
        const values = [];
        for (let index = 0; index < colCount; index++) {
          values.push(row[index] === undefined ? null : row[index]);
        }
        return values;
      });
    }

    // Mock 数据没有尾部格式空白，直接根据数据行数返回有效范围。
    getDataRange(headerInfo, _colCount) {
      const dataStartRow = headerInfo.dataStartRow;
      return {
        lastDataRow: dataStartRow + this.rows.length - 1,
        dataRowCount: this.rows.length,
      };
    }

    // 返回 Mock 工作表名称。
    getSheetName() {
      return this.name;
    }

    // 浏览器 Mock 只在配置的表头行返回真实表头，便于验证指定行功能。
    getHeaderInfo(headerRow) {
      const targetRow = Number(headerRow) || this.headerRow;
      const sameRow = targetRow === this.headerRow;
      const headers = this.headers.map((label, index) => {
        const labelText = sameRow && label !== null && label !== undefined
          ? String(label).trim()
          : "";
        return {
          label: labelText,
          fullLabel: labelText,
          groupLabel: "",
          columnIndex: index,
          column: columnLetter(index),
        };
      });
      return {
        mode: "single",
        headerRows: 1,
        headerRow: targetRow,
        dataStartRow: targetRow + 1,
        headers,
      };
    }

    // 返回当前 Mock 表头名称，供旧的兼容调用使用。
    getSourceHeaders() {
      return this.getHeaderInfo().headers.map((header) => header.fullLabel);
    }

    // 返回 Mock 配置的默认表头行。
    getDefaultHeaderRow() {
      return this.headerRow;
    }

    // 返回工作表名称、起始位置和行列数量等只读信息。
    getSheetInfo() {
      return {
        name: this.name,
        startRow: 1,
        startColumn: 1,
        rowCount: this.headerRow + this.rows.length,
        columnCount: Math.max(
          this.headers.length,
          ...this.rows.map((row) => row.length),
          0
        ),
      };
    }

    // Mock 不监听外部修改；保留空方法以统一适配器接口。
    onDataChange() {}
  }

  // WPS 表格：通过 wps.EtApplication 读取当前活动工作表。
  class WpsAdapter {
    // 获取 WPS 表格应用对象，所有读取都基于当前活动工作簿。
    constructor() {
      this.isMock = false;
      this.app = window.wps.EtApplication();
    }

    // 获取当前活动 Sheet。
    _sheet() {
      return this.app.ActiveSheet;
    }

    // 兼容 UsedRange 是属性或方法的 WPS 版本差异。
    _usedRange(sheet) {
      return typeof sheet.UsedRange === "function" ? sheet.UsedRange() : sheet.UsedRange;
    }

    // 读取一个单元格的原始值，空白 undefined 统一转换为 null。
    _readCell(sheet, row, column) {
      const value = sheet.Cells.Item(row, column).Value2;
      return value === undefined ? null : value;
    }

    // 判断单元格是否为空，空白字符串和纯空格都视为空。
    _isBlankCell(value) {
      return value === undefined || value === null || String(value).trim() === "";
    }

    // UsedRange 可能因格式扩展到空白行，从底部向上找到最后一条有效数据。
    _findLastDataRow(sheet, dataStartRow, candidateLastRow, startColumn, columnCount) {
      for (let row = candidateLastRow; row >= dataStartRow; row--) {
        for (let column = 0; column < columnCount; column++) {
          if (!this._isBlankCell(this._readCell(sheet, row, startColumn + column))) {
            return row;
          }
        }
      }
      return dataStartRow - 1;
    }

    // 计算最后一条有效数据行，并返回需要读取的数据行数。
    getDataRange(headerInfo, colCount) {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used || !colCount) {
        return {
          lastDataRow: headerInfo.dataStartRow - 1,
          dataRowCount: 0,
        };
      }

      const startColumn = Number(used.Column) || 1;
      const candidateLastRow = used.Row + used.Rows.Count - 1;
      const lastDataRow = this._findLastDataRow(
        sheet,
        headerInfo.dataStartRow,
        candidateLastRow,
        startColumn,
        colCount
      );
      return {
        lastDataRow,
        dataRowCount: Math.max(0, lastDataRow - headerInfo.dataStartRow + 1),
      };
    }

    // 读取指定行的表头文字；未指定时默认使用 UsedRange 第一行。
    getHeaderInfo(headerRowOverride) {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) {
        return { mode: "empty", headerRows: 0, headerRow: 1, dataStartRow: 1, headers: [] };
      }

      const firstUsedRow = Number(used.Row) || 1;
      const requestedRow = Number(headerRowOverride);
      const headerRow = Number.isInteger(requestedRow) && requestedRow > 0
        ? requestedRow
        : firstUsedRow;
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
      return {
        mode: "single",
        headerRows: 1,
        headerRow,
        dataStartRow: headerRow + 1,
        headers,
      };
    }

    // 从表头下一行读取数据，跳过 UsedRange 尾部仅有格式的空白行。
    async readRows(colCount, maxRows, headerRow) {
      const sheet = this._sheet();
      const used = this._usedRange(sheet);
      if (!used) return [];

      const headerInfo = this.getHeaderInfo(headerRow);
      const startColumn = Number(used.Column) || 1;
      const dataRange = this.getDataRange(headerInfo, colCount);
      const lastRow = dataRange.lastDataRow;
      const limit = Number.isInteger(maxRows) ? Math.max(0, maxRows) : Infinity;
      const dataRows = [];
      for (let row = headerInfo.dataStartRow; row <= lastRow && dataRows.length < limit; row++) {
        const values = [];
        for (let column = 0; column < colCount; column++) {
          // 以 UsedRange 的实际起始列为准，兼容 BOM 从 B 列或更右侧开始。
          let value = this._readCell(sheet, row, startColumn + column);
          if (value instanceof Date) value = value.toISOString();
          values.push(value);
        }
        dataRows.push(values);
      }
      return dataRows;
    }

    // 返回 WPS 当前活动 Sheet 名称。
    getSheetName() {
      return this._sheet().Name;
    }

    // 返回默认表头行的全部列名。
    getSourceHeaders() {
      return this.getHeaderInfo().headers.map((header) => header.fullLabel);
    }

    // 返回 UsedRange 的起始行，作为源表头行的初始值。
    getDefaultHeaderRow() {
      const used = this._usedRange(this._sheet());
      return used ? (Number(used.Row) || 1) : 1;
    }

    // 返回当前 Sheet 的 UsedRange 信息；空表返回零值。
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

    // 注册 WPS SheetChange 事件；不支持事件的老版本保持静默降级。
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
