// 临时 XLSX 生成器：按平台固定模板输出 BOM 文件，不依赖网络或第三方库。
(function (root) {
  "use strict";

  const MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const HEADERS = [
    ["序号", "partNo"],
    ["名称", "partName"],
    ["材料", "partMaterial"],
    ["规格（宽*长*高）（MM）", "partSpecs"],
    ["数量", "partNum"],
    ["备注", "partRemark"],
    ["加工要求", "partProcessReq"],
  ];

  // 转义 XML 特殊字符，防止单元格内容破坏 XLSX 的 XML 结构。
  function xmlEscape(value) {
    return String(value === undefined || value === null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  // 将 0 基列索引转换为 Excel 列名。
  function columnName(index) {
    let value = "";
    let number = index + 1;
    while (number > 0) {
      const remainder = (number - 1) % 26;
      value = String.fromCharCode(65 + remainder) + value;
      number = Math.floor((number - 1) / 26);
    }
    return value;
  }

  // 文本统一使用 inlineStr，数量和数字保持数值单元格。
  function cellXml(reference, value, forceText) {
    if (!forceText && typeof value === "number" && Number.isFinite(value)) {
      return `<c r="${reference}"><v>${value}</v></c>`;
    }
    return `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
  }

  // 生成工作表 XML：表头固定写入指定行，数据从下一行开始。
  function buildSheetXml(rows, headerRow) {
    const lastRow = Math.max(headerRow, headerRow + rows.length);
    const rowXml = [];
    const headerCells = HEADERS.map(([label], index) =>
      cellXml(`${columnName(index)}${headerRow}`, label, true)
    ).join("");
    rowXml.push(`<row r="${headerRow}" s="1">${headerCells}</row>`);

    rows.forEach((row, rowIndex) => {
      const excelRow = headerRow + rowIndex + 1;
      const cells = HEADERS.map(([, key], columnIndex) =>
        cellXml(
          `${columnName(columnIndex)}${excelRow}`,
          row[key],
          key !== "partNum"
        )
      ).join("");
      rowXml.push(`<row r="${excelRow}">${cells}</row>`);
    });

    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<dimension ref="A${headerRow}:G${lastRow}"/>` +
      `<sheetData>${rowXml.join("")}</sheetData>` +
      `</worksheet>`
    );
  }

  // 将 JavaScript 字符串编码为 UTF-8 字节，兼容旧 WPS WebView。
  function utf8Bytes(text) {
    const encoded = unescape(encodeURIComponent(String(text)));
    const bytes = new Uint8Array(encoded.length);
    for (let index = 0; index < encoded.length; index++) {
      bytes[index] = encoded.charCodeAt(index);
    }
    return bytes;
  }

  // CRC32 用于标准 ZIP 文件头；XLSX 本质上是包含 XML 的 ZIP 包。
  const crcTable = (() => {
    const table = new Uint32Array(256);
    for (let index = 0; index < 256; index++) {
      let value = index;
      for (let bit = 0; bit < 8; bit++) {
        value = (value & 1) ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
      }
      table[index] = value >>> 0;
    }
    return table;
  })();

  // 计算 ZIP 文件所需的 CRC32 校验值。
  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let index = 0; index < bytes.length; index++) {
      crc = crcTable[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  // 按顺序合并多个 Uint8Array 片段。
  function concatBytes(parts) {
    const total = parts.reduce((sum, part) => sum + part.length, 0);
    const result = new Uint8Array(total);
    let offset = 0;
    parts.forEach((part) => {
      result.set(part, offset);
      offset += part.length;
    });
    return result;
  }

  // 以小端序写入 ZIP 头中的 16 位整数。
  function writeUint16(view, offset, value) {
    view.setUint16(offset, value, true);
  }

  // 以小端序写入 ZIP 头中的 32 位整数。
  function writeUint32(view, offset, value) {
    view.setUint32(offset, value >>> 0, true);
  }

  // 将日期转换为 ZIP 所需的 DOS 日期和时间字段。
  function dosDateTime(date) {
    const year = Math.max(1980, date.getFullYear());
    return {
      time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
      date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
    };
  }

  // 使用“存储”方式生成 ZIP，不压缩但结构标准，Excel 和平台解析器都能读取。
  function buildZip(entries) {
    const localParts = [];
    const centralParts = [];
    const now = dosDateTime(new Date());
    let offset = 0;

    entries.forEach((entry) => {
      const nameBytes = utf8Bytes(entry.name);
      const dataBytes = utf8Bytes(entry.content);
      const checksum = crc32(dataBytes);
      const localHeader = new Uint8Array(30);
      const localView = new DataView(localHeader.buffer);
      writeUint32(localView, 0, 0x04034b50);
      writeUint16(localView, 4, 20);
      writeUint16(localView, 6, 0x0800);
      writeUint16(localView, 8, 0);
      writeUint16(localView, 10, now.time);
      writeUint16(localView, 12, now.date);
      writeUint32(localView, 14, checksum);
      writeUint32(localView, 18, dataBytes.length);
      writeUint32(localView, 22, dataBytes.length);
      writeUint16(localView, 26, nameBytes.length);
      writeUint16(localView, 28, 0);
      localParts.push(localHeader, nameBytes, dataBytes);

      const centralHeader = new Uint8Array(46);
      const centralView = new DataView(centralHeader.buffer);
      writeUint32(centralView, 0, 0x02014b50);
      writeUint16(centralView, 4, 20);
      writeUint16(centralView, 6, 20);
      writeUint16(centralView, 8, 0x0800);
      writeUint16(centralView, 10, 0);
      writeUint16(centralView, 12, now.time);
      writeUint16(centralView, 14, now.date);
      writeUint32(centralView, 16, checksum);
      writeUint32(centralView, 20, dataBytes.length);
      writeUint32(centralView, 24, dataBytes.length);
      writeUint16(centralView, 28, nameBytes.length);
      writeUint16(centralView, 30, 0);
      writeUint16(centralView, 32, 0);
      writeUint16(centralView, 34, 0);
      writeUint16(centralView, 36, 0);
      writeUint32(centralView, 38, 0);
      writeUint32(centralView, 42, offset);
      centralParts.push(centralHeader, nameBytes);
      offset += localHeader.length + nameBytes.length + dataBytes.length;
    });

    const localBytes = concatBytes(localParts);
    const centralBytes = concatBytes(centralParts);
    const end = new Uint8Array(22);
    const endView = new DataView(end.buffer);
    writeUint32(endView, 0, 0x06054b50);
    writeUint16(endView, 4, 0);
    writeUint16(endView, 6, 0);
    writeUint16(endView, 8, entries.length);
    writeUint16(endView, 10, entries.length);
    writeUint32(endView, 12, centralBytes.length);
    writeUint32(endView, 16, localBytes.length);
    writeUint16(endView, 20, 0);
    return concatBytes([localBytes, centralBytes, end]);
  }

  // 组装一个最小但完整的 XLSX 包所需的全部 XML 文件和关系文件。
  function packageEntries(sheetName, rows, headerRow) {
    const safeSheetName = xmlEscape(sheetName || "BOM");
    return [
      {
        name: "[Content_Types].xml",
        content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
      },
      {
        name: "_rels/.rels",
        content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      },
      {
        name: "xl/workbook.xml",
        content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${safeSheetName}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      },
      {
        name: "xl/_rels/workbook.xml.rels",
        content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
      },
      {
        name: "xl/styles.xml",
        content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
      },
      {
        name: "xl/worksheets/sheet1.xml",
        content: buildSheetXml(rows, headerRow),
      },
    ];
  }

  // 对外入口：生成 BOM XLSX，优先返回 File，旧环境降级返回 Blob。
  function buildBomFile(rows, fileName, platform) {
    const options = platform || {};
    const headerRow = Number(options.headerRow) || 5;
    const sheetName = options.sheetName || "BOM";
    const fileBytes = buildZip(packageEntries(sheetName, rows || [], headerRow));
    const blob = new Blob([fileBytes], { type: MIME_TYPE });
    const name = fileName || "wps-bom-import.xlsx";

    if (typeof File === "function") {
      return new File([blob], name, { type: MIME_TYPE });
    }
    // 极少数旧 WebView 没有 File 构造函数，用带 name 属性的 Blob 兼容。
    Object.defineProperty(blob, "name", { value: name });
    return blob;
  }

  root.XlsxBuilder = {
    buildBomFile,
    buildSheetXml,
  };
})(typeof self !== "undefined" ? self : globalThis);
