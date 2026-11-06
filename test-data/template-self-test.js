// 使用真实模板配置验证旧、新两套字段映射和临时 XLSX 输出。
const fs = require("fs");
const path = require("path");

globalThis.self = globalThis;
const Core = require("../js/core.js");
require("../js/xlsx.js");

// 从浏览器或 Node 的 File/Blob 中读取未压缩 ZIP 条目。
async function readZipEntry(file, entryName) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder("utf-8");
  let offset = 0;

  while (offset + 30 <= bytes.length && view.getUint32(offset, true) === 0x04034b50) {
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const compressedSize = view.getUint32(offset + 18, true);
    const nameStart = offset + 30;
    const dataStart = nameStart + nameLength + extraLength;
    const name = decoder.decode(bytes.slice(nameStart, nameStart + nameLength));
    if (name === entryName) {
      return decoder.decode(bytes.slice(dataStart, dataStart + compressedSize));
    }
    offset = dataStart + compressedSize;
  }
  throw new Error(`ZIP 中未找到 ${entryName}`);
}

// 统一记录失败项，最后一次性输出，避免第一项失败后漏掉后续检查。
function assert(condition, message, failures) {
  if (!condition) failures.push(message);
}

async function run() {
  const failures = [];
  const configPath = path.join(__dirname, "../data/templates.json");
  const config = JSON.parse(fs.readFileSync(configPath, "utf8"));
  const oldTemplate = config.templates.find((item) => item.id === "tpl_bom_import_v1");
  const fullTemplate = config.templates.find((item) => item.id === "tpl_bom_full_v1");
  const oldFields = Core.flattenTemplate(oldTemplate);
  const fullFields = Core.flattenTemplate(fullTemplate);

  assert(oldTemplate.platform.excelParseConfigId === "2095456118164029441", "旧模板 ID 不正确", failures);
  assert(fullTemplate.platform.excelParseConfigId === "2118616788984389634", "新模板 ID 不正确", failures);
  assert(oldFields.length === 7, "旧模板字段数不是 7", failures);
  assert(fullFields.length === 12, "新模板字段数不是 12", failures);
  assert(fullFields.filter((field) => field.mappable).length === 10, "新模板可映射字段数不是 10", failures);

  // 旧模板必须识别平台原始规格表头，避免可选字段未匹配后静默留空。
  const oldSourceHeaders = [
    "序号", "名称", "材料", "规格（宽*长*高）（MM）", "数量", "备注", "加工要求",
  ].map((label, index) => ({ label, column: Core.colLetter(index) }));
  const oldAuto = Core.autoMatchHeaders(oldSourceHeaders, oldFields);
  const oldMapping = {};
  oldAuto.matches.forEach((match) => {
    if (match.fieldId) oldMapping[match.fieldId] = match.headerIndex;
  });
  const oldResolved = Core.resolveFieldMapping(oldSourceHeaders, oldFields, oldMapping);
  assert(oldResolved.valid, "旧模板字段映射未通过", failures);
  assert(oldResolved.assignedCount === 7, "旧模板 7 个字段未全部自动匹配", failures);

  const sourceHeaders = [
    "零件编号", "零件名称", "材质", "规格", "重量", "零件分类", "物料属性",
    "数量", "备注", "加工要求", "热处理", "零件图",
  ].map((label, index) => ({ label, column: Core.colLetter(index) }));
  const auto = Core.autoMatchHeaders(sourceHeaders, fullFields);
  const mapping = {};
  auto.matches.forEach((match) => {
    if (match.fieldId) mapping[match.fieldId] = match.headerIndex;
  });
  const resolved = Core.resolveFieldMapping(sourceHeaders, fullFields, mapping);
  assert(resolved.valid, "新模板字段映射未通过", failures);
  assert(
    resolved.assignedCount === 10,
    `新模板自动映射数量不是 10，实际为 ${resolved.assignedCount}；未匹配：` +
      resolved.unmappedFields.map((item) => item.field.label).join("、"),
    failures
  );

  const mappedRows = Core.buildMappedRows(resolved.fieldMappings, [[
    "P-001", "测试件", "45#", "100x50", "2.5", "机加工", "",
    "3", "测试备注", "精加工", "调质", "",
  ]]);
  const validation = Core.validateImportRows(mappedRows, 2);
  assert(
    validation.errors.length === 0,
    `新模板数据校验失败：${validation.errors.join("；")}`,
    failures
  );
  assert(validation.rows.length === 1, "新模板校验后没有有效数据行", failures);
  if (validation.rows.length === 0) {
    console.error(failures.join("\n"));
    process.exitCode = 1;
    return;
  }
  assert(validation.rows[0].partWeight === 2.5, "重量未转换为数字", failures);
  assert(validation.rows[0].partClassify === "机加工", "零件分类未保留", failures);
  assert(validation.rows[0].partHotDispose === "调质", "热处理未保留", failures);
  assert(validation.rows[0].materialAttribute === null, "物料属性未保持为空", failures);
  assert(validation.rows[0].partImage === null, "零件图未保持为空", failures);

  // 零件编号为空不应阻断临时文件生成，但必须产生明确警告。
  const blankPartNoRows = Core.buildMappedRows(resolved.fieldMappings, [[
    "", "无编号测试件", "45#", "M8x20", "0.2", "标准件", "",
    "8", "", "", "", "",
  ]]);
  const blankPartNoValidation = Core.validateImportRows(blankPartNoRows, 2);
  assert(blankPartNoValidation.errors.length === 0, "空零件编号不应阻止生成", failures);
  assert(blankPartNoValidation.rows.length === 1, "空零件编号行被错误丢弃", failures);
  assert(blankPartNoValidation.warnings.length === 1, "空零件编号没有产生警告", failures);
  const blankPartNoFile = globalThis.XlsxBuilder.buildBomFile(
    blankPartNoValidation.rows,
    "blank-part-no.xlsx",
    fullTemplate.platform,
    fullFields
  );
  const blankPartNoSheet = await readZipEntry(blankPartNoFile, "xl/worksheets/sheet1.xml");
  assert(blankPartNoSheet.includes('<c r="A2"/>'), "空零件编号没有生成空单元格", failures);
  assert(blankPartNoSheet.includes("无编号测试件"), "空零件编号行未写入临时文件", failures);

  // 相同零件编号但规格不同的行是真实 BOM 常见数据，不应被前端单独按编号拦截。
  const duplicatePartNoRows = Core.buildMappedRows(resolved.fieldMappings, [
    ["P-DUP", "重复编号件A", "45#", "M8x20", "0.2", "标准件", "", "1", "", "", "", ""],
    ["P-DUP", "重复编号件B", "45#", "M8x30", "0.3", "标准件", "", "2", "", "", "", ""],
  ]);
  const duplicatePartNoValidation = Core.validateImportRows(duplicatePartNoRows, 2);
  assert(duplicatePartNoValidation.errors.length === 0, "相同零件编号不应被前端拦截", failures);
  assert(duplicatePartNoValidation.rows.length === 2, "相同零件编号的数据行被错误丢弃", failures);

  const fullFile = globalThis.XlsxBuilder.buildBomFile(
    validation.rows,
    "full.xlsx",
    fullTemplate.platform,
    fullFields
  );
  const fullSheet = await readZipEntry(fullFile, "xl/worksheets/sheet1.xml");
  assert(fullSheet.includes('dimension ref="A1:L2"'), "新模板范围不是 A1:L2", failures);
  assert(fullSheet.includes('r="A1"') && fullSheet.includes("零件编号"), "新模板第 1 行表头缺失", failures);
  assert(fullSheet.includes('r="L1"') && fullSheet.includes("零件图"), "新模板第 12 列表头缺失", failures);
  assert(fullSheet.includes('<c r="G2"/>'), "物料属性不是空单元格", failures);
  assert(fullSheet.includes('<c r="L2"/>'), "零件图不是空单元格", failures);
  assert(fullSheet.includes('<c r="E2"><v>2.5</v></c>'), "重量不是数字单元格", failures);
  assert(fullSheet.includes("调质"), "热处理未写入临时文件", failures);

  // 用 12 行数据确认插件自身不会在第 10 行截断。
  const manyRows = Array.from({ length: 12 }, (_, index) => ({
    ...validation.rows[0],
    partNo: `P-${String(index + 1).padStart(3, "0")}`,
    partName: `测试件${index + 1}`,
  }));
  const manyFile = globalThis.XlsxBuilder.buildBomFile(
    manyRows,
    "many.xlsx",
    fullTemplate.platform,
    fullFields
  );
  const manySheet = await readZipEntry(manyFile, "xl/worksheets/sheet1.xml");
  assert(manySheet.includes('dimension ref="A1:L13"'), "12 行数据未完整输出到第 13 行", failures);
  assert(manySheet.includes("P-012"), "第 12 行数据被截断", failures);

  const oldFile = globalThis.XlsxBuilder.buildBomFile(
    [{ partNo: "P-OLD", partName: "旧件", partNum: 1 }],
    "old.xlsx",
    oldTemplate.platform,
    oldFields
  );
  const oldSheet = await readZipEntry(oldFile, "xl/worksheets/sheet1.xml");
  assert(oldSheet.includes('dimension ref="A5:G6"'), "旧模板范围不是 A5:G6", failures);
  assert(oldSheet.includes('r="A5"') && oldSheet.includes("序号"), "旧模板第 5 行表头缺失", failures);
  assert(oldSheet.includes('r="A6"'), "旧模板数据未从第 6 行开始", failures);

  // 无参数兜底调用仍应保持旧版 7 字段、第 5 行表头的兼容行为。
  const fallbackFile = globalThis.XlsxBuilder.buildBomFile([{ partNo: "P-FALLBACK", partName: "兜底件" }]);
  const fallbackSheet = await readZipEntry(fallbackFile, "xl/worksheets/sheet1.xml");
  assert(fallbackSheet.includes('dimension ref="A5:G6"'), "旧版兜底范围不是 A5:G6", failures);

  if (failures.length > 0) {
    console.error(`FAIL\n${failures.join("\n")}`);
    process.exitCode = 1;
    return;
  }
  console.log("PASS: 旧模板和新模板的字段映射、表头行、数据行及空列均正确。");
}

run().catch((error) => {
  console.error(error.stack || error.message);
  process.exitCode = 1;
});
