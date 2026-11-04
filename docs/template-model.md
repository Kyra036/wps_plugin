# BOM 模板数据结构

模板定义位于 `data/templates.json`，示例工作表位于 `data/mock.json`。
两个文件都必须保持合法 JSON，因此不能在文件内写 `//` 注释。
本文件用于说明模板库和字段结构的含义。

## 模板库

```json
{
  "version": "1.0",
  "defaultTemplateId": "tpl_bom_standard_v1",
  "templates": []
}
```

- `defaultTemplateId`：推荐的内置模板标识。
- 页面当前不会自动选中该模板，用户仍从“请选择模板”开始操作。
- `templates`：内置模板列表，后续可替换为服务端模板接口。

## 模板

```json
{
  "id": "tpl_bom_standard_v1",
  "name": "BOM 标准模板（示例）",
  "builtIn": true,
  "version": 1,
  "groups": []
}
```

- `id`：模板唯一标识。
- `name`：页面显示名称。
- `builtIn`：是否为系统内置模板。
- `version`：模板结构版本，用于后续判断模板是否需要升级。
- `groups`：表头分组，顺序决定字段在工作表中的默认列顺序。

## 分组

```json
{
  "key": "product",
  "label": "产品定义",
  "visible": true,
  "builtIn": true,
  "fields": []
}
```

- `key`：分组唯一标识，同时作为字段路径的第一段。
- `label`：页面显示名称，例如“产品定义”。
- `visible=false`：整组隐藏，展开模板时不会生成字段。
- `builtIn`：是否为内置分组。
- `fields`：分组内的字段列表。

## 字段

```json
{
  "key": "productCode",
  "label": "产品编号",
  "visible": true,
  "builtIn": true,
  "aliases": ["物料编码", "料号"]
}
```

- `key`：字段唯一标识，与分组 key 组合成完整字段路径。
- `label`：当前模板表头显示名称。
- `visible=false`：字段不进入字段列表，也不参与上传。
- `builtIn=false`：预留给后续自定义字段。
- `aliases`：自动识别客户表头时使用的候选名称，当前步骤只保存，尚未参与匹配。

## 当前默认分组

- `product`：产品定义，7 个字段。
- `bom`：物料清单，6 个字段。
- `route`：工艺路线，2 个字段。

当前 BOM 模板不包含时间相关字段。
