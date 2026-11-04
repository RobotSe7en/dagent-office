# R3-03 工作簿包与只读模型

`core` 已实现工作簿严格契约、新建、分块校验、类型化单元格/窗口查询及保存重开，`office` 文档包 ZIP 同样支持工作簿。原生 XLSX 读取、完整公式语法与求值、网格预览和编辑仍在后续任务；不能把本阶段称为 Excel 完整支持。完整范围见[第三轮细则](iteration-3.md)，存储依据见[技术基线](workbook-storage.md)。

## 统一入口与类型分支

按用户选择，三种格式统一使用 `openDocument`。返回模型的 `kind` 是类型判别字段：Word/PPT 使用 `ContentDocumentModel` 和原有 HTML/节点接口，工作簿使用 `WorkbookModel`、`WorkbookSnapshot`、`getCell` 和 `readWindow`。工作簿快照不含 `html`，清单不含 `entrypoint`；查询不先展开全表。现有 Word/PPT 包继续可读，宿主 TypeScript 代码需要在使用格式专属字段前判断类型。

```ts
import { createWorkbook, openDocument } from "@dagent-office/core";

const files = createWorkbook({ documentId: "host-authorized-id", sheetName: "预算" });
const model = openDocument(files);
if (model.kind === "workbook") {
  const sheetId = model.snapshot().metadata.workbook.activeSheetId;
  const cells = model.readWindow({ sheetId, row: 0, column: 0, rows: 40, columns: 20 });
  // 新建工作簿没有实际单元格，cells 为空；未存储的坐标是隐式空白。
  model.getCell(sheetId, 0, 0);
  openDocument(model.serialize(), { kind: "workbook" });
}
```

可选 `kind` 参数是预期格式断言，不做转换；不匹配即失败，并帮助 TypeScript 收窄返回类型。`createDocumentSession` 当前明确拒绝工作簿，直到 R3-08 读取门槛之后开放编辑。Playground 编辑器和 Node 转换目前也拒绝工作簿，不把它当成空白 Word。

## 文件与来源

| 内容 | 来源及校验 |
| --- | --- |
| 清单 | `kind: workbook`，`chunks` 数组；每块包含 sheetId、零基准 row/column、路径和 SHA-256 |
| 元数据 | workbook 的日期系统、计算模式、活动表、表次序/名称/可见性、行列度量、冻结、合并和数字格式 |
| 样式 | 清单列出的 CSS；单元格和富文本的 class 引用必须存在；拉丁/东亚字体角色分别声明 |
| 内容 | 每块 HTML 表格；行、列按坐标递增，原点为 256 行及 32 列的整数倍；块不能重叠或引用不存在的表 |
| 批注/资源 | 当前仅允许空工作簿批注数组；已声明资源仍须校验摘要，完整工作簿批注和绘图在 R3-07 接入 |

空白远端位置不产生任何分块。合并、冻结和行列尺寸只保存逻辑定义；合并可跨块，不随窗口复制。至少一张表可见，活动表必须可见；表身份、大小写不敏感的表名及数字格式 ID 唯一。尺寸区间必须有序且不重叠，合并重叠通过行扫描和列区间索引检查。

每个块是完整的无脚本 HTML 文档，包含一个 table/tbody；以下是其内容示意：

```html
<table data-sheet-id="sheet-1" data-row="0" data-column="0"><tbody>
  <tr data-row="0">
    <td data-column="0" data-value-type="text" class="cell-default">编号 001</td>
    <td data-column="1" data-value-type="number" data-value="12.5" data-format-id="general">12.50</td>
    <td data-column="2" data-value-type="number" data-value="13.5" data-formula="=B1+1" data-cache-state="imported">13.5</td>
  </tr>
</tbody></table>
```

常量类型为 blank/text/number/boolean/error。文本 HTML 是内容源，允许 span 与 br；数字/布尔/错误的原始值在 `data-value`，正文为来源显示缓存。公式表达式是计算源，`data-cache-state` 区分 imported/valid/stale/unavailable；当前阶段只解析并保留声明语法和缓存，不声称已重新求值。新建默认自动计算和 1900 日期系统，也可显式选 1904。

未建模 HTML、未知属性、未知样式或格式、重复坐标、错误摘要、无效 UTF-8、非法标量值及不支持的公式语法均失败。ZIP 打包转发同样的工作簿限额。快照和查询结果冻结，输入及序列化字节隔离，索引不保留全表解析树或第二份单元格 JSON。

## 预算与验证

默认文档字节限制保持原值；`workbookLimits` 单独限制实际单元格 250,000、公式 25,000、一次窗口面积 5,000、模板 25,000 和模板键 4,194,304 个 UTF-16 单位。这些是保护限额，不是容量验收结果。一万行基线为 200,000 实际单元格、20,000 公式。

文档的 `maxNodes`/`maxDepth` 对工作簿每个有界块及元数据适用，完整工作簿另外累计检查实际单元格与公式；源块最大嵌套深度为 16。跨块不能重置整包字节、单元格、公式和模板预算。返回的窗口按行列排序，只包含实际存储的单元格，缺失位置由调用者作为空白显示。

验证覆盖工作簿/原有模型的类型分支、严格 JSON Schema、摘要和跨块身份、中文及变体选择符、类型与公式缓存、稀疏最远坐标、样式引用、合并/维度约束、窗口和累计额度、不可变字节、文档包 ZIP 重开，以及工作区之外的打包 JS/浏览器/Node 类型消费。Word/PPT 原有模型、事务和结构测试继续运行。

```sh
pnpm test packages/core/tests/workbook-model.test.ts packages/office/tests/workbook-package.test.ts
pnpm test tests/package-consumption.test.ts
pnpm probe:workbook-model
```

公开入口的一万行[本地测量](evidence/workbook-model-baseline.json)覆盖 20 万实际单元格和 2 万公式：HTML 包打开约 1.9 秒，通过受限窗口扫描全部单元格约 9.0 秒，快照仅 7,215 字节，进程生涯峰值约 444 MiB。该记录没有宣称原生 XLSX、计算引擎或最终固定环境性能已通过。
