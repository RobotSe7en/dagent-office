# R3-03 工作簿包与只读模型

`core` 已实现工作簿严格契约、新建、分块校验、类型化单元格/窗口查询及保存重开，`office` 文档包 ZIP 同样支持工作簿。[公式语法与引用基础](workbook-formulas.md)已实现，原生 XLSX 完整读取、求值、网格预览和编辑仍在后续任务；不能把本阶段称为 Excel 完整支持。完整范围见[第三轮细则](iteration-3.md)，存储依据见[技术基线](workbook-storage.md)。

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
| 元数据 | workbook 的日期系统、计算模式、活动表、表次序/名称/可见性、行列度量、冻结、合并、数字格式、作用域名称、超链接和打印定义 |
| 样式 | 清单列出的 CSS；单元格和富文本的 class 引用必须存在；拉丁/东亚字体角色分别声明 |
| 内容 | 每块 HTML 表格；行、列按坐标递增，原点为 256 行及 32 列的整数倍；块不能重叠或引用不存在的表 |
| 批注/资源 | 当前仅允许空工作簿批注数组；已声明资源仍须校验摘要，完整工作簿批注和绘图在 R3-07 接入 |

空白远端位置不产生任何分块。合并、冻结和行列尺寸只保存逻辑定义；合并可跨块，不随窗口复制。至少一张表可见，活动表必须可见；表身份、大小写不敏感的表名及数字格式 ID 唯一。尺寸区间必须有序且不重叠，合并重叠通过行扫描和列区间索引检查。

R3-06 接入原生读取时增加了以下可选契约，旧工作簿包无需补字段：

- `names` 保存不透明身份、名称、可选工作表作用域、表达式及明确的引用原点。名称在各自作用域内不区分大小写，不能与单元格地址或保留名称冲突；普通表达式的解析不代表已支持求值。
- 表及行列维度可引用 `numberFormatId`，避免格式化的空行列丢失数字格式。`defaultColumnWidthUnit` 区分原生列宽与不含边距的基础字符宽；缺省沿用原有字符宽语义。`defaultRowsHidden` 表示隐式行的隐藏默认值。
- 数字格式定义为自定义 `code` 或 `builtinId` 二选一，保留内建日期格式随语言变化的语义，不在导入时固定成某个地区的格式串。`themeCssClass` 与 `tabCssClass` 引用 CSS 中的主题和工作表标签外观；既有 `tabColor` 继续可读，但不能与 `tabCssClass` 同时定义。
- `hyperlinks` 保存单元格区域及 HTTP(S)/邮件目标或内部名称/地址，可保留原生说明文字。读取不访问目标，重叠链接及可执行协议失败。
- `print` 保存区域、重复标题、纸张度量、方向、页边距、适配或百分比缩放、分页及类型化页眉页脚文字/字段。未指定的纸张及页边距保持未指定，未提前猜测打印机设置。该契约可保存重开，实际工作簿分页/PDF 在 R3-18 验收。

原生适配的分部件检查覆盖 32 份独立来源及 Calc 重存文件的名称和打印设置、流式行列/冻结/合并/超链接、共享/内联富文本、公式缓存与共享公式。自动行高是原生布局缓存，固定行高才写入源尺寸。当前这些私有读取模块尚未组成公开 `importXlsx`；样式、完整扩展部件和 R3-08 读取验收仍未完成。

原生样式适配已将 32 份样例的有效单元格字体、主题/索引颜色、纯色填充、边框、对齐及数字格式映射到可校验的 CSS/格式定义。主题引用、字体角色、原字体名及明暗参数保留为来源；色阶应用、精确边框和特殊对齐需要 R3-08 渲染器解析。原生命名样式定义及更新标记目前只在导入中间态保留，最终持久化取舍仍待产品决策，不能称为完整样式保真已通过。

`cellXfs` 的 `apply*` 标记控制后续命名样式修改是否覆盖单元格，而不是简单地关闭该格当前格式；实现据 [Microsoft 的 XF 说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/68362a4b-5589-4504-b566-e8154dce1de3)保留当前值与更新标记。索引色采用[标准映射](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.indexedcolors?view=openxml-2.8.1)，工作簿明暗参数保留 [HLS 亮度语义](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.foregroundcolor?view=openxml-3.0.1)。

原生一万行分部件链路已执行两遍有界扫描、共享公式处理、40 块 HTML 生成及统一模型重开，核对 200,000 个单元格、20,000 个公式和末行结果。它不依赖技术基线 HTML 生成器，也尚未包含完整公开导入、视觉或最终固定环境验收。

原生记录管线现在贯通异步 ZIP、全包关系审计、样式/共享字符串、逐表双遍扫描、维度压缩和 HTML 分块。32 份独立文件及 Calc 重存文件通过这条内部链路；两份一万行来源均保留 200,000 个单元格和 20,000 个公式。XML 全局节点、单元格、公式与输出字节跨表累计，元数据另设保留节点限制；取消、宿主进度回调异常及输入字节隔离有独立检查。管线保留尚未消费的绘图、智能表和批注头部供后续适配，尚不对外返回一个缺少这些内容的文档包。

`conditionalRules` 已有数值比较、标量公式和双色/三色色阶来源；差异外观及色阶颜色引用 CSS，阈值与优先级保留为元数据。`dataValidations` 保存列表、整数、小数、日期、时间、文本长度和自定义公式，也允许仅有提示的 `none`。规则都保存精确多区域和公式原点；严格检查身份、公式数量、样式引用、优先级及有歧义的重叠验证。已有违规数据不会在读取阶段被修正；规则求值、违规诊断和各编辑入口执行仍未实现。

原生规则验证依据 [SpreadsheetML 条件格式说明](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting)及[数据验证契约](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidation?view=openxml-2.8.1)，并单独保存[固定 Calc 的探针结果](evidence/xlsx-rule-semantics.json)：传统 `showDropDown="1"` 隐藏列表；多区域公式原点取最左列中最靠上的实际区域起点，不取列表首项或外包矩形角。后者在跨行列样例中分别为 B5/E9，并经外部 ODS 中的引用重定位核对。这是固定 Calc 的独立观测，不冒充 Excel/WPS 实机结论。运行 `python3 scripts/check-xlsx-rule-semantics.py` 重复核对；只有显式 `--record` 更新记录。

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
