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
  model.resolveReference("=A1:B20", { sheetId, row: 0, column: 0 });
  openDocument(model.serialize(), { kind: "workbook" });
}
```

可选 `kind` 参数是预期格式断言，不做转换；不匹配即失败，并帮助 TypeScript 收窄返回类型。`createDocumentSession` 当前明确拒绝工作簿，直到 R3-08 读取门槛之后开放编辑。Playground 编辑器和 Node 转换目前也拒绝工作簿，不把它当成空白 Word。

`resolveReference` 根据本模型的工作表、名称和智能表，返回有序源区域、引用错误或明确的未解析原因；不展开单元格、不求值，也不使用图表缓存。返回的联合区域保留顺序和重复覆盖，调用者不能直接当作集合去重。具体支持和预算见[引用基础](workbook-formulas.md)。

只读预览的内部窗口读取器已按文档与修订绑定完整显示帧，跨窗格合并只读取一次内容锚点。外观查询还包含合并四角、可见行列对应的周界及最近可见邻格；邻格属于另一合并时也读取其四角。整张表大小的合并不展开内部坐标，隐藏的百万行整段跳过。可见单元格默认上限仍为 5,000，外观来源另限 25,000 个地址、累计源文本 8 Mi UTF-16 单位；这些是保护限额，不是容量验收结果。整批查询先检查累计单元格和重复地址，再以最多 256 个坐标的小窗口读取；任务间让出执行以接收取消。快速切换窗口、旧响应、超限及已释放实例不会发布部分结果。内部 DOM 投影保持完整单元格尺寸，再按窗格裁剪，RTL 只改变物理位置；合并内容不会在碎片内重新居中，重复片段只有一个辅助技术单元格身份。绘制失败时保留前一个完整窗口，派生 DOM 不写回源包。

`tests/browser/workbook-grid.spec.ts` 在 Chromium、Firefox 和 WebKit 的几何检查覆盖左右方向、四窗格合并、源命中、第 1 万行与 XFD1048576、DOM 单元格预算和释放。这里使用少量稀疏实际单元格，验证坐标与窗口上限，不能替代 20 万实际单元格的完整 XLSX 首屏验收。此内部适配尚未作为公开 Worker 协议或完整网格预览发布，数字格式、字体和对象绘制仍需接入。

## 文件与来源

| 内容 | 来源及校验 |
| --- | --- |
| 清单 | `kind: workbook`，`chunks` 数组；每块包含 sheetId、零基准 row/column、路径和 SHA-256 |
| 元数据 | workbook 的日期系统、计算模式、活动表、表次序/名称/可见性、行列度量、冻结、合并、数字格式、作用域名称、超链接和打印定义 |
| 样式 | 清单列出的 CSS；单元格和富文本的 class 引用必须存在；拉丁/东亚字体角色分别声明 |
| 内容 | 每块 HTML 表格；行、列按坐标递增，原点为 256 行及 32 列的整数倍；块不能重叠或引用不存在的表 |
| 批注/资源 | `cells` 锚点保留工作表与精确多区域；正文/回复引用有界独立 HTML story，资源校验摘要；原生备注/线程、基础绘图与图表已接入内部管线，兼容定位核对及预览仍待完成 |

空白远端位置不产生任何分块。合并、冻结和行列尺寸只保存逻辑定义；合并可跨块，不随窗口复制。至少一张表可见，活动表必须可见；表身份、大小写不敏感的表名及数字格式 ID 唯一。尺寸区间必须有序且不重叠，合并重叠通过行扫描和列区间索引检查。

R3-06 接入原生读取时增加了以下可选契约，旧工作簿包无需补字段：

- `names` 保存不透明身份、名称、可选工作表作用域、表达式及明确的引用原点。名称在各自作用域内不区分大小写，不能与单元格地址或保留名称冲突；普通表达式的解析不代表已支持求值。
- 表及行列维度可引用 `numberFormatId`，避免格式化的空行列丢失数字格式。`defaultColumnWidthUnit` 区分原生列宽与不含边距的基础字符宽；缺省沿用原有字符宽语义。`defaultRowsHidden` 表示隐式行的隐藏默认值。
- 数字格式定义为自定义 `code` 或 `builtinId` 二选一，保留内建日期格式随语言变化的语义，不在导入时固定成某个地区的格式串。`themeCssClass` 与 `tabCssClass` 引用 CSS 中的主题和工作表标签外观；既有 `tabColor` 继续可读，但不能与 `tabCssClass` 同时定义。
- `columnWidthFontCssClass` 单独引用字符列宽的计量字体；原生源按 `cellStyle builtinId=0 → cellStyleXfs → font` 找到 Normal 字体，不以字体记录零或 `cellXfs` 零代替。依据 [Microsoft Open Specifications 的说明](https://learn.microsoft.com/en-us/answers/questions/1110580/where-is-the-normal-style-font-defined-in-an-xlsx)，本地化名称和修改过的 Normal 样式仍通过内建身份识别。原生 Normal 定义有歧义或缺失时不猜测列宽字体；老 HTML 包没有此可选引用时沿用表的默认字体。这个必要的布局引用不替代尚待决策的完整命名单元格样式保留。
- `hyperlinks` 保存单元格区域及 HTTP(S)/邮件目标或内部名称/地址，可保留原生说明文字。读取不访问目标，重叠链接及可执行协议失败。
- `print` 保存区域、重复标题、纸张度量、方向、页边距、适配或百分比缩放、分页及类型化页眉页脚文字/字段。未指定的纸张及页边距保持未指定，未提前猜测打印机设置。该契约可保存重开，实际工作簿分页/PDF 在 R3-18 验收。

原生适配的分部件检查覆盖 32 份独立来源及 Calc 重存文件的名称和打印设置、流式行列/冻结/合并/超链接、共享/内联富文本、公式缓存与共享公式。自动行高是原生布局缓存，固定行高才写入源尺寸。当前这些私有读取模块尚未组成公开 `importXlsx`；样式、完整扩展部件和 R3-08 读取验收仍未完成。

数字格式的内部源码编译器已按照[格式代码规则](https://support.microsoft.com/en-us/excel/review-guidelines-for-customizing-a-number-format)区分分段、引号/转义、条件/颜色、日期/时长字段、科学计数、货币/地区标记及需要测量的填充/留白字符；引号内的分号不会被错误拆段。十进制、科学计数和常用分数分段已能生成文本、按字符宽度留白及填充部件；分组与小数符号由调用方显式提供。显示舍入另行处理，百分比、逗号缩放和科学计数不先乘除浮点值，避免中间溢出；它不改变原值，也不是公式 `ROUND` 的实现。分数使用有界连分数近似，支持假分数、带分数及不约分的固定分母，位数增加不会逐个枚举分母；原生分母常量与引号中的文本数字分别解释。带分数的强制零分子、分组整数等特殊组合仍明确拒绝，不能默默套用其他格式。[固定 Calc 的 130 个格式化结果](evidence/xlsx-number-format-semantics.json)可通过 `python3 scripts/check-xlsx-number-formats.py` 重核，其中 60 个分数结果已与内部显示部件对照；这不代表实际字体宽度与视觉验收。此证据记录了 1900 系统早期日期与 Excel 的差异，SDK 日期语义仍待用户选择。当前尚无完整数字格式入口：日期及公式 TEXT 仍待实现，预览也尚未接入这些基础模块。

基本分段选择和文本段已接入内部标量格式入口：先用原始值选择正数/负数/零或条件段，再进行显示舍入；提前出现的末尾 `@` 仍是文本段。文本不因外观格式转换为数字，空段保留隐藏含义，纯字面内容不额外加负号。[47 个独立分段/文本结果](evidence/xlsx-format-selection.json)由 `python3 scripts/check-xlsx-format-selection.py` 重核。两个格式参考共 177 项，其中 151 项数值/文本结果已对照；另外 26 项日期或 General 由独立的日历/宽度分支处理，尚未计入完成的全入口格式验收，不把原生缓存字符串当成重算结果。未匹配的条件不自动冒充 General，部分混合显式/隐式条件组合仍待核对。

General 的内部宽度适配已按 [Excel General 说明](https://support.microsoft.com/en-us/excel/reset-a-number-to-the-general-format)和 [OOXML 数字格式规则](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1)实现有界候选与实测选择：普通十进制最多 11 个显示位置，12 位及以上整数使用科学计数，窄列可以减少小数显示精度，不能把小非零值当作零。极端有限数也不先乘浮点数缩放。无可用显示时返回宽度溢出，保留原始标量供宿主查询。此规则与固定 Calc 的部分宽列 General 文本不同，不将 Calc 的原文作为该分支的 Excel 预期。显示部件的留白按实际字符测量，填充字符串整体参与字形测量，并受累计输出限额约束；相邻文字合并测量以保留字距。完整单元格字体、内建地区格式和日期仍需继续接入。

列宽计量已可在样式与本地字体就绪后解析上述 Normal 引用，测量 0–9 的最大数字宽度并返回字体替代诊断。字体解析只作用于所属表面的临时测量节点，不改源 CSS 或其他已挂载界面；成功或失败都会释放临时节点。三浏览器已用独立 DOM 字宽核对该数值，并覆盖旧包默认字体路径。完整原生列宽、自动行高和固定 Office 参考的最终视觉门槛仍未完成。

原生表的默认外观及数字格式同样来自 [Normal 样式](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyle?view=openxml-3.0.1)，不以 `cellXfs[0]` 替代。内部预览按单元格、有效行、列、表默认的顺序选择外观和数字格式引用；两类引用分别回退，不把 CSS 类列表顺序当作级联优先级。已存储但没有显式格式的空白与隐式空白采用同一回退规则，显式格式零则保留覆盖含义。行的原生 `customFormat` 未启用时，不应用该行的 `s`；启用但省略 `s` 时应用原生缺省的格式零。稀疏维度通过二分查找定位，不展开百万行。[54 个独立 Calc 结果](evidence/xlsx-cell-style-semantics.json)覆盖这些区别，包含 Normal 与格式零不同的数值显示和背景色；可用 `python3 scripts/check-xlsx-cell-styles.py` 重核。此项不代表完整条件格式或最终视觉验收。

窗口样式测量已保留源块的 `table/tbody/tr/td` 祖先、坐标属性和富文本 `span/br`，因此基于源结构的 CSS 仍按浏览器级联匹配，不受虚拟网格裁剪容器影响。跨窗格的同一锚点只建立一份测量源；临时节点按窗口累计限额构建，全部校验完成后才连接到测量表面，释放时清空节点及索引。惰性 HTML 解析后的结构和属性白名单会拒绝脚本、事件、资源标签及与查询不符的单元格，失败不留下部分测量树。三浏览器已核对远端块选择器、富文本格式、主题变量和释放；完整单元格绘制仍待接入。

单元格字体读取复用 PPT 的字体角色分段，保留 UTF-16 原文和偏移，东亚、拉丁及复杂文字的字体引用分别解析。组合标记和变体选择符沿用前一个基字的角色，跨富文本 span 也不丢失该关联；`br` 计为一个换行单位。分段前先读取原始祖先上的 CSS，不为测量插入会改变源选择器匹配的 span，字体替代和缺字诊断仍由显示表面负责。字体分段受独立限额约束，不将这种字体读取检查等同于最终排版或缺字验收。

内部富文本显示已把这些字体角色、主题明暗色和源 span 样式投影为独立视图，保留每个文本节点对应的 UTF-16 区间。原生普通/双下划线与删除线独立绘制，局部富文本可关闭原生装饰；CSS 装饰仍保留祖先传播关系。`br` 显示为同偏移的换行，源包不被改写。三浏览器的窗口用例已接入真实富文本与纯色填充，并检查冻结合并、远端坐标和有限节点构建。财务下划线尚需行几何，当前明确失败；这些内部检查不代表完整对齐、自动行高、数字单元格及 Office 视觉门槛已通过。

内部文字方向按 [SpreadsheetML readingOrder 规则](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.alignment?view=openxml-3.0.1)区分显式 LTR/RTL 与上下文方向。上下文只检查首个非空白 Unicode 标量，因此数字或 emoji 开头的混排文本不能直接套用整段 HTML `dir=auto`。单字符分类使用浏览器 Unicode 双向表，短缓存只属于当前测量表面，读后清空测量文字、释放时清空缓存。三浏览器覆盖希伯来语、阿拉伯语、Adlam、方向标记、弱字符开头及宿主 RTL 隔离；最终单元格对齐和原生视觉对照仍待接入。

原生样式适配已将 32 份样例的有效单元格字体、主题/索引颜色、纯色填充、边框、对齐及数字格式映射到可校验的 CSS/格式定义。主题引用、字体角色、原字体名及明暗参数保留为来源；色阶应用、精确边框和特殊对齐需要 R3-08 渲染器解析。原生命名样式定义及更新标记目前只在导入中间态保留，最终持久化取舍仍待产品决策，不能称为完整样式保真已通过。

对齐缩进的原生读取范围为 0–255，251–255 不被截断或误拒绝，负数、小数及 256 明确失败；依据为 [MS-OI29500 v20260818 §2.1.693](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OI29500/%5BMS-OI29500%5D.pdf#page=295)。实际缩进使用单独测量的 Normal 字体空格宽度，不随该单元格的字体改变。

内部文字布局已接入 General／左中右、上中下、Normal 空格缩进、换行和单行缩小适应，并在同窗口批量解析本地字体后测量。富文本装饰和 UTF-16 投影保留；测量成功后才交给完整单元格的裁剪容器，失败和释放不残留测量文字，也不修改源。它返回未裁剪的内容高度，尚未组成自动行高流程。窗口测试已使用这条布局路径；跨空白格溢出、特殊对齐、完整数字／日期显示和最终只读预览仍待实现。

[94 项独立原生 PDF 记录](evidence/xlsx-alignment-semantics.json)可用 `python3 scripts/check-xlsx-alignment.py` 重核，覆盖固定的三种字体、General 类型、显式方向、缩进及横纵对齐。三浏览器对照字宽和物理位置，并检查换行、缩小、字体替代与失败清理；使用关闭字偶距的原生文字度量，WebKit 同时需要显式关闭 OpenType `kern`。这些局部校准不等同于完整页面视觉验收。Calc 的布尔／错误值右对齐与 [OOXML General 居中规则](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1)不同，SDK 遵循后者；Calc 的显式 LTR 希伯来文字和 RTL PDF 边框相对文字位置也单独记录，不能当作 Excel 实测。数值负号和错误值标点按类型显示方向保持顺序，普通文本仍遵循显式阅读方向。日期和默认地区语义没有由此推定。

内部颜色显示已按 [SpreadsheetML 的 HLS 亮度规则](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.foregroundcolor?view=openxml-3.0.1)处理正负 tint，区别于 DrawingML 的 RGB tint/shade。浏览器解析源 CSS 中的主题/索引变量及当前文字颜色，保留源引用；原生 ARGB 高字节不作为单元格透明度。无效颜色或明暗参数明确失败。[108 个固定 Calc 颜色观测](evidence/xlsx-color-tint-semantics.json)与连续 HLS 转换的差别最多为一个 8 位通道单位，记录了舍入边界，不声称逐字节颜色一致；运行 `python3 scripts/check-xlsx-color-tints.py` 可复核。三浏览器另核对变量、颜色语法、透明度和无效值处理。

普通共享边框的内部选择器已对照[固定 Calc 的 14 类边框记录及 588 个 PDF 边界颜色结果](evidence/xlsx-border-semantics.json)：线宽等级优先，双线排在中等与粗线之间，同等级取物理下方/右方，因此 RTL 列轴会反转源顺序中的平局选择。两侧均无边框时不产生描边，不套用浏览器 CSS 表格的冲突规则。可运行 `python3 scripts/check-xlsx-borders.py` 重核；字体清单摘要、Office 归档及 PDF 工具环境随记录保存。独立描边层已按 PDF 的点单位线宽、虚线节距和双线间隔核对，三浏览器另检查裁剪后虚线实线区域完全对应、抗锯齿透明度最多一个通道单位的变化及双线留空。它只接收已选择的物理边，条件格式、合并周界及最终网格组合仍未接入。Calc 将 `slantDashDot` 变为细虚线，这是外部生产者差异，不能把该输出作为斜线保真的证明，当前描边层明确拒绝该项；Excel/WPS 仍未实机核对。

[OOXML Primer §3.2.13](https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf)将合并内容与各格保存的边框格式区别说明。[80 个合并边框原生 PDF 探针](evidence/xlsx-merged-border-semantics.json)进一步记录了固定 Calc 对不一致周界格式的处理，以及外侧较粗邻格边框对局部边界的影响；运行 `python3 scripts/check-xlsx-merged-borders.py` 可复核。RTL 导出可能重复同一描边，记录保留了原始重复项。这些观测不替代 Excel/WPS 实测，也尚未定型 SDK 的完整合并边框组合；窗口已提供必要来源，未将观察到的生产者选择写回文档源。

DrawingML 主题的填充、背景填充、线条和效果列表也已映射为共享 CSS 定义，保留占位颜色与对象引用颜色的关系。复用 PPT 的颜色、填充和轮廓原语，Excel 边界额外检查属性、重复选择、颜色变换顺序及累计样式预算。非空对象默认定义、附加配色及未知扩展目前明确拒绝；本项验证覆盖全部 32 份样式来源，尚不代表绘图对象和预览已完成。

`cellXfs` 的 `apply*` 标记控制后续命名样式修改是否覆盖单元格，而不是简单地关闭该格当前格式；实现据 [Microsoft 的 XF 说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/68362a4b-5589-4504-b566-e8154dce1de3)保留当前值与更新标记。索引色采用[标准映射](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.indexedcolors?view=openxml-2.8.1)，工作簿明暗参数保留 [HLS 亮度语义](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.foregroundcolor?view=openxml-3.0.1)。

原生一万行分部件链路已执行两遍有界扫描、共享公式处理、40 块 HTML 生成及统一模型重开，核对 200,000 个单元格、20,000 个公式和末行结果。它不依赖技术基线 HTML 生成器，也尚未包含完整公开导入、视觉或最终固定环境验收。

只读网格的坐标基础已使用稀疏区间前缀和处理行列度量：隐藏的百万行不展开，远端地址仍保留原始坐标，视口元素受独立限额约束。列宽按照[原生字符宽到像素的换算](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-3.0.1)，接收正常样式字体的实测最大数字宽度，再返回文档点单位。冻结区域和滚动主体已可分为四个互不重叠的逻辑窗格，共享可见单元格预算；跨隐藏行列的查询拆为实际可见窗口，不读取巨大包围矩形。合并单元格用区间索引查找，各窗格只裁剪其完整几何，内容始终读取同一左上角源；被隐藏或滚出视口的合并起点也不会丢失。点击映射与查询共用该坐标，重复窗格投影只读取一次内容。三种绘图锚点也使用相同行列度量，旋转后的可见边界参与窗格裁剪；负尺寸明确失败，不通过翻转猜测修复。逻辑滚动将原生滚动条轨道限制在固定长度，仍保留整张表的点坐标和末行定位。坐标测试不代表字体测量、虚拟 DOM 或完整 R3-08 预览已完成。

原生记录管线现在贯通异步 ZIP、全包关系审计、样式/共享字符串、逐表双遍扫描、维度压缩和 HTML 分块。32 份独立文件及 Calc 重存文件通过这条内部链路及包含批注、绘图和图表的测试包投影；两份一万行来源均保留 200,000 个单元格和 20,000 个公式。XML 全局节点、单元格、公式与输出字节跨表累计，元数据另设保留节点限制；取消、宿主进度回调异常及输入字节隔离有独立检查。投影结束后另核对所有原生部件：可达但未被工作表、表格、绘图或其他来源消费的部件也会失败。命名样式及兼容线程定位仍是公开导入的前置条件，不对外暴露部分导入。

核心文件属性中的标题、语言以及作者、主题、说明、修订、日期等来源信息保存在元数据；`properties` 的作者与应用字段不代表宿主身份或权限，日期字符串保留原始精度及时区。应用中的派生统计列表和包缩略图经过结构/资源检查后丢弃，不作为额外文档源；未知属性内容、签名和非空相对链接基址明确拒绝。原生计算链按[计算链说明](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain)仅视为旧计算次序提示，逐条复核实际公式地址及继承的工作表身份后丢弃。它不能替代依赖图，也不会使任何导入缓存变成有效计算结果。

`conditionalRules` 已有数值比较、标量公式和双色/三色色阶来源；差异外观及色阶颜色引用 CSS，阈值与优先级保留为元数据。`dataValidations` 保存列表、整数、小数、日期、时间、文本长度和自定义公式，也允许仅有提示的 `none`。规则都保存精确多区域和公式原点；严格检查身份、公式数量、样式引用、优先级及有歧义的重叠验证。已有违规数据不会在读取阶段被修正；规则求值、违规诊断和各编辑入口执行仍未实现。

原生规则验证依据 [SpreadsheetML 条件格式说明](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting)及[数据验证契约](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidation?view=openxml-2.8.1)，并单独保存[固定 Calc 的探针结果](evidence/xlsx-rule-semantics.json)：传统 `showDropDown="1"` 隐藏列表；多区域公式原点取最左列中最靠上的实际区域起点，不取列表首项或外包矩形角。后者在跨行列样例中分别为 B5/E9，并经外部 ODS 中的引用重定位核对。这是固定 Calc 的独立观测，不冒充 Excel/WPS 实机结论。运行 `python3 scripts/check-xlsx-rule-semantics.py` 重复核对；只有显式 `--record` 更新记录。

原生智能表保留为 `tables` 定义：表身份和公式名称、区域、表头/汇总行、列身份、计算列、汇总函数/公式/标签、筛选、排序状态和表样式关联。按照 [SpreadsheetML 表定义](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-tables)，公式使用的原生 `displayName` 映射为 `name`；不同的对象模型名称另存 `objectName`。工作簿名称与表名不能冲突；列名/身份、区域宽高及表/合并/普通筛选的重叠均校验。显示表头时，列名称必须与实际单元格的常量文字一致；不猜测两者谁覆盖谁。计算列模板使用真实首行原点，不把越过工作表边界的原点截回末行。普通筛选保存值/日期选择、标量比较、前后 N 项/百分比和相对日期/平均值条件；导入不重新应用筛选或排序。原生已过滤行的隐藏状态仍由行维度保留。

内建表样式使用原生名称引用，自定义 `tableStyles` 通过差异 CSS 定义区域和条带，不把表样式铺平成逐格格式。条带采用 Office 声明的 [1–9 范围](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d91a5d31-0bf4-4801-9139-f6c13a690505)，普通前后 N 项筛选采用 [1–500 范围](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/2961ac37-0f8e-47a8-89f9-981403e619b6)，百分比另限 100。源文件和 Calc 重存文件都通过表格读取与保存重开；计算列执行、筛选重算、表格编辑和准确样式渲染仍属后续验收。原生命名单元格样式在表区域上的关联也仍待对应取舍落地。

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

## 工作簿批注正文

工作簿可选 `manifest.content` 指向同一包内的辅助 HTML，使用 `data-office-kind="workbook-content"` 根节点。它仅包含有界的富文本批注 story，不拼接单元格块，也不创建另一份工作表内容。`snapshot.content` 和 `getNode` 提供正文查询；每个元素有稳定身份，复用无脚本 HTML/CSS 校验。正文和回复各自拥有一个 story；孤立正文、重复归属、未知类名、未来修订和重叠锚点区域明确拒绝。元数据和批注记录共用累计条目限额。

工作表删除后，显式 `needs_reanchor` 的批注可以保留原坐标、正文和回复，等待后续事务重定位；有效批注仍必须指向现有表。Word/PPT 不接受工作簿的 `cells` 锚点，包括需要恢复的项。当前仅实现读取及保存重开；事务映射、完整原生批注验收和交互验收仍在对应任务内。

现代批注的内部读取现在校验工作簿人员、全局 GUID、按表回复图、原生可选时间、正文及解决状态；嵌套回复的父身份保存在格式元数据中。正文仍是辅助 HTML 的唯一可编辑来源。身份、作者显示名和回复关系在打开 HTML 包时重新校验，重复、循环、跨线程父引用及未来修订不接受；导入映射按条目核算重复作者文本后的 JSON 字节，防止共享长名称放大输出。32 份内部记录链路继续通过；固定 Calc 重存的现代线程丢失仍作为外部差异记录。

经典兼容占位备注的核对、缺失定位后的处理及备注框预览尚未完成，不能据此宣称两类原生批注已验收。非空提及/扩展当前明确拒绝，不会触发通知或忽略其元数据；完整读取边界仍需最终复核。格式依据为 Microsoft 的 [CT_ThreadedComment](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/42f9b03d-9662-4204-9783-dbeb324a691c)、[CT_Person](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6274371e-7c5c-46e3-b661-cbeb4abfe968) 和[兼容占位备注](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6383f002-c90b-401c-a1d7-66b97b14cb3e)。

经典备注的内部记录也已读取作者、单元格、富文本和兼容线程标记，保留待处理的原生几何属性。线程身份、工作表和位置一致的兼容备注按照[原生占位规则](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6383f002-c90b-401c-a1d7-66b97b14cb3e)作为缓存消费，不重复显示现代线程；真正的经典备注仍保留。缺失、移位、复制或无法匹配的占位情况仍明确拒绝，等待定位恢复取舍；完全没有旧版备注部件的独立线程继续可读。独立来源的加粗中文与 Calc 重存的普通字重分别按实际来源保存，不互相补造。VML 几何与基础外观现已接入内部管线；[固定 Calc 探针](evidence/xlsx-note-semantics.json)验证其单元格锚点优先于 VML 位置数值，验证脚本为 `python scripts/check-xlsx-note-semantics.py`，仅 `--record` 更新证据。固定 Office 归档和字体清单均检查摘要；此探针不代表 Excel/WPS 或 SDK 视觉验收。

绘图锚点基础契约区分绝对坐标、单端单元格加尺寸和双端单元格锚点。统一使用文档点单位，保留负偏移、超过一行的偏移以及移动/缩放标志；DrawingML EMU/通用长度和 VML 像素在适配边界转换。三类图表、图片、形状及文本框的 12 个独立源/Calc 锚点已通过定向解析检查，零尺寸保留原值；对象与图表映射已接入内部包投影，预览仍在实施。VML 像素解释依据 [Anchor](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.spreadsheet.anchor?view=openxml-3.0.1)；坐标结构和数值界限按仓库固定的官方 Transitional XSD 检查。

工作簿 `objects` 已支持图片、16 个共享预设和文本框的严格源契约。元数据保存工作表、锚点、原生身份和打印/锁定标志，独立 HTML/CSS 保存正文、资源、外观、旋转翻转和裁剪；不重复保存一份由行列推导的位置。对象根必须是辅助内容的独立子节点，正文属于对象内的段落；校验跨表身份、预设调整、裁剪可见区域、资源摘要及累计预算。锁定标志是文件属性，不授予或取消宿主权限。

原生图片、预设形状和文本框的内部适配已通过独立第 08 号来源及 Calc 重存的包重开检查：原始图片字节、矩形裁剪、旋转、三类锚点、精确 UTF-16 正文、主题引用、段落/字符样式和换行样式均进入同一来源。对象和文字超链接绑定其 HTML 节点，沿用工作簿的安全外链/内部地址校验。背景填充引用按原生 `fillRef` 的 1001 起始编号解析，依据 [FillReference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillreference?view=openxml-3.0.1)。Calc 把样例文本框改存为普通矩形，读取器保留该来源差异；当前尚未完成公开原生导入及对象视觉验收。

图表的源契约与原生引用读取已接入工作表绘图层。系列数值、分类、系列名称和引用标题保留公式；原生缓存单独标为 `imported`，保存稀疏索引、缺值、错误值及整体/单点数字格式，不转成独立可写数据表。没有缓存时不补零或假定结果有效。缓存声明长度与元数据字节跨图表累计计量。六个独立/Calc 图表来源的 30 条引用、标题正文、图例、轴、系列及基础标签外观均通过包保存重开。轴标题同样支持独立富文本或单元格引用；格式、样式和正文所有权在核心再次校验。

原生元素缺省属性按固定官方 XSD 解析，包括百分比间距和标签偏移；可见数据、按点着色和超出最大值的标签开关保留元素是否存在，避免把未指定误当成显式设置。Calc 输出的标签引导线扩展按 [MS-ODRAWXML Datalabels](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/80e97775-804a-4f7e-90a1-8e87e1734d83)及 [showLeaderLines](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2013.drawing.chart.showleaderlines?view=openxml-3.0.1)映射；未知、重复和相互矛盾的扩展失败。当前单元格查询驱动的图表显示、完整视觉验收尚未完成。无工作表引用的原生常量向量仍待产品取舍，当前内部适配明确拒绝，不能宣称完整图表读取门槛已经通过。

经典备注框将单元格锚点、显示/打印和随行列变化标志保存在 `comments[].box`，实体填充、描边、箭头样式、单层阴影、内边距及对齐映射到受校验的 CSS 类。原生 VML 不复制进 HTML。普通 `commentPr` 坐标亦可单独读取；同时存在两套几何的情况明确拒绝，等待可靠的优先关系映射。未知控件、宏字段、非备注形状、重复定位、孤立形状和矛盾的显示状态明确失败。后者的 Calc 行为另有固定探针记录，不推断 Excel/WPS 相同。两类批注共用累计输出字节核算，源包重开会复核作者、正文、回复和备注框样式引用。32 份内部记录与包投影检查已包含这些基础映射，仍不代替公开导入及只读网格验收。
