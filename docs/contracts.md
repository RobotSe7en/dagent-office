# 结构契约与 JSON Schema

R1-03 提供 `core` 的数据结构校验、推导类型、JSON Schema、稳定校验错误，以及会话等运行时接口的类型声明。R1-04 的文档包打开、HTML/CSS 解析、节点与元数据引用检查见[文档模型](document-model.md)；事务执行和界面尚未实现。

## 公开入口

从 `@dagent-office/core` 导入契约：

```ts
import {
  getContractJsonSchema,
  parseContract,
} from "@dagent-office/core";

// input 是宿主提供的未知数据；此处不执行文件、网络或模型调用。
const manifest = parseContract("document-manifest", input);
const proposalSchema = getContractJsonSchema("edit-proposal");
```

`parseContract(name, input)` 根据名称推导具体返回类型，成功时返回独立的解析结果。失败抛出带稳定 `code` 的 `DocumentError`。也可使用公开的 Zod Schema 或 `contractSchemas` 注册表进行结构校验；直接调用 Zod 的方法时，返回的是 Zod 自身的错误，而非 SDK 错误码。

`getContractJsonSchema(name)` 每次生成独立、完整的 Draft 2020-12 Schema。根 `$id` 使用包含契约名称及版本的 URN，嵌套定义在同一文件中，不依赖网络解析。源契约无法转换时生成失败，不退化为任意类型。

## 已提供的契约

| 名称 | 数据 |
| --- | --- |
| `asset-descriptor` | 资源的 `path`、`mimeType` 和小写十六进制 `sha256` |
| `document-manifest` | 当前版本单 HTML 入口的文档清单 |
| `document-metadata` | Word 页面设置和节引用；标题及语言可选 |
| `anchor` | 单节点文字、跨节点文字、节点、节点集合和区域 |
| `annotation`、`annotations` | 批注及批注列表 |
| `content-operation` | 当前六种内容操作 |
| `document-operation` | 内容操作和三种批注管理操作 |
| `document-transaction` | 协议、文档、基准修订和操作列表 |
| `interaction-request` | 开放批注、完整区块、样式、元数据、资源和支持操作 |
| `edit-proposal` | 请求关联、提案身份、已处理批注及内容操作 |
| `document-error` | 有界的错误码、消息和结构问题列表 |
| `document-event` | 文档变化、批注变化和错误事件 |

各对象、操作及事件的 TypeScript 类型由同一 Zod 定义推导。持久化文件与宿主消息使用 JSON Schema；`DocumentPackage`、`DocumentSession`、`CandidateChange` 等运行时值不伪装成可序列化消息。

## 严格字段与版本规则

所有协议对象均严格拒绝未知字段，包括元数据、锚点、嵌套操作和上下文。不提供通用 `extensions` 字段，不静默删字段，不自动填默认值，不强制转换字符串数字，不 trim 或归一化原文。

格式和协议的当前设计版本均为 1。显式声明其他整数版本时，SDK 入口返回 `UnsupportedVersion`；缺少版本或把版本写成字符串属于结构错误，不猜测版本。

清单的类型枚举识别 `document`、`presentation`、`workbook`，但识别名称不代表已实现相应文档读取。当前元数据和交互上下文只定义 Word 类型；PPT/Excel 的上下文、几何与单元格操作在对应任务显式加入。分块清单、编号、复杂样式映射、页眉页脚、批注回复和作者显示字段也须在对应任务定义后接收，当前不能通过额外 JSON 字段绕过。

开发草案尚未发布，后续任务可以显式调整契约。发布冻结后，影响严格读取器的新增字段采用显式版本演进和支持的读取器，不能依赖忽略未知字段实现兼容。

## 字段与操作约定

- 修订和文本偏移为非负安全整数；矩形位置非负、尺寸为正数，几何单位固定为 pt。
- 标识是不透明的非空字符串，不按 URL、文件路径或权限解释。
- 包路径采用规范 POSIX 相对路径，拒绝绝对路径、反斜杠、冒号、空段、点段、控制字符和查询/片段分隔符，保留空格和中文。包路径是字面路径；URL 解码、引用映射及授权不能由该正则替代。
- 页面设置要求 pt 宽高及四边距；至少一个节包含 `sectionId` 和 `rootNodeId`。`title` 和 `locale` 可省略，解析器不设置默认值。
- `replaceText` 使用 `nodeId/from/to/expectedText/text`；`replaceFragment` 使用 `nodeId/html`，替换目标的子内容。
- `insertFragment` 使用 `nodeId/position/html`，位置为 `before`、`after`、`start` 或 `end`；前两者插入同级，后两者插入子内容。
- `removeNode` 使用 `nodeId`；`setStyle` 使用 `nodeId/property/value`；`replaceAsset` 使用 `nodeId/assetPath`，资源必须由后续引用校验确认可用。
- `addAnnotation` 携带完整 `annotation`；`updateAnnotation` 使用 `annotationId/changes`，只允许修改锚点、说明及状态；`removeAnnotation` 使用 `annotationId`。
- 区域锚点使用 `surfaceId/rect/nodeIds/layoutKey`，可附带真实 `pageId`。`rect` 为 `x/y/width/height`，空 `nodeIds` 表示空白框选。
- 请求只接收 `open` 批注。提案只接收内容操作，不能直接提交批注状态管理操作。

数组结构允许空事务与空提案；是否净变化、是否更新批注由未来执行层判断。跨节点范围排序、重复 ID、范围先后及原文匹配也是后续语义检查，不由结构校验推断。

## JSON Schema 产物

`pnpm build:packages` 在编译后从同一契约生成 13 份 JSON，位于核心构建目录的 `schemas`。各文件通过明确的公开子路径分发，例如：

```js
import proposalSchema from "@dagent-office/core/schemas/edit-proposal.json"
  with { type: "json" };
```

不在源码中手工维护第二份 JSON Schema。普通构建、`pnpm pack:packages` 和打包消费测试均生成并核对这些文件。开发监听更新 JS 与声明；修改契约后运行 `pnpm build:packages` 刷新静态 JSON。

测试使用独立的 Ajv 2020 校验器，对真实示例和非法输入同时执行 Zod 与导出 Schema 校验，并核对打包产物。Ajv 只属于开发依赖，不进入 SDK。

## 错误与运行时类型

`DocumentError` 包含 `code`、`message` 和 `details`，`toJSON()` 输出 `document-error` 数据。消息最多 256 字符，问题最多 16 项，每项路径最多 16 段；字符串长度上限按 Unicode 码点计算。不保存原始输入或原始校验器错误，不把文档原文回显进错误。

`parseContract` 在完整解析前逐项预检数组，包括嵌套对象和契约分支中的数组；遇到无效项立即停止，保留完整字段路径，不为后续数组项继续收集错误。直接调用公开 Zod Schema 的方法仍遵循 Zod 自身的错误收集行为。

文档结构错误使用 `InvalidDocument`，操作及事务使用 `InvalidOperation`，提案使用 `InvalidProposal`。格式标识不支持为 `UnsupportedFormat`，当前没有实现的交互类型或契约为 `UnsupportedFeature`。读取或校验错误不改变原始字节和输入对象。

运行时类型的职责与实现状态：

- `DocumentPackage` 为只读路径到字节的 Map，`manifest.json` 在其中，避免另一份可独立修改的 manifest。
- `DocumentSnapshot` 为文档源的派生只读视图，不是新的保存格式。
- `DocumentModel` 由 `openDocument` 创建，已实现 `snapshot/getNode/serialize`；返回值冻结或独立复制，第三方 AST 不进入公共接口。
- `CandidateChange` 是非序列化的候选类型；实际会话仍需私有状态验证候选所有权和修订，类型标记不代表授权。
- `DocumentSession` 声明查询、准备、提交、撤销、序列化和释放，当前没有工厂或执行实现。
- 核心事件声明 `protocolVersion/documentId/revision`，文档变化携带节点 ID，批注变化携带批注 ID，错误携带结构化错误数据。

## 校验边界与后续任务

`parseContract` 和导出的 JSON Schema 只证明字段、类型、枚举和结构约束正确。它们不解析 HTML/CSS、不证明脚本安全、不验证 SHA-256 字节、不读取包资源、不验证节点存在，也不执行提案、撤销或宿主保存。

R1-04 的 `openDocument` 已实现受支持 HTML/CSS 的解析与拒绝策略、节点索引、元数据及非失效批注的节点引用和规范化保存；R1-05 已实现资源实际字节摘要、受支持图片容器检查、完整资源收集及 ZIP，见[资源与 ZIP](resources-and-zip.md)。R1-06/R1-07 实现操作、原子性、修订、历史及锚点映射；请求/提案关联、支持操作子集和跨文档匹配在 R1-09 的实际交互中校验。结构校验不能代替这些执行及渲染边界。

主协议见[编辑与宿主交互协议](editing-protocol.md)，格式见[文档格式](document-format.md)，构建和测试见[本地开发](development.md)。
