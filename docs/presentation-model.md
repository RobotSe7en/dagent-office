# Presentation 契约与文档模型

R2-02 实现 Presentation 的严格契约、新建文稿及只读 HTML 文档包模型。当前支持创建、打开、查询和保存重开；PPTX 导入、幻灯片渲染、编辑、请求／提案及原生导出仍按[第二轮细则](iteration-2.md)后续任务实施，不把模型可读当作读取或视觉验收通过。

## 公开入口

```ts
import { createPresentation, openDocument, readPresentationGeometry } from "@dagent-office/core";

const files = createPresentation({ documentId: "host-supplied-id", title: "项目进展" });
const model = openDocument(files);
const snapshot = model.snapshot();
const presentation = snapshot.metadata.presentation;
if (!presentation) throw new Error("Expected a presentation.");
const slide = model.getNode(presentation.slides[0]!.rootNodeId);
const title = model.getNode("title-centered");
const geometry = title ? readPresentationGeometry(title.attributes) : undefined;
const reopened = openDocument(model.serialize());
```

createPresentation 接收宿主提供的 documentId，以及可选 title、locale 和 size；第二个参数接受与 openDocument 一致的读取预算。默认标题为“未命名演示文稿”、语言 zh-CN、尺寸 960 × 540 pt，生成一页空白幻灯片和空白／标题／标题与正文三种布局。新包修订为 0，默认现代批注格式。入口不访问 DOM、文件系统、网络、模型或字体下载服务。

readPresentationGeometry 从源属性读取局部 x/y/width/height，及可选 rotation/flip-h/flip-v。缺少整个变换时返回 undefined，供有匹配占位符的对象继承；部分变换、无效数值、负尺寸或非法布尔值失败。函数不读取 CSS 或屏幕缩放。需要几何的普通对象、可逆组及允许单轴零尺寸的连接线在完整文档校验时进一步区分。

## 类型与来源

document-metadata Schema 是 Word 与 Presentation 两个严格分支的联合。DocumentMetadata 表达两种已实现的元数据；WordDocumentMetadata 表达原有的页面／节／Word 语义分支，PresentationDocumentMetadata 表达 title/locale/presentation 分支。PPT 不包含伪造的 pageSetup、sections 或 word 字段，旧 Word 包仍使用原有 JSON 结构。

PresentationMetadata 显式定义尺寸、批注格式、幻灯片、主题、母版、布局、备注母版、备注／批注内容域、对象、图表、数据集、原生批注关联和链接。对象按 text/image/geometry/connector/table/chart/group 区分；无关字段不能在其他对象类型上混用。AnnotationReply 的纯文字／富文本正文互斥，线程与回复身份必须唯一。Word 的原生线程读取尚未开放，Word 包不能借此嵌入回复。

- 一份 HTML 根标记为 presentation；幻灯片、母版、布局、备注母版、备注和批注各自有登记的源容器及内容域。
- 对象几何只在 data-office-* 属性中保存；组另存 child-x/y/width/height。层级由同级 HTML 节点顺序定义，元数据数组不形成第二份顺序来源。
- 对象和定义只有一份源节点。未提供变换的占位符必须经布局／母版身份找到有效源变换；不能根据临近位置猜测。
- 外观与主题值保存在 CSS；元数据只引用主题类及原生定义。布局、连接线、图表、标题和内容域引用须存在且匹配类型。
- 图表分类与系列常量仅保存在 dataSets；长度、系列身份、饼图单系列和非负值、轴范围均校验。图表定义与对象双向匹配，标题必须位于对应图表源内。
- 图片对象仅包含一个授权图片，裁剪须保留可见面积。表格对象只包含一张表，单元格覆盖、合并和网格不得越界或重叠；展开前检查工作量预算。
- 备注与富文本批注有独立内容域。回复正文不能共享可独立编辑的内容根，来源作者仅用于显示。

HTML 继续由同一无脚本解析、安全样式表和资源校验处理。未知格式版本及 workbook 拒绝；清单类型与元数据或根角色不匹配也拒绝。R2-06 读取验收前，createDocumentSession 对重新校验后的 PPT 源明确返回 UnsupportedFeature，不能用调用者伪造的快照绕过阶段边界。

## 验证与样例

[完整合成包](../examples/presentation/README.md)包含所有第一批对象类型、图表数据、备注、富文本批注和回复。它可由 openDocument 直接读取，并通过 serialize 保存重开；它不声称当前已有图形渲染或 PPTX 生成。

对应工程检查：

```bash
pnpm build:packages
pnpm exec vitest run packages/core/tests/presentation-model.test.ts
pnpm exec tsc --project tsconfig.tools.json
```

模型用例覆盖严格 Schema 与独立 Ajv 校验、来源顺序、负坐标、UTF-16、继承引用、表格合并、图表数据、回复身份、危险 HTML、源隔离、预算、保存重开和编辑入口阻断。公共 Word 类型调用处按 WordDocumentMetadata 或已检查的元数据分支收窄，保留原有 Word 读取、事务及导出回归。
