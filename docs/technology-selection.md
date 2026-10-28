# 技术选型

第一版采用 TypeScript 和浏览器 HTML/CSS 排版。文档协议、批注、事务和提案校验由本项目实现；文本输入、格式写出和浏览器渲染复用成熟基础能力。宿主智能体负责模型调用和运行，SDK 不引入模型供应商依赖。工程工具已在 R1-02 安装并锁定，文档功能依赖仍为计划选型。

## 已锁定的工程依赖

| 依赖 | 版本 | 用途 | 许可证 |
| --- | --- | --- | --- |
| TypeScript | 6.0.3 | 编译、声明和包边界 AST 检查 | Apache-2.0 |
| Vite | 8.3.5 | Playground 开发、构建及外部浏览器消费验证 | MIT |
| Vitest | 5.0.5 | 工程与后续核心测试 | MIT |
| @playwright/test | 1.64.1 | 浏览器烟测与后续视觉/交互验证 | Apache-2.0 |
| @vitejs/plugin-react | 6.1.2 | Playground React 构建 | MIT |
| React、React DOM | 19.3.1 | 仅 Playground 的运行时 | MIT |
| @types/node | 24.19.2 | Node 模块、工具及测试类型 | MIT |
| @types/react、@types/react-dom | 19.3.0 | React 开发类型 | MIT |

开发环境使用 Node.js 24，pnpm 固定为 10.28.2；全部解析版本见根锁文件。许可证信息来自已安装依赖的 manifest，不表示第一方许可证已经确定。

TypeScript 暂用 6.0.3，以保留工程边界检查所需的编译器 API；7.0 尚未提供该程序化 API，待工具链适配后再评估升级。[TypeScript 官方说明](https://devblogs.microsoft.com/typescript/announcing-typescript-7-0/)

## 基础组件

| 责任 | 首选或验证对象 | 本项目补充的能力 |
| --- | --- | --- |
| monorepo 和构建 | pnpm workspace、TypeScript、Vite | 包边界、独立 playground、打包消费检查 |
| 协议与结构校验 | Zod 4 | 同源推导类型及 JSON Schema，业务语义另行校验 |
| HTML 解析与序列化 | parse5 | 支持结构、规范化、稳定节点和内部文档树适配 |
| CSS 解析与检查 | CSS Tree | 样式能力表、值和引用约束、变量及级联源保留 |
| 文本编辑 | Tiptap 开源基础层及 ProseMirror | 文档节点、样式保留、事务映射、分页配合 |
| PPT 编辑 | HTML 定位布局、SVG、自有对象操作 | 几何、层级、框选、组合、母版与主题 |
| Excel 编辑 | 虚拟网格、自有单元格适配器 | HTML 文档源和计算器数据间的转换 |
| Excel 计算 | 优先验证 Univer 基础公式能力 | 支持函数矩阵、计算状态和依赖更新桥接 |
| OOXML 公共基础 | JSZip、fast-xml-parser 等 | 关系、资源、样式继承和三格式读入 |
| DOCX 输出 | `docx` | 当前 HTML 结构、样式及页面语义的编译 |
| PPTX 输出 | PptxGenJS | HTML 对象到原生幻灯片对象的编译 |
| XLSX 读写 | ExcelJS | 文档源映射、公式计算器连接、超出库能力的格式适配 |
| 渲染与 PDF | Chromium、Playwright | 固定字体、布局测量、截图、导出与视觉回归 |
| 测试 | Vitest、Playwright | 核心、协议、输入、可编辑重开、视觉和宿主边界 |

Zod 4 支持原生 JSON Schema 转换；结构契约只使用可导出的类型，节点引用、版本、样式和资源等语义单独校验。不使用转换失败后放宽为任意类型的策略。[JSON Schema 文档](https://zod.dev/json-schema)

parse5 提供 HTML 解析和序列化，CSS Tree 提供 CSS AST、遍历、生成和语法检查。二者经内部适配层接入，不向公共接口暴露第三方 AST。浏览器打包及支持属性通过首轮样例验证。[parse5](https://parse5.js.org/)、[CSS Tree](https://github.com/csstree/csstree)

编辑器以 DOM 组件实现，普通 JS 挂载和可选 React 包装共用同一 UI。React 不进入核心；非 React 宿主可直接挂载。四个库模块的依赖方向见[总体架构](architecture.md)。

Tiptap 的开源核心采用 MIT 许可，官方也提供付费批注和 AI 扩展。本项目只使用基础编辑能力，公共批注和宿主交互协议自行实现；首轮先验证属性、中文输入和事务映射，再完成正式 Word 输入及分页，不采用独立的编辑器保存源或撤销历史。[官方概览](https://tiptap.dev/docs/editor/getting-started/overview)

Univer 以插件和命令组织能力，并能在浏览器与 Node 环境使用。评估公式能力时只引入所需基础模块；其 Pro 转换与协同能力不作为本项目导入导出的前提。[基础概念](https://docs.univer.ai/guides/sheets/getting-started/quickstart)、[Node 运行](https://docs.univer.ai/guides/sheets/getting-started/node)

`docx` 提供浏览器及 Node 的 DOCX 生成能力。PptxGenJS 提供文字、形状、表格、图表和图片输出，其 `tableToSlides` 处理的是 HTML 表格，整页 HTML 的映射由本项目完成。[docx](https://github.com/dolanmiu/docx)、[PptxGenJS 对象能力](https://gitbrent.github.io/PptxGenJS/docs/introduction/)、[HTML 表格转换](https://gitbrent.github.io/PptxGenJS/docs/html-to-powerpoint.html)

ExcelJS 负责工作簿读写。公式保存和公式计算分开实现，不能只写入一个公式表达式就认为依赖结果已更新。[ExcelJS](https://github.com/exceljs/exceljs)

## 原则

自研内容围绕产品差异和文档格式，不重新实现浏览器文字输入、ZIP 压缩或通用 Office 输出的所有基础设施。第一版不引入 Rust 构建链；后续 Rust 或 WASM 实现可以替换格式适配器及性能关键路径。

每个格式适配器声明自己支持的导入、编辑和导出特性，记录具体库版本与许可证。基础库没有覆盖的样式和对象由适配层补齐，无法完成的转换返回明确结果。

第一轮按“HTML 基础 → 独立 Word 闭环 → 达智智能体验证”推进。DOCX、Tiptap 和 PDF 纳入第一轮正式验收，PPTX 和 XLSX 先做小型验证。ZIP 候选还须验证原始路径、重复项和有界展开，不能仅以库能成功打开文件作为通过标准。详细任务见[第一轮实施细则](iteration-1.md)。

独立发布前确定本项目许可证，并提供依赖及字体许可清单。尚未确定第一方许可证，当前目录不代表已作出开源授权。

## ibisOffice 参考

参考提交为 [50287728b005ceb0ed8a4a8f7d825ac4824224ef](https://github.com/ibisjp/ibis-office/tree/50287728b005ceb0ed8a4a8f7d825ac4824224ef)。可借鉴的架构包括 UI 独立模型、格式适配、命令批量执行、一次撤销和共享编辑路径。[模型入口](https://github.com/ibisjp/ibis-office/blob/50287728b005ceb0ed8a4a8f7d825ac4824224ef/crates/ibis-doc-model/src/lib.rs)、[命令实现](https://github.com/ibisjp/ibis-office/blob/50287728b005ceb0ed8a4a8f7d825ac4824224ef/crates/ibis-doc-model/src/command.rs#L1080)

其采用 Rust 文档模型与 GPUI/WebGPU 视图，保存器保留原始包并将模型变化写回部件。[保存器](https://github.com/ibisjp/ibis-office/blob/50287728b005ceb0ed8a4a8f7d825ac4824224ef/crates/ibis-pptx/src/save/mod.rs)

核心库在隔离工作区使用 Rust 1.99.0 的五项冒烟验证通过，覆盖中文 Word 编辑与撤销、Excel 公式重开、PPT 文本框、批量失败回滚和简单 HTML 转换。该结果验证基础接口链路，不构成完整浏览器视觉兼容性或所有 Office 特性的结论。公开仓库的部分测试引用未包含的 `test/editing` 夹具，完整测试需要补齐上游数据。

ibisOffice 采用 AGPL-3.0-only 和商业双授权。本项目当前不依赖、复制或翻译其源码；根据本项目文档格式、公开标准和独立测试实现所需能力。[上游许可](https://github.com/ibisjp/ibis-office/blob/50287728b005ceb0ed8a4a8f7d825ac4824224ef/COMMERCIAL-LICENSE.md)
