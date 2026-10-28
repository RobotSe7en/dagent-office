# 技术选型

第一版采用 TypeScript 和浏览器 HTML/CSS 排版。文档协议、批注、事务、提案校验及 DOCX/PPTX/XLSX 的 OOXML 读写和样式映射由本项目实现；文本输入、浏览器渲染及 ZIP/XML/哈希等通用能力复用成熟基础库。宿主智能体负责模型调用和运行，SDK 不引入模型供应商依赖。R1-02 至 R1-05 已锁定工程、结构校验、HTML/CSS、资源哈希及 ZIP 依赖；编辑、XML 与公式计算依赖按后续任务验证。

## 已锁定的结构校验依赖

| 依赖 | 版本 | 用途 | 许可证 |
| --- | --- | --- | --- |
| Zod | 4.1.12 | 严格结构、推导类型和原生 JSON Schema 输出 | MIT |
| Ajv | 8.20.0 | 仅开发：独立验证 Draft 2020-12 Schema | MIT |

本轮验证中 Zod 4.6.7 的声明直接依赖全局 `URL`，在核心仅 ES2022、无 DOM/Node 类型且不跳过依赖检查的配置下编译失败，故锁定已通过该配置及外部消费验证的 4.1.12。后续升级必须重新通过类型、Schema 一致性和打包消费，不通过增加宿主全局类型掩盖兼容问题。

## 已锁定的文档解析依赖

| 依赖 | 版本 | 用途 | 许可证 |
| --- | --- | --- | --- |
| parse5 | 8.0.1 | HTML5 解析与序列化，内部节点索引 | MIT |
| CSS Tree (`css-tree`) | 3.2.1 | CSS AST、生成、属性语法检查 | MIT |
| `@types/css-tree` | 3.2.0 | 仅核心开发：解析器类型 | MIT |
| entities | 8.1.0 | parse5 间接依赖：HTML 字符引用 | BSD-2-Clause |
| mdn-data | 2.27.1 | CSS Tree 间接依赖：CSS 语法数据 | CC0-1.0 |
| source-map-js | 1.2.2 | CSS Tree 间接依赖：源映射 | BSD-3-Clause |

直接依赖锁定精确版本，间接解析版本记录于 `pnpm-lock.yaml`，许可证来自已安装包的 manifest。R1-04 验证报告与合成 HTML/CSS 的保存重开、严格 ES2022 类型、Node 打包执行和普通 JS 浏览器打包。第三方 AST 留在核心内部；UTF-8 使用支持环境内建的 TextEncoder/TextDecoder，经内部适配器调用，不向消费者引入 DOM 或 Node 全局声明。支持范围和限制见[文档模型](document-model.md)。

## 已锁定的资源与 ZIP 依赖

| 依赖 | 版本 | 用途 | 许可证 |
| --- | --- | --- | --- |
| `@noble/hashes` | 2.4.0 | `core`：同步、跨运行环境的 SHA-256；保持现有打开接口同步 | MIT |
| fflate | 0.8.3 | `office`：DEFLATE 流及 ZIP 写出，使用不依赖 Node 的 browser 入口 | MIT |

资源 MIME 识别和基本容器边界由内部轻量适配器检查，不实现像素解码。ZIP 适配层在解压前按原始记录检查路径、重复项、结构及预算，压缩解压由 fflate 执行；不能只依赖上层解包 Map，因为此时重复名可能已被覆盖。版本与许可来自已安装包，支持范围见[资源与 ZIP](resources-and-zip.md)。[fflate](https://github.com/101arrowz/fflate)、[noble-hashes](https://github.com/paulmillr/noble-hashes)

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
| OOXML 公共基础 | fflate（已锁定）；XML 解析器待验证，如 fast-xml-parser | 自行实现 OPC 包、关系及三格式映射 |
| DOCX 读写 | 自研 OOXML 适配器 | 原生段落、文字、样式、编号、节和页面语义 |
| PPTX 读写 | 自研 OOXML 适配器 | 原生对象、主题、母版、对象树和几何映射 |
| XLSX 读写 | 自研 OOXML 适配器 | 原生单元格、公式、样式、工作表及关系映射 |
| 渲染与 PDF | Chromium、Playwright | 固定字体、布局测量、截图、导出与视觉回归 |
| 测试 | Vitest、Playwright | 核心、协议、输入、可编辑重开、视觉和宿主边界 |

Zod 4 支持原生 JSON Schema 转换；结构契约只使用可导出的类型，节点引用、版本、样式和资源等语义单独校验。不使用转换失败后放宽为任意类型的策略。[JSON Schema 文档](https://zod.dev/json-schema)

parse5 提供 HTML 解析和序列化，CSS Tree 提供 CSS AST、遍历、生成和语法检查。二者经内部适配层接入，不向公共接口暴露第三方 AST。浏览器打包及支持属性通过首轮样例验证。[parse5](https://parse5.js.org/)、[CSS Tree](https://github.com/csstree/csstree)

编辑器以 DOM 组件实现，普通 JS 挂载和可选 React 包装共用同一 UI。React 不进入核心；非 React 宿主可直接挂载。四个库模块的依赖方向见[总体架构](architecture.md)。

Tiptap 的开源核心采用 MIT 许可，官方也提供付费批注和 AI 扩展。本项目只使用基础编辑能力，公共批注和宿主交互协议自行实现；首轮先验证属性、中文输入和事务映射，再完成正式 Word 输入及分页，不采用独立的编辑器保存源或撤销历史。[官方概览](https://tiptap.dev/docs/editor/getting-started/overview)

Univer 以插件和命令组织能力，并能在浏览器与 Node 环境使用。评估公式能力时只引入所需基础模块；其 Pro 转换与协同能力不作为本项目导入导出的前提。[基础概念](https://docs.univer.ai/guides/sheets/getting-started/quickstart)、[Node 运行](https://docs.univer.ai/guides/sheets/getting-started/node)

Office 转换路线已确定为自研 OOXML，不采用 docx、PptxGenJS、ExcelJS。依据公开格式标准实现 DOCX/PPTX/XLSX 的解析、写回、关系和样式映射，以便掌握原文档特性的保留策略；兼容性以真实格式样例、原生可编辑重开及固定环境视觉对照验收，不以自行实现本身作为保真证明。ZIP、XML、哈希等通用能力继续复用成熟库。公式保存和公式计算分开实现，不能只写入一个公式表达式就认为依赖结果已更新。

## 原则

自研内容包括文档格式、OOXML 读写与样式映射，不重新实现浏览器文字输入、通用 ZIP 压缩、XML 解析及哈希基础设施。第一版不引入 Rust 构建链；后续 Rust 或 WASM 实现可以替换格式适配器及性能关键路径。

每个格式适配器声明自己支持的导入、编辑和导出特性，记录具体库版本与许可证。基础库没有覆盖的样式和对象由适配层补齐，无法完成的转换返回明确结果。

第一轮按“HTML 基础 → 独立 Word 闭环 → 达智智能体验证”推进。DOCX、Tiptap 和 PDF 纳入第一轮正式验收，PPTX 和 XLSX 先做小型验证。ZIP 候选还须验证原始路径、重复项和有界展开，不能仅以库能成功打开文件作为通过标准。详细任务见[第一轮实施细则](iteration-1.md)。

独立发布前确定本项目许可证，并提供依赖及字体许可清单。尚未确定第一方许可证，当前目录不代表已作出开源授权。

## ibisOffice 参考

参考提交为 [50287728b005ceb0ed8a4a8f7d825ac4824224ef](https://github.com/ibisjp/ibis-office/tree/50287728b005ceb0ed8a4a8f7d825ac4824224ef)。可借鉴的架构包括 UI 独立模型、格式适配、命令批量执行、一次撤销和共享编辑路径。[模型入口](https://github.com/ibisjp/ibis-office/blob/50287728b005ceb0ed8a4a8f7d825ac4824224ef/crates/ibis-doc-model/src/lib.rs)、[命令实现](https://github.com/ibisjp/ibis-office/blob/50287728b005ceb0ed8a4a8f7d825ac4824224ef/crates/ibis-doc-model/src/command.rs#L1080)

其采用 Rust 文档模型与 GPUI/WebGPU 视图，保存器保留原始包并将模型变化写回部件。[保存器](https://github.com/ibisjp/ibis-office/blob/50287728b005ceb0ed8a4a8f7d825ac4824224ef/crates/ibis-pptx/src/save/mod.rs)

核心库在隔离工作区使用 Rust 1.99.0 的五项冒烟验证通过，覆盖中文 Word 编辑与撤销、Excel 公式重开、PPT 文本框、批量失败回滚和简单 HTML 转换。该结果验证基础接口链路，不构成完整浏览器视觉兼容性或所有 Office 特性的结论。公开仓库的部分测试引用未包含的 `test/editing` 夹具，完整测试需要补齐上游数据。

ibisOffice 采用 AGPL-3.0-only 和商业双授权。本项目当前不依赖、复制或翻译其源码；根据本项目文档格式、公开标准和独立测试实现所需能力。[上游许可](https://github.com/ibisjp/ibis-office/blob/50287728b005ceb0ed8a4a8f7d825ac4824224ef/COMMERCIAL-LICENSE.md)
