# 本地开发与工程检查

当前完成 R1-02 至 R1-09、R1-05A 模型设计与 R1-10 声明范围的 DOCX 读取验收。`core` 提供严格契约、Word 语义和文档包，`office` 提供 ZIP 与 DOCX 导入，`editor` 提供[预览、选择与批注交互](preview-interaction.md)；`node` 仍为工程入口。`core` 还提供[编辑会话、事务及历史](document-session.md)、精确文字及结构锚点映射和区域批注重定位；[工作台与协议闭环](workbench.md)支持多批注请求、提案差异和整批应用；下一步是 R1-11 富文本输入，输入、自动分页及 DOCX/PDF 导出尚未实现。API 见[DOCX 读取](docx-import.md)，顺序见[第一轮细则](iteration-1.md)。

## 环境和安装

开发基线为 Node.js 24 和 pnpm 10.28.2，当前验证环境为 Node.js 24.13.0。`.node-version` 和 `packageManager` 记录开发版本，`pnpm-lock.yaml` 固定依赖解析结果。

```bash
pnpm install --frozen-lockfile
pnpm dev
```

`pnpm dev` 先构建库，再同时运行库编译监听与 Vite。开发服务器绑定本机，访问终端显示的地址；停止命令时一并结束监听进程。

Playground 通过公开入口打开 DOCX/文档包、选择文字／对象／区域、管理批注、缩放、撤销重做、下载当前会话的文档包及重开。开发宿主将公开合成样例作为静态资源提供；SDK 不自行访问样例目录。它使用 React 生命周期挂载同一 DOM 组件，SDK 的 `core`、`office` 和 `node` 不依赖 React；编辑器专用 React 包装通过可选 `@dagent-office/editor/react` 入口提供；普通 JS 入口不要求 React。

内部 workspace 暂用 `@dagent-office/core`、`@dagent-office/editor`、`@dagent-office/office` 和 `@dagent-office/node`。版本均为 `0.0.0`，包均设为 `private`，许可证字段为 `UNLICENSED`；正式名称和第一方许可证仍按发布计划确定。

## 构建与类型检查

| 命令 | 行为 |
| --- | --- |
| `pnpm build:packages` | 按项目引用构建四个库，输出 JS、声明、映射和 13 份 JSON Schema |
| `pnpm typecheck` | 检查库、Playground、工具配置和测试类型 |
| `pnpm build` | 构建库和生产 Playground |
| `pnpm check:boundaries` | 检查允许的依赖方向、公开入口及浏览器模块边界 |
| `pnpm check` | 依次执行包边界、类型、工程测试和生产构建 |

库的 TypeScript 项目引用需要声明输出，所以 `typecheck` 会先构建库；Playground 和工具类型检查使用 `noEmit`。生成物进入各包 `dist`，增量元数据进入根 `.cache`，均不提交。GitHub Actions 的 `CI / check` 在 PR 和主分支推送时，以锁定的 Node/pnpm 和依赖运行 `pnpm check`；固定字体浏览器视觉验收仍按下文单独执行。

库使用 NodeNext ESM 解析；相对 TypeScript 源码导入写对应的 `.js` 扩展名。跨包使用 workspace 包名和公开 `exports`，不配置指向其他包源码的路径别名。

核心与格式库的类型环境仅包含 ES2022，无 DOM 和自动加载的 Node 类型；编辑器增加 DOM，Node 入口单独启用 Node 类型。工具和测试不参与库发布。

核心通过 Zod 推导并生成普通声明，关闭该包的 `isolatedDeclarations` 限制，仍保持 `strict`、声明输出及 `skipLibCheck: false`。避免为满足隔离声明而手工维护第二份数据类型。静态 Schema 从编译后的同一契约生成；开发监听更新 JS 和声明，修改契约后运行 `pnpm build:packages` 刷新静态 JSON。

## 工程测试

```bash
pnpm test
pnpm test:watch
```

当前 Vitest 覆盖结构契约、JSON Schema、包边界及打包消费。核心用例使用真实报告、请求和提案，并由独立 Ajv 校验器检查导出 Schema 对相同输入的接受/拒绝结果。边界用例检查反向依赖、未声明依赖、内部路径、跨包相对导入、Node 内置模块和不可静态检查的动态导入。

打包用例在 workspace 外、包含空格的临时路径中安装四个 tarball，优先复用工作区 pnpm 依赖缓存；冷缓存时允许获取第三方依赖及其注册表元数据。内部包通过显式 tarball 覆盖固定来源，不从注册表下载内部包。验证真实校验 API、公开 JSON、ES2022 类型、私有路径拒绝，以及不安装 React 的普通 JS 浏览器宿主构建。临时消费工程在测试结束后删除。

测试还覆盖只读文档解析、资源摘要、宿主资源收集、ZIP 原始路径及大小校验、保真样例保存重开；DOCX 用例覆盖实际格式输入、独立生产者版本、样式、字段、批注、修订及保存重开；事务用例另覆盖隔离候选、操作、原子性、历史、冲突和预算，精确 UTF-16/结构映射、跨节点与区域批注，以及 DOCX 原生批注编辑后的历史和文档包重开；输入、自动分页和 Office 导出仍待实现。核心用例放入各包 `tests`，浏览器用例放入根 `tests/browser`。

## 浏览器烟测

```bash
pnpm browser:install
pnpm test:browser
```

浏览器安装和测试通过同一包装脚本运行。默认将 Playwright 浏览器保存在仓库的 `.playwright` 缓存，禁用浏览器缓存回收；可用显式的 `PLAYWRIGHT_BROWSERS_PATH` 覆盖路径。

浏览器测试先构建生产 Playground，再启动严格占用本机 4173 端口的临时预览服务，验证 React 挂载、实际 DOCX 读取、文档包下载重开、字体诊断及连续预览截图。服务由 Playwright 管理，测试结束后关闭，不借用已有预览服务。

本机已经安装 Google Chrome 时，可显式选择其渠道：

```bash
DAGENT_OFFICE_BROWSER_CHANNEL=chrome pnpm test:browser
```

未设置渠道时使用对应 Playwright 版本的 Chromium，不进行自动渠道回退。当前截图基线使用锁定的 Chromium 与 Noto 中文字体；补充的仿宋/楷体浏览器用例还需要[固定开源字体环境](fonts.md)。环境和字体摘要见[参考记录](../examples/docx/reference.json)。使用 Chrome 渠道只能作额外验证，不能静默刷新固定基线。独立参考检查使用 `pnpm check:docx-reference`，更新方法见[DOCX 读取](docx-import.md#独立样例与验收)。

仅运行 R1-08 交互验收使用 `node scripts/playwright.mjs test interaction.spec.ts workbench.spec.ts`；它包含普通 JavaScript 宿主及 React Playground。CI 分别执行 `pnpm check` 和浏览器交互／工作台检查，失败时上传 Playwright 结果。文档视觉夹具截图时展开滚动容器，继续核对原完整文档基线；交互滚动与缩放另有坐标验收。

## 本地打包

```bash
pnpm pack:packages
```

四个 tarball 输出到 `.cache/packages`。pnpm 会把内部 `workspace:*` 依赖转换为包版本，`files` 仅包含构建目录；包的公开入口指向构建结果，不暴露 `src` 或 `dist` 内部路径。

使用 `pack:packages` 避免与 pnpm 自带的 `pack` 命令混淆。此命令只生成本地文件，不发布包或推送仓库。

## 依赖与状态

已安装 Zod、HTML/CSS 解析器、哈希、ZIP、工具及依赖的版本、许可证见[技术选型](technology-selection.md)。XML 已通过内部 saxes 适配器接入，Tiptap 等依赖仍按后续任务引入；Office OOXML 读写和样式映射自行实现，不提前提供占位功能 API。

修改后运行相应的窄范围检查与 `git diff --check`。只有验证所覆盖的能力才可在 README 标为已实现，产品范围和宿主边界以[总体架构](architecture.md)及[交互协议](editing-protocol.md)为准。
