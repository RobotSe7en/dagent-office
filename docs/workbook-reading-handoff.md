# 第三轮执行交接

更新于 2026-11-05。本文件记录执行位置和下一步；完整范围、产品决策和退出条件以[第三轮实施细则](iteration-3.md)为准，具体已实现行为见[工作簿模型](workbook-model.md)。不要把局部测试通过解释为完整读取验收通过。

本次续做先检查了 `050287b` 的 CI #62：`check`、`office-reopen`、`safari` 已通过，浏览器项仍在运行。首批补齐表／表列区域的命名单元格样式关联及 DXF 内部边框保留，22 个定向用例通过；不代表表样式显示组合完成。已向用户发起两项显示语义选择（隐藏后的行条纹、直接字体按属性或组件覆盖），答复前不实施依赖这些选择的组合。

首批代码 `6cabf4f` 已推送，额外通过 32 份公开导入／重开回归。续做发现共享 DrawingML 列表的零缩进和零左边距被 `parseFloat(...) || default` 覆盖，现已修复，三种现代浏览器和固定旧 Chromium 的原生源／Calc 源绘图文字布局检查共 4 项通过。CI #62 随新推送取消，后续以新 HEAD 检查结果为准。

零缩进修复 `028e99e` 已推送。另已按规范与 24 项独立 Calc 探针实现 `anchorCtr` 的最小文字块居中，保留左右／混合段落对齐及上中下锚定；四引擎逐字边界小于 1.5pt、6 项绘图源测试及全部附加 XLSX 探针的官方 XSD／关系审计通过。锚定探针可用 `python3 scripts/check-xlsx-text-anchoring.py` 重核，不计作整页验收。

## 新 session 从这里继续

1. 阅读仓库 `AGENTS.md`、`README.md`、[第三轮细则](iteration-3.md)及本文件；实现相应模块前再读其设计文档。
2. 检查 `git status --short --branch` 和当前 PR/CI；保留任何新出现的用户改动，不重置工作区，不切到 main 直接提交。
3. 继续完成 PR #25 的 R3-06～08 读取门槛。下一项实质实现是智能表样式的显示组合，随后是条件格式和数据验证提示；下文列出具体缺口。
4. 每完成可验证的一批就提交、推送并更新状态，不再长期积累本地提交。合并前对完整差异独立审查、修复复核，确认最终提交的四项 CI 全部通过。
5. PR #25 合并后才进入 R3-09～16 编辑计算 PR，再进入 R3-17～20 输出与独立闭环 PR。用户授权覆盖全部阶段，不需要重新请求常规实施或合并许可。

可直接给新 session 的任务：

> 请读取 AGENTS.md、README.md、docs/iteration-3.md 和 docs/workbook-reading-handoff.md，按已定计划继续第三轮。从 feat/workbook-reading / PR #25 的当前提交继续，先检查 CI，再完成剩余读取门槛；不要重新规划、缩减范围或重复已完成工作。每批验证后及时提交推送，最终独立审查、修复、检查全部 CI 后合并，再继续后两个阶段。产品语义或范围需要改变时主动问我。

## 分支、PR 和最近提交

- 仓库：`/home/danny/work/dagent-office`。
- 分支：`feat/workbook-reading`。
- 当前 PR：[RobotSe7en/dagent-office #25](https://github.com/RobotSe7en/dagent-office/pull/25)，仍为草稿，覆盖 R3-02～08；读取门槛未通过，不能合并。
- 本交接记录之前的代码提交为 `66662fb0f1a45d6054d951d865e8a9a7e4f16846`，已推送。交接文档另有后续提交，恢复时以实际 HEAD 为准。
- 最近三笔代码提交：`998d6c1` 修复 Word/PPT 多文字范围批注回归；`29c1af5` 应用富文本数字格式；`66662fb` 接入 Playground 的所属 Worker 和只读预览。
- 更早的关键提交：`73789bd` 公共原子预览、`e6752cd` Worker 只读传输、`7c7c127` 全部 60 个规范表样式、`ece17b6` 十三类表区域与边框投影。
- 五个阶段 PR 中，#23 工作台和 #24 路线/设计已合并；#25 进行中；编辑计算和输出两个 PR 尚未建立。
- 代码提交 `66662fb` 的 [CI #61](https://github.com/RobotSe7en/dagent-office/actions/runs/42002466746) 在交接时运行中，四项为 `check`、`browser-interaction`、`office-reopen`、`safari`。新的文档提交会触发后续 CI；新 session 应查询最新 HEAD，不能沿用旧提交的检查结论。

## 已定决策和授权

既有决策已写入[细则](iteration-3.md)；续做中新增的待答显示边界见本文件开头，答复前不改变相关语义。特别保留：

- 统一 `openDocument`，按 `kind` 提供快照与查询，不增加公开 `openWorkbook`。
- 基线是 **1 万行 × 20 列、20 万实际单元格、2 万公式**；不是十万行，也不能用空坐标或资源保护上限代替实际容量验收。
- 排序区域外公式保持地址引用；区域内公式随记录按引用规则调整，批注跟随数据。
- 剪切引用跟随数据，范围被拆开时精确改写为多区域，不使用包围矩形替代。
- 普通文本与富文本统一应用文本格式；保留局部样式，格式显式颜色覆盖整段，派生前后缀不写回原值。
- 高级单元格外观、命名样式、智能表、基础规则、三类图表、绘图和两类批注都在本轮范围，不得以拒绝导入代替已选实现。
- 现代线程定位矛盾时保留正文/回复并标记 `needs_reanchor`；原生图表常量是唯一来源，引用图表读取当前单元格；负值饼图按绝对值绘制并保留带符号标签和诊断。
- 日期兼容 Excel 1900/1904；地区依次取格式、宿主、文档、zh-CN。
- 用户已授权提交、推送、独立审查、修复及合并所有阶段 PR。不得直接提交 main、发布 npm 或使用子代理。涉及产品语义改变或缩减范围时主动询问。
- R1-15 继续延期，R1-16 宿主验收未完成；Excel/WPS 仍未实机验证。外部自动验证固定 LibreOffice 25.8.1.1。

## 已实现的读取基础

- 工作簿严格契约、256×32 分块、紧凑索引、公式解析/重写/结构化引用基础、统一打开及保存重开。编辑会话仍明确拒绝工作簿。
- 公开 `importXlsx`，可取消且有预算的 ZIP/XML 管线；32 份独立源/Calc 文件可导入重开，17 类负例整包拒绝。公式缓存状态与真实重算结果分开。
- 公共只读预览：有界窗口、冻结/合并/RTL、选择、工作表导航、地址/公式栏、绘图、只读批注面板和源投影。候选失败或迟到不替换已接受窗口。
- 宿主所属 Worker 的专用只读端口，按需查询单元格/图片/图表；预览销毁不终止宿主 Worker，连接和宿主分别拥有自己的生命周期。
- Playground 原生 XLSX/工作簿 ZIP 打开、阶段进度、取消、首屏成功后原子切换、失败保留旧视图、缓存 Worker 复用及卸载释放。Word/PPT 编辑历史隔离。仍为只读试用。
- 字体角色、命名样式、数字/日期和文本格式、图案/渐变、斜线、旋转/竖排的读取和显示基础；普通/富文本格式组合已按用户选择接通。
- 三类图表当前值读取、常量和引用分离、混合类型、隐藏点、标题及基础布局；图片/形状/批注也已有窗口组合。细节和完整原生视觉验收仍有缺口。
- 全部 60 个规范表样式的数据目录，以及十三类表区域、条纹和边框投影；**尚未接通单元格的最终显示组合**。

## 当前 PR 剩余实施顺序

### 1. 智能表、规则和单元格排版

- 接通表样式、直接格式、命名样式与条件格式的优先关系；按属性组合，不能用整体 CSS 覆盖蒙混。处理隐藏/筛选后的条纹，表头/汇总/角落和内部/外部边框。
- 实现已选基础条件格式：数值比较、标量表达式、双色/三色色阶，含优先级及停止规则；依赖读取、缓存状态和不可计算诊断要明确，不能把导入旧缓存当成当前求值。
- 数据验证的七类元数据已可读取；补只读输入/错误提示及规则状态。编辑时的统一执行属于后续 R3-09～16，不在读取阶段提前开放编辑。
- 补普通文本跨空白格溢出、自动行高、财务下划线和特殊对齐；审查区域/年代/日历/数字格式的其余组合和明确支持边界。
- 补视口外旋转文字发现、共享边框与旋转几何的交叠，以及大文字/极端几何的限额行为。

### 2. 图表、绘图、批注及原生视觉

- 补图表字体/间距/标签边界、分类网格、颜色和 `varyColors`、引导线、自动布局及极端数据裁剪；按固定字体对照原生页面。
- 复核图片 `rotWithShape=false`、正文 overflow 属性。`anchorCtr=true` 的文字块居中及 `presentation/text.ts` 中列表零缩进／零左边距已实现并通过四引擎检查；完整排版组合仍待验收。
- 补经典原生备注框外观；可靠处理 VML 与 `commentPr` 同时出现的几何优先关系。复核现代提及/`xr:uid` 的有界保留和已定非通知范围，不能静默丢正文。
- 对全部已选来源形成完整读取/显示矩阵，保留规范、SDK 源重开、外部应用的独立结论；Calc 扁平化或丢失的效果不能算保真通过。

### 3. 固定资源验收、独立审查和合并

- 固定 **4 vCPU / 8 GiB**、浏览器与字体；实际一万行数据的 XLSX 到可交互首屏 ≤30 秒、交互 p95 ≤100 ms、可取消；1280×800 所有窗格单元格 DOM 合计 ≤5000；浏览器和 Worker 增量峰值 ≤2 GiB。
- 已有 `docs/evidence/workbook-native-worker.json` 是本机四引擎功能记录，环境并非固定 4 vCPU / 8 GiB，也未测峰值内存，不能据此宣称性能门槛通过。
- 全量计算 ≤15 秒、编辑 p95、排序/历史等门槛仍归后续编辑计算阶段，保留原编号和目标。
- 对完整 PR 独立审查、修复和复核，再在最终 HEAD 跑完四项 CI。完整读取门槛通过后才能合并 #25、进入下一阶段。

## 下一项：表样式组合的定位线索

- `packages/render/src/workbook/table-style-data.ts`、`table-styles.ts`、`table-style-types.ts`：全部规范目录、只读查询和差异格式类型。
- `packages/render/src/workbook/table-regions.ts`：十三类区域的规范优先级、尺寸为 1～9 的条纹、边框槽映射；不会展开整表，RTL/视口不重置源区域。
- `packages/core/src/contracts/workbook-tables.ts`、`workbook-cell-styles.ts`：表与命名/原生单元格格式契约。
- `packages/office/src/xlsx/styles.ts`、`tables.ts`、`rules.ts`：原生样式/表/规则读取；原生完整 XF CSS 与部分 DXF CSS 的语义不同。
- `packages/render/src/workbook/cell-source.ts`、`window-layout.ts`、`cell-borders.ts`：源祖先、实际 CSS 测量和单元格/边框组合入口。
- 当前源格式选择涉及 cell/row/column/Normal，原生 XF CSS 已展平默认字体/填充/边框，不能直接追加一层表 CSS。`applyFont/applyFill/...` 也不能未经证明当成当前直接格式的优先标志。
- `scripts/check-xlsx-table-styles.py`、`docs/evidence/xlsx-table-style-semantics.json`、`packages/office/tests/fixtures/xlsx-table-style-precedence.xlsx` 是独立探针。固定 Calc **没有显示这批探针的表样式**，所以不能据它确认表样式与直接格式优先关系或隐藏行条纹。
- 规范 ECMA-376 Part 1 §18.8.41 给出区域顺序，§18.18.77 有区域图。原附件和 PDF 在本地 `.cache`；正文需要时重新定位，不依赖聊天记忆。
- 本地规范源 `.cache/presetTableStyles.xml`，SHA-256 `74c6cf73d551bb2bd33a51690d4922095260b6647b2a38fcbc0d6da6c6275d7e`。`python3 scripts/generate-workbook-table-styles.py .cache/presetTableStyles.xml --check` 可核对生成目录。

## 最近验证结果与回归修复

以下是代码提交 `66662fb` 前后的本地结果，不能替代下一提交的相关检查或最终 CI：

- 包构建、Playground 生产构建、工具 TypeScript、依赖边界、`git diff --check` 通过。
- 完整单元批次 1582 个用例通过；打包消费套件因 `pnpm exec vitest` 缺少 `npm_execpath` 首次启动失败，随后通过 `pnpm test tests/package-consumption.test.ts --maxWorkers=2` 单独重跑，4 个用例通过。不是尚待处理的代码失败。
- Playground 三项加原生普通/富文本格式组合一项，三种现代浏览器共 12 项、固定旧 Chromium 共 4 项通过。结果目录 `.cache/workbook-playground/{modern,legacy}`。
- 旧 CI #60 的三个 Word 批注浏览器失败用例修复后全部通过，目录 `.cache/workbook-annotation-regression`。根因是无定位 cells 批注联合分支在有界数组预检中误把 text-set 的 ranges 当作坐标；`998d6c1` 保留 kind 分派并增加回归用例，没有放宽通用预算校验。
- 旧图表负例还在拒绝用户已批准的原生常量向量；现已改为接受单一常量、拒绝常量与引用混用。
- 固定 LibreOffice 25.8.1.1 的 DOCX 导出/重开检查通过，结果 `.cache/docx-export-acceptance/result.json`，包含源/生产者及编辑案例。
- 13 项独立文本格式探针、对应原生 XLSX、官方 XSD/关系审计已检查。Calc 对三个富文本案例忽略格式，SDK 按用户选择统一应用，差异已在文档记录，不声称 Excel/WPS 验证过。

## 本地命令和环境

- Node 24、pnpm 10.28.2。先构建包，再运行使用包入口的测试。
- 有针对性的单元检查可用 `pnpm test <test-files> --maxWorkers=2`。不要用不受支持的 `--minWorkers`；打包消费测试要走 `pnpm test` 生命周期。
- 现代浏览器：`node scripts/playwright.mjs test <spec-files> --output=<unique-cache-dir>`。运行前构建 Playground；同一时间只开一个 Playwright/Vite 4173 服务。
- 固定旧 Chromium：同上并加 `--project=chromium`，环境变量 `DAGENT_OFFICE_BROWSER_EXECUTABLE="$PWD/.cache/legacy-browser/chromium-1041/chrome-linux/chrome"`。
- Playground 相关测试：`tests/browser/workbook-playground.spec.ts`；富文本组合在 `workbook-grid.spec.ts`，名称含 `composed native plain and rich text formats`。
- 静态检查：`pnpm exec tsc --project tsconfig.tools.json --noEmit`、`pnpm check:boundaries`、`git diff --check`；完整阶段检查见 CI 和[开发文档](development.md)。
- 固定 LibreOffice 包装器：`PATH="$PWD/.cache/acceptance-office/bin:$PATH"`；依赖路径 `LD_LIBRARY_PATH="$PWD/.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu"`。版本 `25.8.1.1`、build `54047653041915e595ad4e45cccea684809c77b5`。
- 官方 Office 归档 `.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz` 的 SHA-256 为 `f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c`。字体清单 424 项，摘要 `c69980066d071afe8f248c7774f5265a6eb82990db1f2f78843580df8d2325b4`。
- Python lxml 环境 `.cache/xlsx-venv`；XSD/关系审计入口 `scripts/check-xlsx-output.py`。普通检查不改证据基线；只有明确增加来源并审查差异时才用 `--record`。
- GitHub 插件读取 PR/Actions 可用；前次 PR 正文写入曾返回 403，已使用授权的 `gh pr edit --body-file <file>`。多行正文写入文件，避免 shell 插值。

交接时没有运行中的本地构建、浏览器或 Office 验证任务；远端 CI 仍需新 session 查看。后续执行应持续更新本文件的当前状态，已解决缺口及时移除，完整计划和任务编号仍维护在第三轮细则。
