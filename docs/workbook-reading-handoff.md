# 第三轮工作簿读取交接

当前任务是完成并合并 [PR #25](https://github.com/RobotSe7en/dagent-office/pull/25)，分支 `feat/workbook-reading`。完整范围和退出条件以 [iteration-3.md](iteration-3.md) 为准；本记录不降低门槛。R3-09～16 编辑计算、R3-17～20 输出验收须在读取阶段合并后顺序实施。

## 授权与约束

用户已授权阶段 PR 的提交、推送、单独复审、修复复核和最终提交全部 CI 成功后的合并；无需再次询问合并许可。不直接提交 main、不发布 npm。用户于 2026-11-06 明确允许本次 PR #25 使用子智能体独立审查，覆盖此前禁令，仅限本次审查；后续阶段仍不得自行派生。产品语义出现新分歧时主动询问；已定选择继续有效，不重复询问。

基础七项选择见 AGENTS.md。额外七项为：财务下划线铺满单元格；隐藏列内容按未隐藏宽度参与 AutoFit；冲突经典备注框以 VML 为准且保留两份来源及差异；旧 lineChart/marker 只保留不控制显示；坐标轴按 Office 的跨越轴语义；渐变随形状刚性旋转并保留原覆盖；负柱反转保留符号方向，填白色并保留原边框。

Playground 服务监听 `0.0.0.0:5174`，`http://192.168.31.221:5174/`。5173 属于另一个项目，勿操作。当前任务期间保留 5174 服务。

## 已实现的读取范围

- 公开原生导入、统一 `openDocument`、严格工作簿包、保存重开、专用 Worker 与有界读取通道。
- 单元格类型、富文本及精确 UTF-16；日期／地区及常用数字格式；高级填充、边框、旋转、竖排、命名样式关联。
- 隐藏行列、冻结合并、RTL、跨空白格溢出、跨格居中、实际来源自动行高及隐式字体来源。
- 智能表、基础条件格式及七类验证；当前源求值和明确不可用诊断；打印元数据保留。
- 三类基本图表：当前引用或唯一原生常量，轴、标记、标题、实际字体／数字字段、标签、图例／引导线和极端值裁剪。
- 图片、基础形状／文本框、三种锚点、裁剪／翻转／旋转、渐变覆盖及源投影。
- 经典备注和现代线程、可见备注框及面板、VML 优先、异常定位 needs_reanchor、提及及 uid 保留。
- 公共只读预览、Playground 的打开／取消／切换、资源与字体诊断和释放。

实现细节与独立来源见 [workbook-model.md](workbook-model.md)、[读取验收](workbook-reading-acceptance.md)和 `docs/evidence/xlsx-*.json`。不要把公式读取内核称为 R3-10 的完整依赖计算或编辑能力。

## 最近提交与已完成检查

- `87e4091`：完整标签布局及白色负柱。
- `e64dd2c`：极端图表值域／描边裁剪、72pt 图例预留；修复旧图例选择器及通用 XSD 扫描器对已知非法生产者的分类。
- `b58b46d`：图表实际测量数字字段与全部公开读取矩阵。现代 111 项、旧 Chromium 37 项通过；包括全部 16 源／16 实际生产者文件，四引擎共 128 份公开读取报告。`pnpm check` 的 142 文件／1734 测试、类型、边界和完整构建通过。
- `a10819a`、`3092fcc`：GitHub runner 的 Azure Ubuntu 镜像下载极慢；浏览器和 Office 作业改用官方 archive/security 镜像，配置下载超时／重试。测试与性能门槛不变。
- `3092fcc` 的 CI 工程、Office 重开及 Safari 已通过，浏览器任务仍在进行。所有上述检查不是当前未提交运行时代码的最终通过记录。

主矩阵 32 份来源通过官方 Transitional XSD 与 OPC 审计。额外标签探针中 4 份 Calc 图表存在 numFmt/dLbl 子元素顺序错误：保留原文件和精确错误，仅在验证副本检查其余结构，不能称原文件合规。Calc 还会扁平化线程、丢弃部分图片／渐变开关和改写标签／负柱；生产者差异不能用原源偷偷修复。Excel/WPS 未实机验证。

## 当前未提交批次与剩余工作

最终性能复核暴露旧 Chromium 滚轮 p95 略超 100ms、Firefox 首屏 30.160s。用户已将读取滚轮 p95 调整到 120ms，并明确要求保留并完成优化；其余目标不变。失败和被撤回实现的完整样本保存在 [失败记录](evidence/workbook-reading-budget-failures.json)；带 CPU profile 或改变输入方式的诊断不计为验收。

当前改动：

1. 只对已授权预读来源，AutoFit 最多重叠当前与下一批回复；每批仍不超过 4096 查询槽，累计预算、原子发布和取消不变。Firefox 首屏在该改动后实测 25.748s。
2. 根据上次真实垂直位移准备一个有界候选网格，正式请求仍重新读取并逐项核对来源和字体版本。隐藏候选与当前网格共用节点／片段额度，先完成浏览器布局，接受时复用。
3. 网格以 `overscroll-behavior:none` 保持原生滚动边界，采用被动滚轮监听；单元格滚动仍映射到受校验的视口，备注面板继续使用原生滚动。可中断整窗及其他未达标实验已撤回，不作为当前实现。
4. 新增同修订值变化、非法 HTML、错修订、字体变化、总 DOM 额度、预读关闭、预取取消和迟到清理检查。现代三引擎 12 项、旧 Chromium 4 项通过；额外备注面板的实际原生滚动断言已在旧 Chromium 通过，现代三引擎随最终矩阵复核。
5. 资源包装器记录客体系统报告的 CPU 拓扑。最终四引擎统一使用 `AllowedCPUs=0,16,32,48`、CPUQuota=400%、8GiB、无交换；这是四个逻辑 CPU，不声称独占物理核。先前 0–3 位于同一客体 core 报告，历史失败仍保留。

当前四引擎资源检查及汇总已通过，滚轮 p95 为 49.3／29／48／90.3ms，首屏、内存、DOM、真实计数及取消全部达标。必须继续到完成：刷新当前公共包的 128 份矩阵；完整工程检查；从 main 基线完成三份子智能体独立审查，修复发现并复核；同步 README／AGENTS／第三轮任务状态／验收记录；提交推送。最终 HEAD 的 check、browser-interaction、office-reopen、safari 全部通过后，将 PR 转 ready 并以 expected-head SHA 守卫 squash merge。不得凭旧提交或单一子集成功合并。

## 验证入口

```sh
pnpm build:packages
pnpm exec tsc --project tsconfig.tools.json --noEmit
node scripts/playwright.mjs test workbook-preview.spec.ts --grep 'predicted viewports|idle neighbor|automatic row reads'
node scripts/collect-workbook-reading-matrix.mjs --record
node scripts/collect-workbook-reading-budget.mjs --record
pnpm check
```

资源命令见 [workbook-reading-budget.md](workbook-reading-budget.md)，各浏览器依次独占运行，不与构建或其他浏览器测试争用。运行浏览器期间不改写 dist。普通功能与实际性能不可互相替代。

固定 Office 为官方 LibreOffice 25.8.1.1，归档 SHA-256 `f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c`。本机 `.cache/xlsx-venv/bin/python` 提供 lxml；默认 python3 未装 lxml。正式来源与证据不得被诊断脚本重写。

## 子智能体独立审查修复批次（当前工作区）

用户已明确授权本次子代理审查；三份只读审查覆盖 core、office、render/Worker，并提出 8 项 P2（无已确认 P1）：

1. 规则／图表求值的批次文本消耗未累计：scalar 结果新增内部 `codeUnits` 与剩余额度，两个消费者统一累计；响应不增加字段。
2. 全局名称在联合／交集／范围中丢失歧义检查：沿已有 AST 传递并检查名称上下文，不新增公式引擎。
3. 相邻行压缩漏比较 autoHeight：补齐比较，保留显式自动覆盖及相邻固定默认行。
4. 图表 XMLNS 被误当业务属性：统一仅忽略命名空间声明，保持其他属性严格校验；覆盖 XLSX/PPT 的局部前缀及缺省命名空间。
5. 分页符 max 缺省值漏读：按官方 XSD 使用 0，保存重开保留。
6. AutoFit 等待忽略取消的下一批时不释放原型：等待与 AbortSignal 竞速，外部取消后在迟到回复释放前清理。
7. 合法 viewport 字体 resize 测量使用旧尺寸：保存 providers，viewport/container 单位、CSS 转义、条件 CSS 或辅助 HTML 行内样式在独立候选 surface 重建；支持连续 resize 取代、失败保留，无整本序列化。
8. Firefox 备注面板原生滚动后网格滚轮带动宿主：单纯 CSS containment 不足，已恢复处理网格轮事件的非被动 preventDefault，CTRL／备注面板先排除；隐藏预布局和 AutoFit 预取优化保留。

三份审查已只读复核并全部关闭上述发现。审查代理名称为 `pr25_review_core`、`pr25_review_office`、`pr25_review_render`；不得将之前本人复审称为子代理审查。

修复前新 Office 5 项及 core 3 项均复现失败；当前新 Office 7 项、core 3 项及图表／相关单元合计 37 项通过。现代三引擎最新 15 项通过（`.cache/pr25-review-typography-target.log`），包含 CSS 转义、HTML 字符引用行内样式、连续 resize 取消、迟到回复前清理、宿主与备注面板原生滚动。工具类型检查正在 `.cache/pr25-review-tools-types.log` 进行；最新源码构建通过。

**旧的最终记录须重做：** 已提交 `bf4d294` 的四引擎性能记录与 142/1734 工程检查曾通过，但以上修复改变运行时；必须用当前最后源码重新跑四引擎资源汇总、全量工程、128 公开矩阵及最终 HEAD 四项 CI。上一轮现代矩阵 96 来源全部通过，额外 Firefox 面板/宿主边界失败，因此该轮整体失败，不能作为当前完整通过。当前未有其他浏览器运行；下一步先旧 Chromium 9 项定向复核，然后提交推送修复批次并进行最终各门槛。当前新工作区还未提交，勿恢复旧实现或重新试阈值。

追加取消复核发现 rules.source.read 同源路径，以及共享取消 Promise 保留已完成回复的风险。已提取 read-cancellation.ts 单次可释放 helper，同时用于 AutoFit cell/rules，两类来源的迟到回复均不阻止清理；现代三引擎追加 9 项通过，render 审查代理再次关闭该项。此小批正在提交；之后不要再更改运行时，直接完成最终完整检查及四引擎／128 矩阵。
