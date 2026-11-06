# 工作簿读取资源检查

这是 [R3-08](iteration-3.md) 的显式性能检查入口。读取门槛仍以完整计划为准；脚本存在或普通 CI 成功不表示性能通过。

[当前四引擎记录](evidence/workbook-reading-budget.json)对应 `7ec49aa` 的运行时代码，四组均使用内核四核、8 GiB、无交换和相同字体；汇总器已核对源码及 Worker／宿主包摘要。它仅关闭这一版本的读取性能子集，后续运行时代码变化后须重新核对，完整 R3-08 仍未通过。

| 浏览器 | 首屏（秒） | 选择 p95（ms） | 滚动 p95（ms） | 最大片段 | 浏览器高水位合计（MiB） |
| --- | ---: | ---: | ---: | ---: | ---: |
| chromium | 14.620 | 12.6 | 62.3 | 336 | 984.7 |
| firefox | 28.187 | 6.0 | 31.0 | 300 | 1300.3 |
| webkit | 17.189 | 19.0 | 65.0 | 336 | 1189.2 |
| legacy-chromium | 21.973 | 12.8 | 98.2 | 336 | 843.0 |

旧 Chromium 的 98.2ms 余量有限，Firefox 首屏 28.187 秒也接近 30 秒目标；保留所有慢样本和失败历史，不据一次通过扩大目标机器范围。

先构建项目并安装仓库固定的 Playwright 浏览器及 `scripts/install-legacy-chromium.mjs` 固定旧版本。Linux 主机须有 systemd、cgroup v2 的 CPU／cpuset／内存控制器和既有固定字体；检查不会下载字体。选择主机实际可用的四个逻辑 CPU，将下面的用户／组和 Node 路径替换为当前用户及 Node 的绝对路径。临时系统 unit 由管理权限创建，测试进程仍以普通用户运行；不更改持久服务。每次用不同的 unit 名称，四组浏览器依次运行，避免争用预览服务。

```sh
pnpm build
sudo -n systemd-run --wait --pipe --collect \
  --unit=dagent-workbook-budget-chromium \
  --uid=danny --gid=danny \
  --working-directory="$PWD" --setenv=PATH="$PATH" \
  -p MemoryMax=8G -p MemorySwapMax=0 -p CPUQuota=400% -p AllowedCPUs=0-3 \
  /absolute/path/to/node \
  scripts/workbook-reading-budget.mjs --browser chromium
```

`--browser` 也接受 `firefox`、`webkit`、`legacy-chromium`。包装器在启动测试前检查实际 `memory.max`、`cpu.max`、`memory.swap.max`、`cpuset.cpus.effective` 和 CPU affinity。浏览器子进程继承该组，采样时复核进程及线程没有逃出该集合，允许线程主动收窄；条件不符会失败。SDK、测试、预览服务和浏览器共同受 8 GiB、四 CPU 配额及内核四核集合约束。只有 `taskset` 的用户 unit 不够：WebKit 的辅助线程曾自行扩大 affinity，所以现在必须由 cpuset 强制限制，不能忽略该失败。

实际来源是 [一万行 XLSX](../examples/xlsx/15-ten-thousand-rows.xlsx)。使用独立空宿主页加载公开 SDK，避免移除 Playground 的 React DOM 后仍留有示例异步任务；Playground 功能另有完整浏览器套件。专用 Worker 从原生字节导入，公开预览通过读取通道获取窗口；计时后独立遍历确认 20 万实际单元格、2 万公式、40 块。首次打开使用新的浏览器、Worker 和模型，不复用文档缓存。字体文件在运行前读取并计算摘要，会预热操作系统字体文件缓存；报告明确记录这个条件。首屏计时从原生字节交给 Worker 开始，覆盖导入、模型检查、自动行高和首次窗口提交，排除浏览器／测试服务启动及 SDK 模块加载。

交互通过浏览器真实鼠标事件发送，采集 30 次选择和 40 次正反滚动。起点采用事件时间戳，并记录送达延迟；滚轮监听器等待 SDK 登记候选后，再等待 `settled()` 和下一动画帧，确认视口确实改变。完成点是浏览器下一帧回调机会，未独立测量 GPU 像素呈现。不能把捕获监听器中的微任务当成完成，否则可能在 SDK 气泡处理器执行前计时结束。p95 使用排序后第 `ceil(n × .95)` 项，不丢弃慢样本。失败也保存全部样本。窗口为 1280×800，网格为 880×450 pt，计数覆盖实际已挂载的全部窗格及视口外文字来源，并与公开 `cellFragmentCount` 交叉核对。另用新 Worker 检查导入取消和已接受画面保留。

输出位于 `.cache/workbook-reading-budget/<browser>/`：`browser/**/workbook-reading-budget.json` 保存时延、源码／测试包摘要、实际计数、取消结果及用户代理；`resources.json` 保存 CPU／内存约束、字体文件与浏览器二进制摘要、采样和进程高水位。浏览器二进制摘要在子任务完成后读取，不计入交互时间；字体摘要在启动前读取。普通检查不会修改已提交的证据。

四组浏览器都完成后，可显式运行 `node scripts/collect-workbook-reading-budget.mjs --record`。它重新构建测试 Worker／宿主，核对四轮包摘要与当前代码一致，重新校验全部目标、样本数量、CPU／内存及共同字体，才写入 `docs/evidence/workbook-reading-budget.json`；任何一组失败或代码不匹配都会拒绝更新。性能子集仍不代替完整 XLSX 阅读门槛或独立视觉矩阵；后续实现变化后应以最终代码重新验证。

内存每 20 ms 从同一 cgroup 的浏览器进程读取 `/proc` RSS 和内核 `VmHWM`，Worker 位于渲染进程内。所有已观察进程的高水位相加会重复计算共享页和不同时刻的峰值；浏览器绝对合计小于 2 GiB，比文档增量小于 2 GiB 更严格。采样可能遗漏极短进程，因此另记内核 cgroup `memory.peak`；该数值还含 Node、预览服务和文件缓存，不等于浏览器内存。不得将这两种数值混用。报告注明采样限制，不能声称采样就是无遗漏的分配跟踪。

性能用例关闭 Playwright 的逐操作 DOM 快照、截图和录像，避免诊断采集干扰原生输入时延；功能用例仍保留原有失败追踪。可选 `DAGENT_WORKBOOK_CPU_PROFILE=1` 采集 Chromium CPU profile 及有界的原生样式／布局追踪，只供定位热点；带这些诊断的结果不作为最终性能记录。

本检查显式开启，普通浏览器套件不自动在未受限主机上运行它。最终记录必须同时满足首屏、选择、真实滚动、DOM、内存、取消和既有功能检查；性能失败不能靠放宽断言或减少数据转为通过。
