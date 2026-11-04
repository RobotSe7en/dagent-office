# R3-02 工作簿存储技术基线

本阶段验证一万行来源的分块读取、紧凑索引、公式模板共享及 Worker 窗口传输。本页记录技术原型证据；后续 [R3-03](workbook-model.md)已将其接入公开的新建、统一读取和窗口查询。完整 XLSX 读取门槛、编辑、求值和打印尚未完成，任务范围见[第三轮细则](iteration-3.md)。

## 来源与常驻状态

独立生成器 [workbook-baseline-fixture.mjs](../scripts/workbook-baseline-fixture.mjs) 产生 1 万行 × 20 列，20 万个实际单元格、2 万个公式，含数字、中文、emoji 和变体选择符。每块最多 256 行 × 32 列，包内保存 HTML 字节；原型每块保留一份私有字节、排序坐标、UTF-16 起止偏移和公式模板编号。解析树及逐单元格临时对象在校验后释放。返回窗口只切出请求单元格，序列化返回字节副本。

原型严格检查块身份、行列范围和次序、重复单元格、显式单元格边界、值类型、UTF-8、HTML 标签/属性及预算；未建模内容拒绝。样式、资源、公式名称解析和跨块引用将由正式工作簿模型继续检查。当前块不能独立充当完整工作簿验证器。

技术原型首先支持常量、错误、缺省参数、标量运算、函数调用、名称及 A1/跨表/混合引用。相对地址转为模板偏移，2 万公式共享两个 AST；不使用 `eval` 或第三方公式库。解析器限制源码、token、AST 深度及模板预算。后续 [R3-05 引用基础](workbook-formulas.md)已增加整行列、结构化引用、共享公式和引用重写；不能可靠解释的语法明确失败，完整计算仍待 R3-10。

解析依据为 [MS-XLSX 公式语法](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3d025add-118d-4413-9856-ab65712ec1b0)及 [Excel 运算优先级](https://support.microsoft.com/en-us/excel/calculation-operators-and-precedence-in-excel)。未知函数可以形成保留引用的 AST，计算支持与名称解析由后续阶段判断，解析成功不代表可求值。

## 可重复验证

```sh
pnpm probe:workbook-storage
pnpm exec vitest run packages/core/tests/workbook-chunk.test.ts packages/core/tests/workbook-formula.test.ts
pnpm build
node scripts/playwright.mjs test workbook-storage.spec.ts
```

Node 报告写入 `.cache/workbook-baseline/node.json`。浏览器报告作为各引擎 Playwright 附件及输出文件保存。固定生成源的 SHA-256 为 `bc1454f1a04f735837f6a6d9500caa286e6c8266d43117d17d0f658ee7621270`，来源 16,506,794 字节，紧凑索引 2,120,000 字节，公式模板键合计 436 个 UTF-16 单位。

本地记录见[测量证据](evidence/workbook-storage-baseline.json)：Node 来源生成、校验和公式模板构建总计约 2.4 秒，采样 RSS 增量约 265 MiB，进程生涯峰值约 325 MiB。此环境没有固定为最终 4 vCPU / 8 GiB 配置，数字只说明该原型的实测状态，不作为最终性能验收。

38 个单元测试及 Chromium/Firefox/WebKit 的 Worker 用例通过。Worker 测试确认实际单元格与公式数量、模板共享、窗口回复小于 10 KB、宿主事件循环持续运行，以及取消新载入后仍可读取已打开数据；测试结束释放 Worker 和 Blob URL。浏览器总内存、SDK 会话历史共享、原生 XLSX 打开和完整计算仍需后续门槛测量。

## 下一阶段

R3-03 将把此技术路线接入严格契约、分块清单、新建和只读查询；R3-04/05 增加独立 XLSX 来源和完整引用语义。原型不提前开放工作簿编辑，不把导入缓存当作已计算结果。公开 API、能力矩阵和性能检查在对应任务完成后再声明支持。
