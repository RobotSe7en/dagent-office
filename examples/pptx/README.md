# 独立 PPTX 样例与参考环境

本目录只有公开合成数据，不含用户文档。源文件由 [Python ZIP/XML 生成器](../../scripts/generate-pptx-fixtures.py)依据 [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)、[MS-PPTX](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/) 和 [MS-ODRAWXML](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/) 自行编写，使用 Python 标准库，不调用 SDK 导出器或第三方 Office 库。源 ZIP 固定顺序、时间与存储方式，独立于 SDK 实现。

| 文件 | 原始内容 |
| --- | --- |
| 01-rich-text.pptx | 中文、Latin、emoji、扩展字和变体序列，字形角色、粗斜体、列表、换行及外部超链接 |
| 02-theme-placeholders.pptx | 主题颜色／字体、母版和布局占位符、幻灯片继承几何 |
| 03-masters-layouts.pptx | 多母版、布局及主题引用 |
| 04-shapes.pptx | 16 种预设、填充、渐变、透明度、轮廓、虚线和外阴影 |
| 05-groups.pptx | 嵌套组、局部坐标、负坐标、旋转、翻转和缩放 |
| 06-connectors.pptx | 直线／直角连接、箭头及对象连接引用 |
| 07-images.pptx | PNG、JPEG、WebP、裁剪、旋转、翻转及替代文字 |
| 08-table-merges.pptx | 原生 DrawingML 表格、网格及横纵合并 |
| 09-charts.pptx | 柱状／折线／饼图、缓存及独立编写的内嵌 XLSX 常量数据 |
| 10-notes.pptx | 备注母版和逐页富文本备注 |
| 11-classic-comments.pptx | 经典作者、时间、正文和位置，原生坐标 576 对应 72 pt |
| 12-modern-comments.pptx | 现代作者、正文、回复、解决状态及幻灯片／对象／UTF-16 文字锚点 |
| 13-multi-slide.pptx | 三页组合文稿、隐藏页、图表、备注及两类批注 |
| 15-strict-namespaces.pptx | Strict 命名空间 |
| 16-alternate-prefixes.pptx | 不同 XML 前缀 |
| 17-renamed-parts.pptx | 改名／移动部件与绝对根关系，不能按固定路径读取 |
| invalid/ | 缺件、重复身份、DTD、未支持对象、动画、外部图片／数据、公式、未知作者和损坏 ZIP |

非法与未支持输入一类位于 `invalid/`，当前 16 份合法源、10 份异常输入，覆盖计划中的 14 类；细分边界随读取与编辑实现补充，不代表全部实现已通过。

[expected.json](expected.json)记录独立内容预期及输入 SHA-256。[reference.json](reference.json)记录实际 LibreOffice／Poppler／fontconfig 版本、全部已安装参考／回退字体摘要、源与生产者文件摘要及逐页图像摘要。`producer/` 是 LibreOffice 实际重存文件，`reference/` 是源 PPTX 的 PDF 与 96 DPI PNG。所有源文件保持不变，生产者结果单独保存。

本地参考使用 LibreOffice 25.8.1.1，Impress／Calc／Draw 与原有 Core 同版本，未升级 Word Core。使用已有 Noto Sans CJK SC、Liberation Sans 和系统回退字体，没有安装或下载新字体。Noto CJK 来自系统 `fonts-noto-cjk`（1:20240730+repack1-1，SIL OFL 1.1，`/usr/share/doc/fonts-noto-cjk/copyright`）；Liberation 来自系统 `fonts-liberation`（1:2.1.5-3，SIL OFL 1.1，`/usr/share/doc/fonts-liberation/copyright`）。补充中文字体的来源、版本和许可证沿用[既有环境记录](../fonts/environment.json)。字体二进制不随仓库分发，也不成为 SDK 运行依赖。

已观察到的兼容性限制：

- LibreOffice 25.8 重存会删除现代批注及回复，`sourceFeatures`／`producerFeatures` 明确记录数量差异；现代线程的 SDK 验收必须使用原始 PPTX。经典批注、备注和三类图表的数量保留。
- 现代文字锚点样例的 U+20000 在 PDF 文本中缺失，emoji 的提取顺序变化；独立源保留全部 Unicode 和精确 UTF-16 范围，PDF 只检查明确列出的分段文字。字形可用性需要后续预览诊断，不能据 PDF 提取结果宣称无损。
- 试验中 LibreOffice 对带空格且百分号编码的部件名输出空白页。因此当前外部参考使用移动后的无空格部件名；编码路径另在 SDK 关系解析用例验证，不能据此宣称外部应用支持。
- 此记录只验证外部参考环境，原生 PPTX 导入和只读画面由 R2-04 至 R2-06 验收。尚未验证 PowerPoint／WPS 实际应用，也未验证现代批注在这两者中的编辑。

```bash
pnpm fixtures:pptx
pnpm check:pptx-fixtures
pnpm check:pptx-reference
# 仅在审查过环境或源变化、确实需要更新参考后执行：
pnpm record:pptx-reference
```

独立审计检查可重复字节、关系目标、类型、原生身份、内嵌常量数据和 UTF-16 范围。参考检查使用私有临时 Office profile、超时及进程清理，逐份核对打开、重存、再次打开、页数及可见文字，不信任 Office 的零退出码。每次均检查已提交参考文件的摘要；环境一致时还必须逐像素摘要相等。CI 使用 Ubuntu 24.04 的实际 Office 和字体；环境不同时照常运行内容检查，明确记录 `environment-differs-pixels-not-compared`，不将它标成像素验收通过，也不自动更新基线。结果在 `.cache/pptx-reference/result.json`。
