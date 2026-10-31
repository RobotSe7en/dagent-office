# Presentation 文档包样例

这是 R2-02 的公开合成 HTML 文档包，供 core 的打开、查询、严格校验、资源摘要和保存重开验收使用。它不是独立 PPTX 输入样例，不代替 R2-03 的独立 ZIP/XML 生成器或 R2-06 的视觉验收。

包内包含一页幻灯片、主题和三种布局定义，以及文字、圆角矩形、嵌套在组中的对象、连接线、表格、柱状图数据、图片、演讲者备注和带富文本回复的现代批注。图片复用本项目 [fidelity 样例](../fidelity/README.md)的合成 PNG；不包含用户文档。

文字和顺序保存在 HTML，外观保存在 CSS，图表数据及定义引用保存在 metadata.json。图表 SVG、页面画布和原生 XLSX 不保存在包中。当前读取 API 见 [Presentation 模型](../../docs/presentation-model.md)，完整第二轮范围见[实施细则](../../docs/iteration-2.md)。
