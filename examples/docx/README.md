# DOCX 读取样例

本目录只含本项目制作的公开合成文档，不含用户文档。`01` 至 `08` 覆盖第一轮读取范围，`09` 专门检查中西文字体、扩展区汉字、字体替代和缺字提示。生成脚本为 [generate-docx-fixtures.py](../../scripts/generate-docx-fixtures.py)。

`expected.json` 给出独立内容预期；`producer/` 是 LibreOffice 实际重存的 DOCX；`reference/` 是原文件在固定字体环境下的参考页面。环境、输入和图像摘要见 [reference.json](reference.json)，连续预览回归另存于浏览器测试目录。

| 文件 | 验证内容 |
| --- | --- |
| 01-text-numbering.docx | 中文、行内格式、标题、层级编号和重启 |
| 02-table-merges.docx | 横纵合并、边框和表格网格 |
| 03-images.docx | 三种图片、包内关系和图片尺寸 |
| 04-styles-sections.docx | 样式继承、直接覆盖、分节及页眉页脚 |
| 05-long-page-break.docx | 长内容、明确分页及连续预览 |
| 06-links-tabs.docx | 书签、链接、制表位和引导符 |
| 07-fields.docx | 简单、复杂、嵌套及跨段落字段缓存 |
| 08-comments-revisions.docx | 经典富文本批注、范围及文字插入/删除修订 |
| 09-chinese-fonts.docx | 中西文字体分离、全角标点、缺失字体和 U+20000 缺字诊断 |

参考应用可能更新字段或使用不同的修订显示模式；字段缓存和修订内容按实际源 XML 另行检查。参考页不代表浏览器已有自动分页或 DOCX 导出。完整运行方法和限制见 [DOCX 读取说明](../../docs/docx-import.md)。参考截图使用本机已有字体，字体二进制不随本目录分发。
