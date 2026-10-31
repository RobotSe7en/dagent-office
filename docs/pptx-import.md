# PPTX 基础读取

R2-04 提供 `importPptx(bytes, { documentId, limits?, maxZipBytes? })`，将宿主授权的 PPTX 字节映射到同一 HTML/CSS/元数据／资源文档包。它不调用 Office 进程、网络或第三方演示文稿库。当前覆盖基础对象；图表、备注和两类批注在 R2-05 实施，全部读取和只读画面验收属于 R2-06。编辑入口继续关闭。

```ts
import { openDocument } from "@dagent-office/core";
import { importPptx, writeDocumentZip, readDocumentZip } from "@dagent-office/office";

const files = importPptx(authorizedPptxBytes, { documentId: "host-id" });
const model = openDocument(files);
const snapshot = model.snapshot();
const packageBytes = writeDocumentZip(model.serialize());
const reopened = openDocument(readDocumentZip(packageBytes));
```

输入字节保持不变。文档包不包含原 PPTX 或隐藏的转换回退副本；输出可直接打开、查询、保存重开。ZIP、XML、生成节点与字节、资源和最终源校验均有预算。DTD、实体、危险 XML、外部图片、动画、未知可见对象和失效关系明确失败，不提供默认部分导入。

| 内容 | 当前映射 |
| --- | --- |
| OPC | 按关系和内容类型定位，支持改名／移动部件、根绝对关系、Strict／Transitional 和任意 XML 前缀；不按 `ppt/slides/slide1.xml` 猜路径 |
| 文稿 | 宿主身份、标题、语言、原始 pt 尺寸、幻灯片顺序、隐藏及母版显示标志 |
| 主题和继承 | 颜色、角色字体及补充脚本字体、格式样式列表、颜色映射、母版文字样式、布局与占位符身份；缺失几何继续通过源引用继承 |
| 文字 | UTF-16、emoji／变体序列、字符／段落格式、列表、换行、横排、内边距、垂直对齐、三种自动适配、默认制表间隔及文字／对象超链接 |
| 形状 | 声明的 16 种预设、常量调整值、实色／线性渐变、透明度、轮廓、虚线和基础外阴影 |
| 组与连接线 | 局部坐标、负坐标、嵌套、旋转、翻转；`line`／`straightConnector1` 及 `bentConnector2` 至 `5` 的原始预设、调整值和对象引用 |
| 图片 | PNG／JPEG／WebP 原始字节、摘要、矩形裁剪、变换和替代文字 |
| 表格 | 原生 DrawingML 网格、行高、横纵合并、显式填充／边框／对齐和富文本；合并起点在 HTML 中只有一份内容 |

未支持的表格样式 ID、显式制表位、垂直／多栏／变形文字、文字描边／复杂效果、任意自由曲线、图案／纹理填充、非矩形图片、音视频、嵌入控件、宏及高级演示特性均明确失败。空的静态 transition 容器没有过渡效果，可读取；带效果或自动翻页时序的容器失败。PowerPoint／WPS 应用兼容结果尚未记录。

外观唯一保存在 CSS：主题原色为 `--office-palette-*`，有效颜色角色为 `--office-theme-*`；颜色映射引用原色，避免交换角色时产生变量循环。主要／次要字体保存在 `--office-major-*`／`--office-minor-*`，字符字体分 ascii／h-ansi／east-asia／complex-script。格式样式列表映射为主题 CSS 类；对象引用相应类并保留局部覆盖。颜色变换、渐变停止点、线条和阴影使用 `--office-*` 自定义属性，不保存原生 XML 作为第二份编辑来源。文字样式、列表和字段缓存保持源语义，画面适配在 R2-06 读取这些属性。

位置和组子坐标仍由 `data-office-x/y/width/height` 等源属性保存。原生预设常量为 `data-office-adjust-*`，连接线原始类型为 `data-office-connector-preset`；核心校验检查其类型、数值和 routing 一致性。幻灯片内原生 ID 保存在类型化对象元数据，用于连接引用和后续批注映射。它们只表示身份，不授予访问权限。

验收使用[独立输入与 LibreOffice 重存文件](../examples/pptx/README.md)。基础读取用例对 11 份源及对应 11 份外部生产者文件分别验证文字、对象和引用、输入不变、确定性以及文档包 ZIP 保存重开；另覆盖图片原始字节、主题与占位符、嵌套坐标、连接引用、合并、异常输入及预算。LibreOffice 重存可能改变连接线预设和小数网格尺寸，读取保留重存文件的实际值，不改回原合成值。现代批注在生产者重存中丢失，不使用无批注重存文件代替现代读取验收。

```bash
pnpm build:packages
pnpm exec vitest run packages/office/tests/pptx-import.test.ts
```

此阶段的重开是 HTML 文档包重开，尚未提供 PPTX 导出；原生可编辑导出和外部编辑后的重新导入在 R2-14 至 R2-16 验收。
