# 资源与 ZIP

R1-05 在 `core` 实现资源摘要、图片容器识别及宿主资源收集，在 `office` 实现文档包 ZIP 读写。两者接收授权字节，不访问文件系统、网络或模型。R1-10 的 DOCX 导入和连续只读预览见[读取说明](docx-import.md)；R1-06 的[事务与历史](document-session.md)已实现；富文本、自动分页和 Office 导出尚未实现。

## 资源校验与收集

`openDocument` 检查清单中全部资源，包括未被正文引用的声明：缺少字节返回 `AssetUnavailable`，SHA-256 不符返回 `InvalidDocument`。PNG、JPEG、WebP 还检查文件签名及基本容器边界是否符合声明的 MIME，不根据扩展名判断，不重新编码图片。该检查不是完整像素解码或图像解码器安全审计。

其他 MIME 的资源可作为不解释、不执行的附件保存，并校验摘要；当前不能将它们用于 `img`。因此保留 SVG、字体或原始文件附件不代表已支持它们的显示、字体嵌入或 Office 转换。正文图片仍须使用已声明的包内 PNG、JPEG、WebP。样式支持范围仍见[文档模型](document-model.md#htmlcss-支持范围)。

```ts
import { collectDocumentAssets, createAssetDescriptor, openDocument } from "@dagent-office/core";
import type { AssetProvider, DocumentPackage } from "@dagent-office/core";

declare const imageBytes: Uint8Array;
const descriptor = createAssetDescriptor("assets/image.png", "image/png", imageBytes);
// 宿主将 descriptor 写入 manifest.assets；此函数不修改任何文档包。

declare const files: DocumentPackage;
declare const provider: AssetProvider; // 绑定当前主体和会话的授权资源读取接口。
const complete = await collectDocumentAssets(files, provider);
const document = openDocument(complete);
```

`AssetProvider.readAsset(descriptor)` 异步返回字节或 `undefined`。收集器只请求清单声明且当前缺失的资源，顺序读取，在保留响应前检查单文件和总字节预算，再复制字节并校验摘要与受支持图片格式。已有资源先校验，不静默用宿主响应覆盖错误内容。宿主抛出的原始错误不回显。输入 Map、源字节和宿主响应都不被修改，也不在会话之间缓存资源；路径和摘要不授予存储访问权。

同步的 `openDocument` 只接受资源齐全的包。需要从宿主存储收集资源时，先调用异步收集器；不在解析正文时隐式发起读取。`createAssetDescriptor` 返回冻结的声明，保持现有清单 Schema 和格式版本不变。

## ZIP 读取与保存

```ts
import { openDocument } from "@dagent-office/core";
import { readDocumentZip, writeDocumentZip } from "@dagent-office/office";

declare const archiveBytes: Uint8Array;
const files = readDocumentZip(archiveBytes);
const document = openDocument(files);
const savedArchive = writeDocumentZip(document.serialize());
const reopened = openDocument(readDocumentZip(savedArchive));
```

`readDocumentZip` 负责容器校验和展开，返回独立的虚拟文件 Map；随后由 `openDocument` 校验清单、版本、HTML/CSS、引用及资源。`writeDocumentZip` 先检查文档及全部声明资源的完整性，再从复制的源文件字节生成 ZIP，不自动改写 HTML/CSS。需要规范化保存时先调用模型的 `serialize()`。未引用的包文件和资源原字节也保留。

ZIP 压缩解压复用 fflate，本项目适配层在解压前按[ZIP 格式规范](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT)检查原始中央目录和本地记录：

- 路径必须为规范包内路径；不先规范化或清除 `..`，不接受重复记录、文件/目录冲突或本地与中央目录名称不一致。
- 接受 ASCII 和显式标记的 UTF-8 文件名，保留 Unicode 和大小写，不猜测系统代码页；Unicode 路径扩展若存在，必须与原始名称一致。
- 支持单卷经典 ZIP 的 STORE、DEFLATE、目录及数据描述符；检查记录范围、重叠、实际展开长度和 CRC-32。目录不产生虚拟文件，但计入条目预算。
- 加密、多卷、ZIP64、符号链接、其他压缩方法及非 UTF-8 旧代码页返回 `UnsupportedFeature`。这些是当前适配器的能力边界。
- 先检查声明的展开大小，再以小块压缩输入执行流式解压，按实际输出继续检查大小。声明大小被伪造时也不能完成一次不受限的完整展开。

`DocumentZipOptions` 接受 `limits?: Partial<DocumentReadLimits>` 及正安全整数 `maxZipBytes`。后者默认 64 MiB，由 `DEFAULT_MAX_ZIP_BYTES` 导出，同时约束 ZIP 输入与输出；展开后的文件数、单文件及总字节采用[文档读取预算](document-model.md#输入预算)，其中 `maxFiles` 对 ZIP 包括目录记录。超限返回 `ResourceLimitExceeded`；格式损坏返回有界的 `InvalidDocument`，不回显文件内容或解析器错误。

ZIP 写出按路径排序并固定时间戳，对相同字节生成确定结果。它保存的是文档包，不是 DOCX、PPTX 或 XLSX。

## 保真验证范围

[保真样例](../examples/fidelity/)覆盖中文字体声明、段落、行内格式、列表、合并表格、图片和页面设置。当前验证规范 HTML/CSS 的保存重开一致、资源原字节和摘要不变、包与 ZIP 往返不丢文件。Node 与普通 JS 浏览器打包消费覆盖公开 API。

ZIP 测试另读取由 Python 标准库 `zipfile` 独立生成的保真样例归档，并使用 Node 的 zlib 构造压缩及异常输入，避免只验证同一库的自读自写。测试不要求消费者安装 Python。

这些检查未验证字体实际可用、分页或像素级视觉一致，也未实现任何编辑操作或 Office 导入导出。源样式保留、预览、编辑和导出需要分别验收，后续按[保真目标](document-format.md#保真目标)补齐固定字体和原文档视觉对照。
