# 核心事务与历史

R1-06 提供 `createDocumentSession`，从已授权的 `DocumentModel` 建立隔离的编辑会话。`openDocument` 继续只读；会话通过相同的 HTML 文档包保存，不保留另一份编辑器 JSON。

```ts
import { createDocumentSession, openDocument } from "@dagent-office/core";
import type { DocumentPackage } from "@dagent-office/core";

declare const files: DocumentPackage;
const session = createDocumentSession(openDocument(files));
const { documentId, revision } = session.snapshot().manifest;
const candidate = session.prepare({
  protocolVersion: 1,
  documentId,
  baseRevision: revision,
  operations: [
    { op: "setStyle", nodeId: "n-title", property: "color", value: "#1f4b99" },
  ],
});
// candidate.preview、changes 和 mapping 可用于候选预览；当前文档仍未变化。
const receipt = session.commit(candidate);
session.undo();
session.redo();
const saved = session.serialize();
const reopened = openDocument(saved);
session.dispose();
```

工厂返回 `DocumentSession`（扩展 `DocumentEditingSession`），实现 `snapshot/getNode/getAnnotations/serialize`、`prepare/commit/discard`、`undo/redo/history/clearHistory`、`subscribe/dispose`。R1-09 已实现请求生成与 `prepareProposal`，完整规则见[工作台与协议](workbench.md)。富文本 UI 通过[正文输入适配器](rich-text.md)复用本入口；分页和 Office 导出另行实现。

## 操作和候选

支持七类内容操作以及三类批注操作，结构与[交互协议](editing-protocol.md#文档操作)一致。全部操作先经严格结构校验，目标 ID 必须属于基准快照；然后按顺序在隔离树执行，文字范围和 `expectedText` 对照前序操作后的状态。最终重新检查 HTML/CSS、资源、节、Word 语义和批注引用。失败不会改变文档、修订或历史。

| 操作 | 当前执行规则 |
| --- | --- |
| `replaceText` | UTF-16 范围；拒绝拆开代理对或字符与变体选择符。新文字继承起点格式，范围外文字和格式保留；LF 写为换行节点。跨段落、单元格、内容域、链接或修订等原生语义边界的范围拒绝；图片、原生制表位和明确分页通过结构或资源操作修改 |
| `setStyle` | 修改目标行内声明，保留其他声明和原样式表；属性和值必须组成单条受支持声明，不能注入额外 CSS |
| `insertFragment` | 按父元素上下文解析 `before/after/start/end`；片段不能自带 `data-node-id`，新身份由核心分配，不能在同批操作里猜测引用 |
| `replaceFragment` | 替换目标的子内容，保留目标身份；浏览器若会重新移动或丢弃插入结构则拒绝 |
| `removeNode` | 删除完整子树；不能删除文档根；破坏必要节或内容域引用时整批失败 |
| `replaceAsset` | 当前只替换 `img`，新路径必须已在当前文档清单声明并提供有效字节；保留尺寸和替代文字，不隐式上传或授权资源 |
| 批注操作 | 增加、修改和删除进入同一事务及历史；检查身份、文档、修订及最终锚点原文。富文本批注正文仍通过内容域引用保存 |

Word 字段、原生批注及插入/删除修订在局部编辑后继续保留。完整删除源对象时移除对应语义定义；只删除复杂字段的部分必要标记等情况明确失败。字段结果文字被修改时标记 `dirty`，不执行字段指令或重算结果。

中西文字体通过源 CSS 变量及 `data-word-script` 保留。Word 文字替换按脚本划分新增文字，使新插入的中文继续采用 `--word-font-eastAsia`，西文采用 `--word-font-ascii/hAnsi`；变体选择符随基字保留。只修改东亚字体时可对段落或文字片段执行 `setStyle`，属性为 `--word-font-eastAsia`，值为带引号的源字体名称。常用宋体、仿宋、楷体名称变体及字体环境见[中文字体](fonts.md)。

候选包含冻结的 `preview`、节点增删改列表 `changes`、`affectedNodeIds`、`changedAnnotationIds` 和 `mapping`。节点差异描述自身结构及属性，受影响节点另包含祖先；布局和视觉差异由预览层派生。文字替换、片段插入/替换和节点删除返回按操作顺序的 UTF-16 偏移映射，覆盖受影响的已有行内节点及祖先；结构删除同时返回节点失效映射。映射的 Map 在运行时也不暴露修改方法。

候选通过私有运行时身份绑定会话和基准修订，复制公开字段不能伪造候选。候选只能消费一次；未使用的候选可 `discard`。有效提交、撤销或重做使其他候选过期并释放其内部文档引用，旧候选的冻结预览仍可查看。

## 修订、历史和批注边界

一次有效提交增加一个修订，并形成一步撤销。无净变化的事务不增加修订、不产生历史、不清空重做分支；无操作可撤销或重做时返回 `changed: false`。撤销和重做恢复对应内容及批注，但使用新的单调修订号。撤销后提交新修改清空重做分支。

历史采用前后状态快照，内部共享不可变资源字节。`serialize()` 始终返回独立 Map 和字节副本；输入、输出及其他会话不能修改当前状态。`commit` 只修改内存会话，不代表宿主正式保存。

R1-07 在同一候选内完成精确文字与结构锚点映射。有效批注提交时绑定新修订，保持原有 `open` 或 `resolved` 状态；定位失败时整条标记 `needs_reanchor`，保留映射前的完整锚点、摘录及基准修订。不猜测语义位置，不自动解决批注。撤销重做直接恢复对应快照中的批注及锚点，同时生成新的当前修订号。

## 锚点映射与重新定位

| 编辑或锚点 | 规则 |
| --- | --- |
| 范围之前的文字变化 | 按 UTF-16 长度差移动起点和终点；之后的变化不影响范围 |
| 边界及内部插入 | 起点插入时整体后移，终点插入时不扩展；范围内部插入纳入摘录 |
| 局部替换或删除 | 保留可确定的剩余范围，包含相交的替换文字，更新 `quote`；整段范围被覆盖、删除或整体替换则失效 |
| 零长度文字锚点 | 跟随其后的源文字；同位置插入时后移，删除内部的点失效；替换边界的点保持可确定位置 |
| 行内与结构操作 | 对同一规范文字流记录偏移；换行占一个 LF，图片占一个 U+FFFC。整块 `replaceFragment` 使目标及其子树的文字锚点失效，即使替换前后文字相同；目标自身的节点身份保留 |
| 跨节点 `text-set` | 逐范围映射，保持文档顺序及同一内容域；任一范围失效或映射后范围重叠时整条失效，不裁掉丢失部分 |
| `node` / `node-set` | 保留存在的稳定身份；集合中任一节点删除时整条失效，不自动缩小集合 |
| 原有 `needs_reanchor` | 保留原始恢复数据及修订，不尝试自动修复 |

读取和最终候选均检查文字原文、代理对及变体选择符边界、跨节点顺序/重叠及实际覆盖的内容域（包括祖先节点上的范围）、节点集合重复项，以及区域渲染面和命中节点的归属。`needs_reanchor` 的旧锚点仅作恢复数据，可继续引用已经删除的节点。仅经文字操作且没有净文字变化的节点保留原范围，结构替换仍按身份与覆盖范围判定；没有净文档或批注变化的事务仍不推进修订或清除重做。

`addAnnotation` 和显式设置 `updateAnnotation.changes.anchor` 使用**最终候选**的坐标及摘录，不再自动叠加同批偏移，操作前后顺序不改变此约定。其有效目标必须属于事务基准快照，不能猜测同批新分配的节点 ID。只修改说明或状态时，已有有效锚点仍参与自动映射。重新定位旧批注时显式提交新 `anchor` 和 `status: "open"` 或 `"resolved"`；仅修改状态不能重新启用旧锚点；可以在同批中分别设置锚点和有效状态。原文或节点校验失败则整批不变。

区域锚点的矩形属于指定 `layoutKey` 下的派生布局。核心不测量 DOM、不平移或缩放旧矩形。任何净 HTML 内容、样式或资源引用变化都会使未显式重定位的区域批注失效，包括空白区域及命中节点之外的变化；仅批注事务不使区域失效。原 `surfaceId/rect/nodeIds/layoutKey/pageId` 及基准修订完整保留。

渲染器或宿主发现字体、视口等环境变化导致 `layoutKey` 变化时，通过 `updateAnnotation` 显式标记 `needs_reanchor`；完成重新测量或人工选择后，使用同一事务接口提交新的区域锚点和有效状态。核心校验渲染面是当前容器、命中节点存在且位于该容器内；`layoutKey` 和派生 `pageId` 的实际测量由调用方负责。重新定位同样受修订冲突、原子性和历史规则约束。R1-08 已提供这些操作的[界面和布局检测](preview-interaction.md)。

事件为 `documentChanged`、`annotationChanged` 和 `error`。变更事件在提交完成后同步发送，回调读取到对应修订；回调不能重入修改会话，回调异常不改变提交结果。订阅返回取消函数；`dispose()` 可重复调用，释放候选、历史和订阅，后续读取或编辑返回 `InvalidOperation`。

## 预算和验证

`createDocumentSession(model, { limits, sessionLimits })` 中，`limits` 沿用[读取预算](document-model.md#输入预算)。会话预算默认值通过 `DEFAULT_DOCUMENT_SESSION_LIMITS` 导出：

| 项目 | 默认值 | 语义 |
| --- | --- | --- |
| `maxOperations` | 1024 | 单批操作数，结构解析前检查 |
| `maxTransactionBytes` | 8 MiB | 事务字符串按 JSON 转义上界计算，解析和字符串化前检查；循环引用拒绝 |
| `maxCandidates` | 8 | 当前仍保留内部状态的候选数 |
| `maxCandidateBytes` | 64 MiB | 活跃候选的序列化文件字节；共享资源只计一次 |
| `maxHistorySteps` | 100 | 撤销及重做共用的历史步数 |
| `maxHistoryBytes` | 64 MiB | 历史状态的序列化文件字节；共享资源只计一次，不等同于 JS 堆内存 |

预算须为正的安全整数，未知项拒绝。超限返回 `ResourceLimitExceeded`，失败候选不被消费，当前状态和历史不变；不自动裁剪已有历史。宿主可调整预算或显式 `clearHistory()` 后继续提交。`history()` 返回撤销步数、重做步数和保留字节。默认值是保护边界，尚不是测量后的性能承诺。

核心测试覆盖批量原子性、顺序范围、无变化、会话隔离、候选身份、旧修订、历史分支、UTF-16、片段安全、资源和超限；DOCX 测试覆盖九类源文件及其 LibreOffice 重存版本的编辑、文档包 ZIP 保存重开，以及中文字体名称和角色保留。DOCX 原生可编辑导出与参考应用重开仍在 R1-13 验收。
