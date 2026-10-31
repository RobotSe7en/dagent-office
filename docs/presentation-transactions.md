# PPT 核心事务与锚点

R2-07 将 Presentation 接入已有 `createDocumentSession`。候选隔离、一次提交／一次撤销、单调修订、旧候选拒绝、事件与预算沿用 [核心事务](document-session.md)。持久化仍只有 HTML/CSS、元数据、批注和不可变资源，浏览器测量和矩阵不成为第二份源。

```ts
import { createDocumentSession, openDocument, readPresentationGeometry } from "@dagent-office/core";

const session = createDocumentSession(openDocument(files));
const { documentId, revision } = session.snapshot().manifest;
const source = readPresentationGeometry(session.getNode(objectId)!.attributes)!;
const candidate = session.prepare({
  protocolVersion: 1, documentId, baseRevision: revision,
  operations: [{ op: "setPresentationGeometry", nodeId: objectId,
    geometry: { ...source, x: source.x + 24 } }],
});
// 当前源未变，candidate.preview 可在同一个只读画布预览。
// candidate.affectedSlideIds 包含直接及间接受影响的稳定幻灯片身份。
session.commit(candidate);
session.undo();
```

`setPresentationGeometry` 接受完整的 x/y/width/height/rotation/flipH/flipV；坐标和尺寸单位为 pt，允许负位置，二维对象和组必须有正尺寸，连接线允许单轴零尺寸。组可同时传 `childCoordinates: { x, y, width, height }`，子坐标尺寸必须为正。占位符未声明本地几何时，显式操作创建局部覆盖；不会修改其继承来源。数值相同的已有几何保留来源属性的写法和缺省值，不制造无变化修订。

`setPresentationSize` 接受 `{ unit: "pt", width, height }`，尺寸改变涉及所有幻灯片；仅元数据变化也会出现在候选、收据和 `documentChanged.affectedSlideIds` 中。Word 拒绝这两种格式操作，Word 提案的七类操作枚举保持现有范围。

绑定连接线在同批几何操作完成后，使用原生连接点和组合变换更新源缓存端点，包含嵌套组、旋转、翻转和布局占位符。端点、对象、文字和批注共同撤销；无效连接点或不可逆组变换使整批失败。保持来源路由类型和预设调整，不承诺 Office 自动布线算法。删除完整对象时移除其对象／图表定义、已不再使用的对应数据集和超链接，连接线解除到已删除对象的绑定并保留缓存几何；必要表格、备注或定义引用被部分破坏时失败。

七类通用内容操作也适用于对应的 PPT 内容域。文字仍用精确 UTF-16 和变体边界校验；原生字体角色由预览按新文字脚本派生。`setStyle` 不能改对象或内容面的宽高及外边距，`reconcileFragment` 不能改已有对象的类型化几何。直接设置 `color`／`background-color` 时同步原生颜色及变换，设置 `font-family` 时同步四类字体角色；单独修改东亚字体可写 `--office-font-east-asia`。片段不得自带新节点身份，最终源必须通过完整格式校验。

## 授权图片资源

```ts
import { createAssetDescriptor } from "@dagent-office/core";

const descriptor = createAssetDescriptor("assets/new.png", "image/png", authorizedBytes);
const candidate = session.prepare({
  protocolVersion: 1,
  documentId: session.snapshot().manifest.documentId,
  baseRevision: session.snapshot().manifest.revision,
  operations: [{ op: "replaceAsset", nodeId: imageSourceNodeId, assetPath: descriptor.path }],
}, { assets: [{ descriptor, bytes: authorizedBytes }] });
session.commit(candidate);
```

资源描述和字节由宿主授权提供，路径、摘要和节点身份不授予读取权限。准备时复制字节，校验摘要、容器、清单和读取／候选预算；不调用网络或资源提供者。新增路径必须由最终图片源引用，不能覆盖包文件或已有不同字节。同路径同摘要与 MIME 可复用，重复输入路径拒绝。任何后续操作失败均不留下资源；撤销恢复源引用、清单和字节，重做一起恢复。此准备选项也适用于 Word，不改变原有单参数调用。

## 按页影响与批注

PPT 区域依赖稳定幻灯片源和其母版、布局、主题、对象定义、图表数据及页面尺寸。仅修改另一页或备注／批注正文不会使该页区域失效，页顺序本身也不改变归属；相关几何、字体、继承定义或尺寸改变后保留旧矩形、锚点与基准修订，进入 `needs_reanchor`。环境字体／视口变化仍由渲染器或宿主显式报告，不猜测旧矩形的新位置。

文字和对象锚点复用精确映射；无法映射时保留完整恢复数据。原生现代线程原来的 open/resolved/closed 状态保存在 `recoveryStatus`，显式重定位后清除。新增批注默认现代格式，来源格式保留。删除线程同时移除其独占富文本正文和回复内容域，撤销完整恢复。

R2-07 提供核心事务。幻灯片和对象管理界面、原生输入、定义管理、图表和线程编辑依次归 R2-08 至 R2-12；PPT 请求／提案归 R2-13。Word 交互／输入适配器和 PPT 协作入口在对应阶段前明确拒绝 PPT。可编辑 PPTX 导出、Node PPT 渲染和独立闭环仍归 R2-14 至 R2-16。

验证覆盖原子失败、无变化及历史分支、候选身份和冲突、资源完整性与预算、继承几何／组／连接线、跨页区域、原生现代 UTF-16 锚点和 ZIP 重开；浏览器验证事务更新同一画布及撤销重做，保留全部读取和 Word 回归。

提交后的单独审查复现并修复了五项回归：无原生批注登记的旧包空事务被误判为变更、对象容器区域未失效、外边距可形成额外坐标、元数据超链接的文字边界被跨越，以及 Word 工作台拒绝 PPT 时留下部分界面。失败用例在修复前均复现，修复后与原有 Word 锚点及协调测试、三个浏览器引擎一并复核。
