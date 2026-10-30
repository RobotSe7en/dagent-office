# 正文富文本输入

R1-11 使用 Tiptap 3.32.1 的开源核心与 ProseMirror，提供连续正文和表格内输入、跨段落替换、分段、字符／段落格式、受支持内容粘贴及中文输入法。工作台通过“编辑正文”和“返回预览与批注”切换同一会话的输入及预览界面。

## 挂载与状态

```ts
import { mountDocumentTextEditor } from "@dagent-office/editor";

const input = mountDocumentTextEditor(container, session, {
  onError(error) { host.showError(error); },
});
await input.ready;
input.focus();
// session 由宿主创建与释放。
input.dispose();
```

独立入口接受 `DocumentEditingSession`；普通工作台默认提供输入入口，`editable: false` 可隐藏它。React 包装复用相同组件。正文输入位于独立的无脚本 iframe，宿主代码挂载输入适配器；源文档的脚本、事件处理器和未授权资源不会执行。

核心文档包是唯一可保存来源。Tiptap 节点、JSON、DOM、光标和输入法中间态均为派生视图，不单独保存。适配器保留已有节点身份、标签、样式、语义属性及区块间空白；新节点由核心分配身份。字段、原生批注及修订继续引用同一源节点。全篇替换保留文档根及分节容器，页设置与页眉页脚引用不因删除正文而消失。

输入法组合期间不提交修订，组合结束统一提交一次。期间若宿主修改了当前修订，待提交输入以 `RevisionConflict` 拒绝，视图恢复到当前核心源。非法结构、Unicode 边界或预算失败也恢复有效状态。卸载取消中间态、监听、输入视图及资源 URL，宿主会话继续可用。

Tiptap 没有独立的持久化撤销历史。输入、格式、批注、脚本和提案共用 `prepare/commit` 及核心历史。Ctrl/Command+Z 和 Shift+Ctrl/Command+Z 使用核心撤销重做。

## 身份保留与精确映射

新增内容操作 `reconcileFragment`：

```json
{
  "op": "reconcileFragment",
  "nodeId": "root",
  "html": "<p data-node-id=\"p\">新文</p>",
  "textEdits": [
    { "nodeId": "p", "from": 0, "to": 1, "expectedText": "原", "text": "新" }
  ]
}
```

它协调目标的子内容：只允许保留基准目标子树内已有且唯一的身份，不接受目标根自身、外部节点或猜测的新 ID。没有身份的新元素由核心分配。完整结果仍通过 HTML、CSS、资源及 Word 引用校验；模型使用该操作时也受请求完整子树范围限制。

`textEdits` 是按实际输入步骤产生的 UTF-16 映射证据，顺序范围与 `expectedText` 对照前序步骤后的文字。核心验证代理对、变体选择符、源文字及最终文字一致性。重复文字使用实际编辑位置，不能通过字符串相似性猜测插入点。缺少映射证据的已改变文字目标显式进入 `needs_reanchor`；纯格式变化且文字及身份不变时保留文字锚点。目标之外受到文字变化影响但缺少证据的祖先也明确失效。

映射证据计入会话操作预算，HTML 与证据共同受事务字节、节点、候选和历史预算约束。准备不改变当前状态，失败整批不变。保留原字体名称与中西文角色；Word 输入按字符角色补充字体作用域，保留 UTF-16 和变体序列。

## 输入范围与验证

工具栏提供字重、斜体、下划线、颜色、字号及段落对齐。Enter 分段，Shift+Enter 换行；直接包含行内内容的单元格内 Enter 换行，Tab／Shift+Tab 在已有单元格间移动。当前编辑正文域；页眉、页脚和原生批注正文继续在预览中显示。

粘贴先在不连接页面的模板中读取，通过同一核心 HTML/CSS 和资源校验后才交给输入视图。复制得到的身份被移除并由核心重新分配。单次文本或 HTML 粘贴上限为 1 MiB，另外受核心预算约束。不支持的结构、样式和资源明确失败；混合直属区块与普通行内文字的容器暂不提供富文本输入，原文仍可预览及保存。

浏览器验收使用真实 Chromium 组合输入，覆盖重复文字锚点、中文与 emoji／变体序列、跨段落、表格、格式、粘贴、冲突、全篇替换、卸载及 Playground 下载重开。固定字体编辑视图基线与原 DOCX 预览基线分开核对。自动分页及 PDF 在 R1-12，原生 DOCX 导出在 R1-13。
