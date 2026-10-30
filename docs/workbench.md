# 工作台与请求／提案闭环

R1-09 在已有预览、选择和批注交互上增加多批注请求、宿主交接、固定提案导入、候选差异及整批应用。普通 JavaScript 和可选 React 包装使用同一界面。文档源、事务、锚点及历史仍由 `core` 会话管理；模型执行、身份和正式存储由宿主负责。

## 核心 API

`createDocumentSession(model, options)` 返回 `DocumentSession`，包含原有 `DocumentEditingSession` 方法及以下入口：

```ts
const request = session.buildInteractionRequest({
  annotationIds: ["a-summary", "a-heading"],
  supportedOperations: ["replaceText", "setStyle"],
});
const candidate = session.prepareProposal(request, proposal);
// candidate.preview、changes、changedAnnotationIds 供查看；准备不会改变文档。
session.commit(candidate);
// 拒绝时使用 session.discard(candidate)。
```

请求包含当前修订下开放批注的冻结副本。文字所在行内节点提升到完整区块；父子区块重叠时合并，并保持文档顺序。原生富批注包含完整批注正文。空白区域提供所属容器的完整子树。当前采用保守上下文：保留有序的全部样式、解释区块需要的完整元数据以及当前包内可用资源描述，资源字节不进入请求。

默认 `DEFAULT_INTERACTION_LIMITS` 为 4 MiB 请求 JSON、10,000 个具有稳定身份的源节点及 256 个资源描述。`limits: { maxBytes, maxNodes, maxAssets }` 可整体覆盖；各项必须为正安全整数，最终请求还受会话文档包大小预算约束。超限明确失败，不截断子树、不丢弃已选批注。生成请求不创建修订、历史或文档变更事件。

`prepareProposal` 验证版本、文档及请求关联、当前修订、批注副本、完整源区块、样式／元数据／资源和操作子集。操作目标必须位于请求完整子树内；HTML 仍通过原事务解析及安全校验。请求可 JSON 保存后传回；相同文档 ID 不提供访问授权，宿主负责绑定主体和会话。

应用有效内容变化时，提案声明已处理且映射有效的批注与内容在同一修订、同一撤销步骤中提交。未声明批注保持开放；映射失败的目标保留 `needs_reanchor`。空提案或没有净内容变化的提案不自动解决批注、不推进修订，也不清除重做历史。旧请求、旧提案和旧候选均按修订拒绝；查看旧差异不表示可以应用。

## 普通 JavaScript 挂载

```ts
import { mountDocumentWorkbench } from "@dagent-office/editor";

const view = mountDocumentWorkbench(container, session, {
  onSubmitAnnotations: async (request) => host.submitDocumentRequest(request),
  onError: (error) => host.showError(error),
});
// 宿主取得对应原请求及返回产物后：
view.loadProposal(request, proposal);
// 宿主销毁自己的界面；session 仍由宿主管理。
view.dispose();
```

界面提供开放批注复选、生成请求、下载请求 JSON、统一提交、提案 JSON 文件导入、逐处差异、候选预览、整批应用及拒绝。内容差异显示前后文字，并可展开查看结构和样式。应用后的文档包下载由宿主通过 `session.serialize()` 和 `office` ZIP 入口完成。

公开视图保留 `selection/fonts/setMode/setZoom/focusAnnotation/ready/dispose`，并提供 `request/buildRequest/submitAnnotations/loadProposal/applyProposal/rejectProposal`。`ready` 等待当前正文及候选预览。卸载释放监听、候选、预览及资源 URL，不销毁外部会话；晚到提交和文件读取结果不重新挂载界面。

工作台一次审阅一个候选；应用或拒绝当前提案后再生成新请求或载入下一份提案。界面在审阅期间禁用对应入口，公开 API 也返回明确的 `InvalidOperation`，保留当前候选。该流程只需一个候选的数量及字节预算，不会叠加未提交候选。

`onEvent` 提供 `selectionChanged`、`annotationsSubmitted` 和 `proposalReady`。提交事件表示宿主成功接受交接，不表示模型执行完成或文件已保存。宿主提交失败返回 `HostSubmissionFailed`，文档与批注不变；未配置提交回调时仍可下载请求。

## React 与 Playground

可选入口 `@dagent-office/editor/react` 导出 `DocumentWorkbench`，接受 `session`、工作台选项及 `onViewChange`。会话由宿主创建和释放；普通入口不导入 React，非 React 消费者不需要安装 React。回调变化通过最新引用传递，不重建文档视图；会话及上下文配置变化才重新挂载。

Playground 提供报告协作样例和固定提案，使用公开 React 入口。固定示例提案调整当前请求的关联身份和修订后仍经过完整校验；它只能用于包含相应节点和内容的报告。界面交接给本地演示宿主，不执行模型。测试覆盖双批注请求、提案、应用、撤销重做、文档包下载重开及实际 DOCX 导入语义。

富文本输入、自动分页和 DOCX／PDF 导出分别进入 R1-11、R1-12 和 R1-13。完整任务与验收顺序见[第一轮细则](iteration-1.md)。
