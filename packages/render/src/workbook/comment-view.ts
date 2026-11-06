import { DocumentError } from "@doffi/core";
import type { WorkbookSnapshot } from "@doffi/core";
import type { FontResolution } from "../contracts.js";
import { resolvePreviewFontsInScope } from "../fonts.js";
import { markViewNode } from "../view-nodes.js";
import { createWorkbookCommentText } from "./comment-text.js";
import type { WorkbookCommentText } from "./comment-text.js";
import type { WorkbookContentSources } from "./content-source.js";
import type { WorkbookSurface } from "./surface.js";
import { checkWindowIdentity } from "./window-source.js";

export interface WorkbookCommentViewLimits {
  readonly maxNodes?: number; readonly maxSourceNodes?: number;
  readonly maxCodeUnits?: number; readonly maxReplies?: number;
}
export interface WorkbookCommentView {
  readonly element: HTMLElement; readonly nodes: number;
  readonly fonts: readonly FontResolution[];
  readonly bodies: readonly { readonly replyId?: string; readonly body: WorkbookCommentText }[];
  dispose(): void;
}
const statuses = { open: "待处理", resolved: "已处理", closed: "已关闭", needs_reanchor: "需要重新定位" };
const recoveryReasons = { "missing-location": "缺少原始单元格位置", "missing-placeholder": "缺少兼容位置记录", "mismatched-placeholder": "位置记录不一致",
  "duplicate-placeholder": "存在重复位置记录", "orphan-placeholder": "兼容记录对应的线程已缺失", "mismatched-reply-location": "回复与主评论的位置不一致",
  "duplicate-thread-location": "多个线程记录了同一位置", "invalid-mention-range": "提及的文字范围需要重新定位", "mismatched-placeholder-identity": "兼容记录的身份不一致" };
const noteBoxProperties = { anchor: "位置", appearance: "外观", visible: "可见性", printable: "打印", locked: "对象锁定", lockText: "文字锁定",
  defaultSize: "默认尺寸", disabled: "禁用状态", autoFill: "自动填充", autoLine: "自动线条", autoScale: "自动缩放", justifyLastLine: "末行对齐", altText: "替代文字" };

/** One detached, read-only thread candidate. Only its requested body/reply
 * stories enter the source surface. The owner supplies panel width, placement,
 * opening/closing controls and focus; native note-box geometry is separate. */
export function createWorkbookCommentView(surface: Pick<WorkbookSurface, "source" | "view">, snapshot: WorkbookSnapshot, sources: WorkbookContentSources,
  annotationId: string, input: WorkbookCommentViewLimits = {}): WorkbookCommentView {
  const limits = { maxNodes: 50_000, maxSourceNodes: 50_000, maxCodeUnits: 8 * 1024 * 1024, maxReplies: 1000, ...input };
  const document = surface.source.ownerDocument;
  if (!surface.source.isConnected || !surface.view.isConnected || surface.view.ownerDocument !== document
    || Object.keys(input).some(key => !["maxNodes", "maxSourceNodes", "maxCodeUnits", "maxReplies"].includes(key))
    || !Object.values(limits).every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Comment views require one connected surface and supported positive budgets.");
  checkWindowIdentity(sources.identity, snapshot);
  const annotation = snapshot.annotations.find(annotation => annotation.annotationId === annotationId);
  if (!annotation) throw new DocumentError("InvalidOperation", "The requested comment does not belong to this workbook.");
  const replies = annotation.replies ?? [];
  if (replies.length > limits.maxReplies) throw new DocumentError("ResourceLimitExceeded", "This comment exceeds its reply-count budget.");
  const ids = [annotation.bodyNodeId, ...replies.map(reply => reply.bodyNodeId)].filter((id): id is string => !!id);
  const projection = sources.project(surface.source, ids, limits.maxSourceNodes), element = markViewNode(document.createElement("section"));
  element.dataset.officeView = "workbook-comment"; element.dataset.annotationId = annotationId;
  element.setAttribute("role", "region"); element.setAttribute("aria-label", "批注与回复");
  element.style.cssText = "box-sizing:border-box;overflow-wrap:anywhere;white-space:normal;direction:ltr;text-align:start;font:11pt sans-serif;color:#20242a;background:#fff;padding:12px;border:1px solid #dce0e6";
  const native = snapshot.metadata.workbook.comments?.find(comment => comment.annotationId === annotationId);
  const people = new Map(snapshot.metadata.workbook.commentAuthors?.map(author => [author.authorId, author.name]));
  const nativeReplies = new Map(native?.format === "modern" ? native.replies.map(reply => [reply.replyId, reply]) : []);
  const bodies: { replyId?: string; body: WorkbookCommentText }[] = [];
  let nodes = 1, units = 0, disposed = false;
  const account = (count: number): void => { nodes += count; if (nodes > limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Comment views exceed their cumulative node budget."); };
  const textUnits = (text: string): void => { units += text.length; if (units > limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Comment views exceed their cumulative text budget."); };
  const label = (parent: HTMLElement, text: string, name = "p"): HTMLElement => {
    account(2); textUnits(text); const node = document.createElement(name); node.textContent = text; node.style.margin = "0 0 8px"; parent.append(node); return node;
  };
  const body = (parent: HTMLElement, sourceId: string | undefined, instruction: string | undefined, replyId?: string): void => {
    if (!sourceId) { label(parent, instruction!); return; }
    const source = projection.get(sourceId);
    if (!source) throw new DocumentError("InvalidDocument", "A comment body is missing from its owning source projection.");
    if (nodes >= limits.maxNodes || units >= limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "This comment has no remaining body budget.");
    const text = createWorkbookCommentText(source, limits.maxNodes - nodes, limits.maxCodeUnits - units);
    account(text.nodes); units += text.sourceCodeUnits; textUnits(text.element.textContent ?? "");
    parent.append(text.element); bodies.push(Object.freeze({ ...(replyId ? { replyId } : {}), body: text }));
  };
  const mentions = (parent: HTMLElement, replyId?: string): void => {
    const values = native?.format === "modern" ? replyId ? nativeReplies.get(replyId)?.mentions : native.mentions : undefined;
    for (const mention of values ?? []) {
      const name = people.get(mention.authorId);
      if (name === undefined) throw new DocumentError("InvalidDocument", "A comment mention has no preserved person.");
      const labelNode = label(parent, "提及：" + name);
      labelNode.dataset.officeView = "workbook-comment-mention"; labelNode.dataset.officeMentionId = mention.mentionId;
    }
  };
  const dispose = (): void => { if (disposed) return; disposed = true; element.remove(); element.replaceChildren(); };
  try {
    label(element, (native?.format === "classic" ? "经典批注" : native?.format === "modern" ? "现代批注" : "批注") + " · " + statuses[annotation.status], "h3");
    label(element, (annotation.author?.name ?? "未署名") + (annotation.date ? " · " + annotation.date : ""));
    if (native?.format === "classic" && native.boxDifferences?.length) {
      const warning = label(element, "原生备注框信息不一致：" + native.boxDifferences.map(name => noteBoxProperties[name]).join("、") + "。已采用 VML，另一份信息已保留。");
      warning.dataset.officeView = "workbook-comment-box-differences"; warning.style.color = "#8a4600";
    }
    if (annotation.status === "needs_reanchor") {
      const reasons = native?.recoveryReasons?.map(reason => recoveryReasons[reason]).join("；");
      const warning = label(element, "当前定位不可用" + (reasons ? "：" + reasons : "") + (native?.recoveryStatus ? "。原状态：" + statuses[native.recoveryStatus] : ""));
      warning.dataset.officeView = "workbook-comment-recovery"; warning.style.color = "#8a4600";
    }
    body(element, annotation.bodyNodeId, annotation.instruction); mentions(element);
    const parents = new Map(native?.format === "modern" ? native.replies.map(reply => [reply.replyId, reply.parentReplyId]) : []);
    const replyById = new Map(replies.map(reply => [reply.replyId, reply]));
    for (const reply of replies) {
      account(1); const section = document.createElement("section"); section.dataset.replyId = reply.replyId;
      section.style.cssText = "margin-top:12px;padding-top:8px;border-top:1px solid #dce0e6";
      const parentId = parents.get(reply.replyId); if (parentId) section.dataset.parentReplyId = parentId;
      label(section, (reply.author?.name ?? "未署名") + (parentId ? " · 回复 " + (replyById.get(parentId)?.author?.name ?? "未署名") : "")
        + (reply.status ? " · " + statuses[reply.status] : "") + (reply.date ? " · " + reply.date : ""));
      body(section, reply.bodyNodeId, reply.instruction, reply.replyId); mentions(section, reply.replyId); element.append(section);
    }
    // Resolve fonts on the isolated display surface, then detach the complete
    // candidate. Measurement never replaces an already published comment.
    const probe = markViewNode(document.createElement("div")); account(1);
    probe.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;width:0;height:0;overflow:hidden";
    probe.dataset.officeView = "workbook-comment-measurement"; probe.append(element); surface.view.append(probe);
    let fonts: readonly FontResolution[];
    try { fonts = resolvePreviewFontsInScope(document, element, "[data-office-text-style],[data-office-script]"); }
    finally { element.remove(); probe.remove(); nodes--; }
    return Object.freeze({ element, nodes, fonts, bodies: Object.freeze(bodies), dispose });
  } catch (error) { dispose(); throw error; }
  finally { projection.dispose(); }
}
