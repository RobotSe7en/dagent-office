import { AnchorSchema } from "@dagent-office/core";
import type { DocumentEditingSession, DocumentOperation, DocumentSelection } from "@dagent-office/core";
export interface CommentControlsHost {
    readonly session: DocumentEditingSession;
    selection(): DocumentSelection | undefined;
    slideRoot(): string;
    perform(operations: DocumentOperation[]): void;
    edit(annotationId: string, replyId?: string): void;
    locate(annotationId: string): void;
    attempt(action: () => void): void;
}
export interface CommentControls {
    refresh(): void;
    dispose(): void;
}
const escape = (text: string): string => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br>");
export function createCommentControls(document: Document, parent: HTMLElement, host: CommentControlsHost): CommentControls {
    const pane = document.createElement("fieldset"), legend = document.createElement("legend"), list = document.createElement("div"), controls = new AbortController();
    let listControls = new AbortController();
    legend.textContent = "批注与回复";
    pane.append(legend);
    parent.append(pane);
    const field = (label: string): HTMLInputElement => { const wrap = document.createElement("label"), node = document.createElement("input"); node.type = "text"; node.setAttribute("aria-label", label); wrap.append(label, node); pane.append(wrap); return node; };
    const author = field("批注显示署名"), draft = field("新批注正文"), format = document.createElement("select");
    format.setAttribute("aria-label", "新批注格式");
    for (const [value, label] of [["modern", "现代批注"], ["classic", "经典批注"]]) {
        const option = document.createElement("option");
        option.value = value!;
        option.textContent = label!;
        format.append(option);
    }
    pane.append(format);
    const button = (label: string, action: () => void, parent: HTMLElement = pane, signal = controls.signal): void => { const node = document.createElement("button"); node.type = "button"; node.textContent = label; node.addEventListener("click", () => host.attempt(action), { signal }); parent.append(node); };
    button("添加批注", () => { const before = new Set(host.session.getAnnotations().map(a => a.annotationId)), anchor = host.selection()?.anchor ?? { kind: "node" as const, nodeId: host.slideRoot() }; host.perform([{ op: "addPresentationComment", anchor: AnchorSchema.parse(anchor), html: "<p>" + escape(draft.value) + "</p>", format: format.value as "classic" | "modern", ...(author.value ? { author: { name: author.value } } : {}) }]); draft.value = ""; host.edit(host.session.getAnnotations().find(a => !before.has(a.annotationId))!.annotationId); });
    pane.append(list);
    return { refresh() {
            listControls.abort();
            listControls = new AbortController();
            list.replaceChildren();
            const p = host.session.snapshot().metadata.presentation!;
            for (const annotation of host.session.getAnnotations()) {
                const item = document.createElement("section"), title = document.createElement("p"), body = annotation.bodyNodeId ? host.session.getNode(annotation.bodyNodeId)?.text : annotation.instruction;
                title.textContent = (annotation.author?.name ?? "未署名") + " · " + (p.comments.find(c => c.annotationId === annotation.annotationId)?.format === "classic" ? "经典" : "现代") + " · " + ({ open: "待处理", resolved: "已解决", closed: "已关闭", needs_reanchor: "待重新定位" }[annotation.status]) + " · " + (body ?? "");
                item.append(title);
                list.append(item);
                const action = (label: string, callback: () => void): void => button(label, callback, item, listControls.signal);
                action("编辑批注正文", () => host.edit(annotation.annotationId));
                action("定位批注", () => host.locate(annotation.annotationId));
                action("删除批注", () => host.perform([{ op: "removeAnnotation", annotationId: annotation.annotationId }]));
                for (const [label, status] of [["解决批注", "resolved"], ["重新打开批注", "open"], ["关闭批注", "closed"]] as const)
                    action(label, () => host.perform([{ op: "updatePresentationComment", annotationId: annotation.annotationId, changes: { status } }]));
                action("按当前选择重新定位", () => { const anchor = host.selection()?.anchor; if (!anchor)
                    throw new Error("请先选择新的幻灯片内容。"); const status = annotation.status === "needs_reanchor" ? p.comments.find(c => c.annotationId === annotation.annotationId)?.recoveryStatus ?? "open" : annotation.status; host.perform([{ op: "updatePresentationComment", annotationId: annotation.annotationId, changes: { anchor: AnchorSchema.parse(anchor), status } }]); });
                const reply = document.createElement("input");
                reply.setAttribute("aria-label", "回复正文");
                item.append(reply);
                action("添加回复", () => { const before = new Set(annotation.replies?.map(r => r.replyId)); host.perform([{ op: "addPresentationCommentReply", annotationId: annotation.annotationId, html: "<p>" + escape(reply.value) + "</p>", ...(author.value ? { author: { name: author.value } } : {}) }]); host.edit(annotation.annotationId, host.session.getAnnotations().find(a => a.annotationId === annotation.annotationId)!.replies!.find(r => !before.has(r.replyId))!.replyId); });
                for (const entry of annotation.replies ?? []) {
                    const section = document.createElement("div");
                    section.textContent = (entry.author?.name ?? "未署名") + "：" + (entry.bodyNodeId ? host.session.getNode(entry.bodyNodeId)?.text : entry.instruction);
                    item.append(section);
                    button("编辑回复", () => host.edit(annotation.annotationId, entry.replyId), section, listControls.signal);
                    button("删除回复", () => host.perform([{ op: "removePresentationCommentReply", annotationId: annotation.annotationId, replyId: entry.replyId }]), section, listControls.signal);
                }
            }
        }, dispose() { controls.abort(); listControls.abort(); pane.remove(); } };
}
