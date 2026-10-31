import type { Anchor, Annotation } from "../contracts/annotations.js";
import type { PresentationOperation } from "../contracts/presentation-operations.js";
import type { DocumentState } from "../document/state.js";
import type { PresentationEditContext } from "./presentation-edit.js";
import { anchorNodes } from "./anchors.js";
import { invalid } from "../document/validation.js";
export function commentOperation(operation: PresentationOperation): boolean { return ["addPresentationComment", "addPresentationCommentReply", "removePresentationCommentReply", "setPresentationCommentBody", "updatePresentationComment"].includes(operation.op); }
export function validatePresentationComment(base: DocumentState, operation: PresentationOperation): void {
    const anchor = operation.op === "addPresentationComment" ? operation.anchor : operation.op === "updatePresentationComment" ? operation.changes.anchor : undefined;
    if (anchor && anchorNodes(anchor).some(id => !base.html.hasNode(id)))
        invalid("Comment anchors must belong to the base source.");
    if ("annotationId" in operation) {
        const annotation = base.snapshot.annotations.find(a => a.annotationId === operation.annotationId);
        if (!annotation)
            invalid("Comment targets must belong to the base snapshot.");
        if ("replyId" in operation && operation.replyId !== undefined && !annotation.replies?.some(r => r.replyId === operation.replyId))
            invalid("Reply targets must belong to the base thread.");
    }
}
function slideFor(ctx: PresentationEditContext, anchor: Anchor): string {
    const html = ctx.parse(), surfaces = new Map(ctx.p.slides.map(s => [s.rootNodeId, s.slideId])), owners = new Set<string>();
    for (const nodeId of anchorNodes(anchor)) {
        let id: string | null = nodeId;
        while (id !== null && !surfaces.has(id))
            id = html.peekNode(id)?.parentNodeId ?? null;
        if (id === null)
            invalid("Native comment anchors require slide content.");
        owners.add(surfaces.get(id)!);
    }
    if (owners.size !== 1)
        invalid("A presentation comment thread must be anchored to one slide.");
    return [...owners][0]!;
}
function body(ctx: PresentationEditContext, html: string, slideId?: string): string {
    const source = ctx.parse(), root = source.nodeIds().find(id => source.peekNode(id)?.attributes["data-office-kind"] === "presentation")!;
    const rootNodeId = ctx.edit.insertSource(root, '<div data-office-kind="comment" data-office-story="comment">' + html + '</div>').rootNodeIds[0]!;
    ctx.p.stories.push({ kind: "comment", storyId: ctx.allocate(), rootNodeId, ...(slideId ? { slideId } : {}) });
    return rootNodeId;
}
export function editPresentationComment(ctx: PresentationEditContext, operation: PresentationOperation): void {
    const { p, annotations, explicit } = ctx;
    if (operation.op === "addPresentationComment") {
        const slideId = slideFor(ctx, operation.anchor), annotationId = ctx.allocate(), bodyNodeId = body(ctx, operation.html, slideId), annotation: Annotation = { annotationId, documentId: ctx.base.snapshot.manifest.documentId, baseRevision: ctx.base.snapshot.manifest.revision, status: "open", anchor: operation.anchor, bodyNodeId, ...(operation.author ? { author: operation.author } : {}), ...(operation.date !== undefined ? { date: operation.date } : {}) };
        annotations.push(annotation);
        explicit.add(annotationId);
        p.comments.push({ annotationId, format: operation.format ?? "modern" });
        return;
    }
    if (!("annotationId" in operation))
        invalid("Comment editing requires a thread.");
    const annotation = annotations.find(a => a.annotationId === operation.annotationId) ?? invalid("The thread was removed by an earlier operation."), definition = p.comments.find(c => c.annotationId === annotation.annotationId);
    const storySlide = annotation.bodyNodeId ? p.stories.find(s => s.rootNodeId === annotation.bodyNodeId)?.slideId : annotation.status !== "needs_reanchor" ? slideFor(ctx, annotation.anchor) : undefined;
    if (operation.op === "addPresentationCommentReply") {
        const replyId = ctx.allocate(), bodyNodeId = body(ctx, operation.html, storySlide);
        (annotation.replies ??= []).push({ replyId, bodyNodeId, ...(operation.author ? { author: operation.author } : {}), ...(operation.date !== undefined ? { date: operation.date } : {}) });
        return;
    }
    if (operation.op === "removePresentationCommentReply") {
        annotation.replies = annotation.replies?.filter(r => r.replyId !== operation.replyId);
        if (!annotation.replies?.length)
            delete annotation.replies;
        if (definition?.replyAuthors)
            definition.replyAuthors = definition.replyAuthors.filter(r => r.replyId !== operation.replyId);
        return;
    }
    if (operation.op === "setPresentationCommentBody") {
        const target = operation.replyId ? annotation.replies?.find(r => r.replyId === operation.replyId) ?? invalid("The reply was removed.") : annotation;
        if (target.bodyNodeId) {
            ctx.edit.removeSource(target.bodyNodeId, true);
            ctx.edit.insertSource(target.bodyNodeId, operation.html);
        }
        else {
            const bodyNodeId = body(ctx, operation.html, storySlide);
            delete target.instruction;
            target.bodyNodeId = bodyNodeId;
        }
        return;
    }
    if (operation.op === "updatePresentationComment") {
        const { format, ...changes } = operation.changes;
        if (changes.anchor) {
            const slideId = slideFor(ctx, changes.anchor);
            explicit.add(annotation.annotationId);
            for (const target of [annotation, ...annotation.replies ?? []])
                if (target.bodyNodeId) {
                    const story = p.stories.find(s => s.rootNodeId === target.bodyNodeId);
                    if (story)
                        story.slideId = slideId;
                }
        }
        Object.assign(annotation, changes);
        if (definition) {
            if (format !== undefined && format !== definition.format) {
                definition.format = format;
                delete definition.authorId;
                delete definition.replyAuthors;
            }
            if (changes.author !== undefined)
                delete definition.authorId;
        }
        else if (format !== undefined)
            p.comments.push({ annotationId: annotation.annotationId, format });
        return;
    }
    invalid("Unsupported comment operation.");
}
