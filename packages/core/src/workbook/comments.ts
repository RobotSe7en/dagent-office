import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookCommentAuthor, WorkbookCommentMention, WorkbookMetadata } from "../contracts/workbook.js";
import type { ParsedHtml } from "../document/html.js";
import { checkLimit, invalid } from "../document/validation.js";
import { resolveWorkbookCommentMention } from "./comment-mentions.js";

/** Validate the optional native comment bindings without requiring every host
 * annotation to adopt an Office author or export format. */
export function validateWorkbookComments(workbook: WorkbookMetadata, annotations: readonly Annotation[], maxItems: number, classes: ReadonlySet<string>, html: ParsedHtml | undefined): number {
  const authors = new Map<string, WorkbookCommentAuthor>(), nativeAuthors = new Set<string>(), bindings = new Set<string>();
  let items = 0;
  for (const author of workbook.commentAuthors ?? []) {
    const nativeKey = JSON.stringify([author.format, author.nativeId]);
    if (authors.has(author.authorId) || author.nativeId !== undefined && nativeAuthors.has(nativeKey)) invalid("Workbook comment author identities must be unique.");
    if (author.format === "classic" && (author.userId !== undefined || author.providerId !== undefined)) invalid("Classic note authors do not have provider identities.");
    authors.set(author.authorId, author); if (author.nativeId !== undefined) nativeAuthors.add(nativeKey); checkLimit(++items, maxItems);
  }
  const annotationMap = new Map(annotations.map(annotation => [annotation.annotationId, annotation]));
  const author = (authorId: string | undefined, format: "classic" | "modern", display: Annotation["author"]): void => {
    if (authorId === undefined) return;
    const source = authors.get(authorId);
    if (!source || source.format !== format || source.name !== display?.name) invalid("A workbook comment author binding disagrees with its displayed author.");
  };
  const nativeComments = new Set<string>();
  const revisionUid = (uid: string): void => {
    checkLimit(++items, maxItems);
    if (!/^\{[\da-f]{8}-(?:[\da-f]{4}-){3}[\da-f]{12}\}$/i.test(uid)) invalid("Native comment revision identities must be GUIDs.");
  };
  const mentionIds = new Set<string>();
  const mentions = (values: readonly WorkbookCommentMention[] | undefined, body: { instruction?: string | undefined; bodyNodeId?: string | undefined }, recoverable: boolean): void => {
    if (!values?.length) return;
    const text = body.instruction ?? (body.bodyNodeId ? html?.getNode(body.bodyNodeId)?.text : undefined);
    if (text === undefined) invalid("Comment mentions require their exact source body.");
    for (const mention of values) {
      checkLimit(++items, maxItems);
      const id = mention.mentionId.toUpperCase();
      if (mentionIds.has(id) || authors.get(mention.authorId)?.format !== "modern") invalid("Comment mentions require unique identities and an existing modern person.");
      mentionIds.add(id);
      if (!resolveWorkbookCommentMention(text, mention) && !recoverable) invalid("Invalid comment mention ranges require explicit thread recovery.");
    }
  };
  const nativeIdentity = (id: string | undefined): void => {
    if (id === undefined) return;
    const key = id.toUpperCase();
    if (nativeComments.has(key)) invalid("Native threaded comment and reply identities must be unique.");
    nativeComments.add(key);
  };
  for (const comment of workbook.comments ?? []) {
    const annotation = annotationMap.get(comment.annotationId);
    if (!annotation || bindings.has(comment.annotationId)) invalid("A workbook native comment must bind a unique existing annotation.");
    bindings.add(comment.annotationId); author(comment.authorId, comment.format, annotation.author); items++;
    if (comment.recoveryStatus !== undefined && annotation.status !== "needs_reanchor") invalid("Only an invalidated comment may retain a recovery status.");
    if (comment.recoveryReasons && (annotation.status !== "needs_reanchor" || comment.recoveryStatus === undefined || new Set(comment.recoveryReasons).size !== comment.recoveryReasons.length)) invalid("Comment recovery reasons require an invalidated annotation, its original status and unique reasons.");
    if (comment.format === "classic") {
      if (comment.nativeUid !== undefined) revisionUid(comment.nativeUid);
      if (comment.compatibilityThreadId !== undefined && (!/^\{[\da-f]{8}-(?:[\da-f]{4}-){3}[\da-f]{12}\}$/i.test(comment.compatibilityThreadId) || !comment.recoveryReasons?.includes("orphan-placeholder"))) invalid("An orphan compatibility note must retain its missing thread identity and recovery reason.");
      if (comment.recoveryReasons && (comment.recoveryReasons.length !== 1 || comment.recoveryReasons[0] !== "orphan-placeholder" || comment.compatibilityThreadId === undefined)) invalid("Classic comment recovery reasons describe an orphaned compatibility placeholder.");
      for (const box of [comment.box, comment.commentPrBox]) if (box) {
        if (box.cssClass !== undefined && (!classes.has(box.cssClass) || /\s/.test(box.cssClass))) invalid("A classic note box has an undefined appearance class.");
        items++;
      }
      if (comment.boxSource !== undefined && !comment.box) invalid("A native note box source requires its selected box.");
      if (comment.commentPrBox && (!comment.box || comment.boxSource !== "vml")) invalid("An alternate commentPr box requires the selected VML source.");
      if (comment.boxDifferences && (!comment.commentPrBox || new Set(comment.boxDifferences).size !== comment.boxDifferences.length)) invalid("Native note box differences require two sources and unique property names.");
      items += comment.boxDifferences?.length ?? 0;
      if (annotation.replies?.length) invalid("Classic notes cannot contain modern replies.");
      continue;
    }
    if (comment.recoveryReasons?.includes("orphan-placeholder")) invalid("An existing modern thread is not an orphaned compatibility placeholder.");
    if (comment.recoveryReasons?.includes("missing-location") && (annotation.anchor.kind !== "cells" || annotation.anchor.ranges.length)) invalid("A missing native location must not invent a cell range.");
    nativeIdentity(annotation.sourceId);
    for (const uid of comment.compatibilityUids ?? []) revisionUid(uid);
    const recoverable = comment.recoveryReasons?.includes("invalid-mention-range") ?? false;
    mentions(comment.mentions, annotation, recoverable);
    const replies = new Map(annotation.replies?.map(reply => [reply.replyId, reply]) ?? []), links = new Map<string, string | undefined>();
    for (const binding of comment.replies) {
      const reply = replies.get(binding.replyId);
      if (!reply || links.has(binding.replyId)) invalid("A threaded comment reply binding is missing or duplicate.");
      nativeIdentity(reply.sourceId); author(binding.authorId, "modern", reply.author); links.set(binding.replyId, binding.parentReplyId); items++;
      mentions(binding.mentions, reply, recoverable);
    }
    if (links.size !== replies.size) invalid("Every threaded comment reply requires a native binding.");
    const visited = new Set<string>();
    for (const id of links.keys()) {
      let cursor: string | undefined = id; const path = new Set<string>();
      while (cursor !== undefined && !visited.has(cursor)) {
        if (!links.has(cursor) || path.has(cursor)) invalid("Threaded reply parents must form an acyclic graph within one annotation.");
        path.add(cursor); cursor = links.get(cursor);
      }
      for (const node of path) visited.add(node);
    }
    checkLimit(items, maxItems);
  }
  checkLimit(items, maxItems); return items;
}
