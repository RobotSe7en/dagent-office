import { DocumentError } from "@dagent-office/core";
import type { Annotation, AnnotationReply, WorkbookComment, WorkbookCommentAuthor } from "@dagent-office/core";
import { fail } from "../ooxml/xml.js";
import type { NativeCommentThread, NativeThreads } from "./threaded-comments.js";
import { WorkbookContent } from "./content.js";
import { utf8Bytes } from "./chunks.js";
import { textHtml } from "./values.js";

export interface WorkbookCommentRecords {
  readonly authors: WorkbookCommentAuthor[];
  readonly annotations: Annotation[];
  readonly comments: WorkbookComment[];
}
/** Called after native compatibility reconciliation. Missing anchors must be
 * resolved or explicitly represented as recovery by that stage, never guessed. */
export function mapThreadedComments(native: NativeThreads, threads: readonly NativeCommentThread[], content: WorkbookContent, documentId: string): WorkbookCommentRecords {
  let annotationBytes = 2, metadataBytes = 2;
  // Account each small record before collecting large reply arrays. In
  // particular, one shared long author name must not amplify into an
  // unbounded JSON annotation file during final serialization.
  const account = (record: unknown, metadata: boolean): void => {
    const bytes = utf8Bytes(JSON.stringify(record)) + 64;
    const total = metadata ? metadataBytes += bytes : annotationBytes += bytes;
    if (total > content.limits.maxFileBytes || total > content.limits.maxPackageBytes) throw new DocumentError("ResourceLimitExceeded", "Workbook comment metadata exceeds its source byte budget.");
  };
  const authors: WorkbookCommentAuthor[] = [];
  for (const [index, person] of native.people.entries()) {
    const author: WorkbookCommentAuthor = { authorId: `xlsx-person-${index}`, format: "modern", nativeId: person.id, name: person.name,
      ...(person.userId === undefined ? {} : { userId: person.userId }), ...(person.providerId === undefined ? {} : { providerId: person.providerId }) };
    account(author, true); authors.push(author);
  }
  const people = new Map(authors.map(author => [author.nativeId!, author])), annotations: Annotation[] = [], comments: WorkbookComment[] = [];
  for (const thread of threads) {
    const cell = thread.root.cell;
    if (!cell) fail(thread.root.source, "A threaded comment has no reconciled cell anchor.");
    const annotationId = `xlsx-thread-${thread.root.id}`, rootAuthor = people.get(thread.root.personId)!;
    const replies: AnnotationReply[] = [], bindings: Extract<WorkbookComment, { format: "modern" }>["replies"] = [];
    for (const reply of thread.replies) {
      const replyId = `xlsx-reply-${reply.id}`, person = people.get(reply.personId)!;
      const mappedReply: AnnotationReply = { replyId, author: { name: person.name }, ...(reply.date === undefined ? {} : { date: reply.date }), sourceId: reply.id, status: reply.resolved ? "resolved" : "open", bodyNodeId: content.comment(thread.sheetId, textHtml(reply.text)) };
      account(mappedReply, false); replies.push(mappedReply);
      const binding = { replyId, authorId: person.authorId, ...(reply.parentId === thread.root.id ? {} : { parentReplyId: `xlsx-reply-${reply.parentId}` }) };
      account(binding, true); bindings.push(binding);
    }
    const annotation: Annotation = { annotationId, documentId, baseRevision: 0, status: thread.root.resolved ? "resolved" : "open", author: { name: rootAuthor.name }, sourceId: thread.root.id,
      ...(thread.root.date === undefined ? {} : { date: thread.root.date }), bodyNodeId: content.comment(thread.sheetId, textHtml(thread.root.text)),
      anchor: { kind: "cells", sheetId: thread.sheetId, ranges: [{ from: { ...cell }, to: { ...cell } }] } };
    account(annotation, false); if (replies.length) annotation.replies = replies; annotations.push(annotation);
    account({ annotationId, format: "modern", authorId: rootAuthor.authorId }, true);
    comments.push({ annotationId, format: "modern", authorId: rootAuthor.authorId, replies: bindings });
  }
  return { authors, annotations, comments };
}
