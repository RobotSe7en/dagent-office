import { DocumentError, resolveWorkbookCommentMention } from "@dagent-office/core";
import type { Annotation, AnnotationReply, WorkbookComment, WorkbookCommentAuthor, WorkbookCommentMention, WorkbookNoteBox } from "@dagent-office/core";
import { fail } from "../ooxml/xml.js";
import type { NativeCommentThread, NativeThreads } from "./threaded-comments.js";
import type { NativeClassicNotes } from "./classic-comments.js";
import { WorkbookContent } from "./content.js";
import { utf8Bytes } from "./chunks.js";
import { textHtml } from "./values.js";

export interface WorkbookCommentRecords {
  readonly authors: WorkbookCommentAuthor[];
  readonly annotations: Annotation[];
  readonly comments: WorkbookComment[];
}
type RecoveryReason = NonNullable<WorkbookComment["recoveryReasons"]>[number];
/** One cumulative source budget and output collection for both native comment
 * formats. Compatibility reconciliation must finish before either is mapped. */
class CommentMapper {
  readonly records: WorkbookCommentRecords = { authors: [], annotations: [], comments: [] };
  private annotationBytes = 2;
  private metadataBytes = 2;
  constructor(private readonly content: WorkbookContent, private readonly documentId: string) {}
  private account(record: unknown, metadata: boolean): void {
    const bytes = utf8Bytes(JSON.stringify(record)) + 64;
    const total = metadata ? this.metadataBytes += bytes : this.annotationBytes += bytes;
    if (total > this.content.limits.maxFileBytes || total > this.content.limits.maxPackageBytes) throw new DocumentError("ResourceLimitExceeded", "Workbook comment metadata exceeds its source byte budget.");
  }
  modern(native: NativeThreads, threads: readonly NativeCommentThread[], recovery: ReadonlyMap<string, readonly RecoveryReason[]> = new Map(), legacyUids: ReadonlyMap<string, readonly string[]> = new Map()): void {
    const authors: WorkbookCommentAuthor[] = [];
    for (const [index, person] of native.people.entries()) {
      const author: WorkbookCommentAuthor = { authorId: `xlsx-person-${index}`, format: "modern", nativeId: person.id, name: person.name,
        ...(person.userId === undefined ? {} : { userId: person.userId }), ...(person.providerId === undefined ? {} : { providerId: person.providerId }) };
      this.account(author, true); authors.push(author);
    }
    const people = new Map(authors.map(author => [author.nativeId!, author])), { annotations, comments } = this.records;
    const locations = new Map<string, number>();
    for (const thread of threads) if (thread.root.cell) {
      const key = JSON.stringify([thread.sheetId, thread.root.cell.row, thread.root.cell.column]); locations.set(key, (locations.get(key) ?? 0) + 1);
    }
    const mentions = (message: NativeCommentThread["root"]): WorkbookCommentMention[] | undefined => {
      if (!message.mentions?.length) return undefined;
      const result: WorkbookCommentMention[] = []; this.account({ mentions: [] }, true);
      for (const mention of message.mentions) {
        const mapped = { mentionId: mention.mentionId, authorId: people.get(mention.personId)!.authorId, startIndex: mention.startIndex, length: mention.length };
        this.account(mapped, true); result.push(mapped);
      }
      return result;
    };
    for (const author of authors) this.records.authors.push(author);
    for (const thread of threads) {
      const cell = thread.root.cell;
      const reasons = [...(recovery.get(thread.root.id) ?? [])];
      if (!cell) reasons.unshift("missing-location");
      if (cell && locations.get(JSON.stringify([thread.sheetId, cell.row, cell.column]))! > 1) reasons.push("duplicate-thread-location");
      if (cell && thread.replies.some(reply => reply.cell && (reply.cell.row !== cell.row || reply.cell.column !== cell.column))) reasons.push("mismatched-reply-location");
      if ([thread.root, ...thread.replies].some(message => message.mentions?.some(mention => !resolveWorkbookCommentMention(message.text, mention)))) reasons.push("invalid-mention-range");
      const originalStatus: "resolved" | "open" = thread.root.resolved ? "resolved" : "open";
      const annotationId = `xlsx-thread-${thread.root.id}`, rootAuthor = people.get(thread.root.personId)!;
      const replies: AnnotationReply[] = [], bindings: Extract<WorkbookComment, { format: "modern" }>["replies"] = [];
      for (const reply of thread.replies) {
        const replyId = `xlsx-reply-${reply.id}`, person = people.get(reply.personId)!;
        const mappedReply: AnnotationReply = { replyId, author: { name: person.name }, ...(reply.date === undefined ? {} : { date: reply.date }), sourceId: reply.id, status: reply.resolved ? "resolved" : "open", bodyNodeId: this.content.comment(thread.sheetId, textHtml(reply.text)) };
        this.account(mappedReply, false); replies.push(mappedReply);
        const binding = { replyId, authorId: person.authorId, ...(reply.parentId === thread.root.id ? {} : { parentReplyId: `xlsx-reply-${reply.parentId}` }),
          ...(reply.cell ? { nativeCell: { ...reply.cell } } : {}), ...(reply.mentions?.length ? { mentions: mentions(reply)! } : {}) };
        this.account({ ...binding, mentions: undefined }, true); bindings.push(binding);
      }
      const annotation: Annotation = { annotationId, documentId: this.documentId, baseRevision: 0, status: reasons.length ? "needs_reanchor" : originalStatus, author: { name: rootAuthor.name }, sourceId: thread.root.id,
        ...(thread.root.date === undefined ? {} : { date: thread.root.date }), bodyNodeId: this.content.comment(thread.sheetId, textHtml(thread.root.text)),
        anchor: { kind: "cells", sheetId: thread.sheetId, ranges: cell ? [{ from: { ...cell }, to: { ...cell } }] : [] } };
      this.account(annotation, false); if (replies.length) annotation.replies = replies; annotations.push(annotation);
      const identity: Omit<Extract<WorkbookComment, { format: "modern" }>, "replies"> = { annotationId, format: "modern", authorId: rootAuthor.authorId,
        ...(legacyUids.get(thread.root.id)?.length ? { compatibilityUids: [...legacyUids.get(thread.root.id)!] } : {}),
        ...(thread.root.mentions?.length ? { mentions: mentions(thread.root)! } : {}),
        ...(reasons.length ? { recoveryStatus: originalStatus, recoveryReasons: reasons } : {}) };
      this.account({ ...identity, mentions: undefined }, true); comments.push({ ...identity, replies: bindings });
    }
  }
  classic(native: NativeClassicNotes, boxes: ReadonlyMap<string, WorkbookNoteBox>, orphans: ReadonlyMap<string, string>): void {
    const people = new Map(native.authors.map(author => [author.authorId, author]));
    for (const author of native.authors) { this.account(author, true); this.records.authors.push({ ...author }); }
    for (const note of native.notes) {
      if (note.placeholderId !== undefined) fail(note.source, "A legacy thread placeholder must be reconciled before annotation mapping.");
      const author = people.get(note.authorId)!;
      const key = `${note.sheetId}:${note.cell.row}:${note.cell.column}`, annotationId = `xlsx-note-${key}`, box = boxes.get(key);
      const compatibilityThreadId = orphans.get(key);
      const annotation: Annotation = { annotationId, documentId: this.documentId, baseRevision: 0, status: compatibilityThreadId ? "needs_reanchor" : "open", author: { name: author.name },
        sourceId: note.guid ?? key, bodyNodeId: this.content.comment(note.sheetId, note.html), anchor: { kind: "cells", sheetId: note.sheetId, ranges: [{ from: { ...note.cell }, to: { ...note.cell } }] } };
      const comment: WorkbookComment = { annotationId, format: "classic", authorId: author.authorId, ...(box === undefined ? {} : { box }), ...(note.shapeId === undefined ? {} : { nativeShapeId: note.shapeId }),
        ...(note.uid === undefined ? {} : { nativeUid: note.uid }),
        ...(compatibilityThreadId ? { compatibilityThreadId, recoveryStatus: "open", recoveryReasons: ["orphan-placeholder"] } : {}) };
      this.account(annotation, false); this.account(comment, true); this.records.annotations.push(annotation); this.records.comments.push(comment);
    }
  }
}
export function mapThreadedComments(native: NativeThreads, threads: readonly NativeCommentThread[], content: WorkbookContent, documentId: string): WorkbookCommentRecords {
  const mapper = new CommentMapper(content, documentId); mapper.modern(native, threads); return mapper.records;
}
export function mapWorkbookComments(native: NativeThreads, threads: readonly NativeCommentThread[], notes: NativeClassicNotes, boxes: ReadonlyMap<string, WorkbookNoteBox>, content: WorkbookContent, documentId: string): WorkbookCommentRecords {
  const roots = new Map(threads.map(thread => [thread.root.id, thread])), matched = new Set<string>();
  const recovery = new Map<string, RecoveryReason[]>(), orphans = new Map<string, string>(), legacyUids = new Map<string, string[]>(), classic: NativeClassicNotes["notes"][number][] = [];
  const issue = (id: string, reason: RecoveryReason): void => {
    let reasons = recovery.get(id); if (!reasons) { reasons = []; recovery.set(id, reasons); }
    if (!reasons.includes(reason)) reasons.push(reason);
  };
  for (const note of notes.notes) {
    if (note.placeholderId === undefined) { classic.push(note); continue; }
    const thread = roots.get(note.placeholderId), cell = thread?.root.cell;
    if (!thread) {
      // No authoritative modern body exists for this cache. Preserve its exact
      // classic text and provenance for recovery, without inventing a thread.
      const { placeholderId, ...orphan } = note;
      orphans.set(`${note.sheetId}:${note.cell.row}:${note.cell.column}`, placeholderId); classic.push(orphan); continue;
    }
    if (matched.has(note.placeholderId)) issue(note.placeholderId, "duplicate-placeholder");
    if (note.uid !== undefined) {
      const uids = legacyUids.get(thread.root.id) ?? []; uids.push(note.uid); legacyUids.set(thread.root.id, uids);
      if (note.uid !== thread.root.id) issue(thread.root.id, "mismatched-placeholder-identity");
    }
    if (thread.sheetId !== note.sheetId || cell && (cell.row !== note.cell.row || cell.column !== note.cell.column)) issue(note.placeholderId, "mismatched-placeholder");
    // A recognized cache never replaces the modern body or changes its cell,
    // even if it was moved/copied. Reconciliation changes only recovery state.
    matched.add(note.placeholderId);
  }
  for (const thread of threads) if (notes.parts.has(thread.sheetId) && !matched.has(thread.root.id)) {
    issue(thread.root.id, "missing-placeholder");
  }
  const mapper = new CommentMapper(content, documentId); mapper.modern(native, threads, recovery, legacyUids); mapper.classic({ ...notes, notes: classic }, boxes, orphans); return mapper.records;
}
