import { DocumentError } from "@dagent-office/core";
import type { Annotation, AnnotationReply, WorkbookComment, WorkbookCommentAuthor, WorkbookNoteBox } from "@dagent-office/core";
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
  modern(native: NativeThreads, threads: readonly NativeCommentThread[]): void {
    const authors: WorkbookCommentAuthor[] = [];
    for (const [index, person] of native.people.entries()) {
      const author: WorkbookCommentAuthor = { authorId: `xlsx-person-${index}`, format: "modern", nativeId: person.id, name: person.name,
        ...(person.userId === undefined ? {} : { userId: person.userId }), ...(person.providerId === undefined ? {} : { providerId: person.providerId }) };
      this.account(author, true); authors.push(author);
    }
    const people = new Map(authors.map(author => [author.nativeId!, author])), { annotations, comments } = this.records;
    for (const author of authors) this.records.authors.push(author);
    for (const thread of threads) {
      const cell = thread.root.cell;
      if (!cell) fail(thread.root.source, "A threaded comment has no reconciled cell anchor.");
      const annotationId = `xlsx-thread-${thread.root.id}`, rootAuthor = people.get(thread.root.personId)!;
      const replies: AnnotationReply[] = [], bindings: Extract<WorkbookComment, { format: "modern" }>["replies"] = [];
      for (const reply of thread.replies) {
        const replyId = `xlsx-reply-${reply.id}`, person = people.get(reply.personId)!;
        const mappedReply: AnnotationReply = { replyId, author: { name: person.name }, ...(reply.date === undefined ? {} : { date: reply.date }), sourceId: reply.id, status: reply.resolved ? "resolved" : "open", bodyNodeId: this.content.comment(thread.sheetId, textHtml(reply.text)) };
        this.account(mappedReply, false); replies.push(mappedReply);
        const binding = { replyId, authorId: person.authorId, ...(reply.parentId === thread.root.id ? {} : { parentReplyId: `xlsx-reply-${reply.parentId}` }) };
        this.account(binding, true); bindings.push(binding);
      }
      const annotation: Annotation = { annotationId, documentId: this.documentId, baseRevision: 0, status: thread.root.resolved ? "resolved" : "open", author: { name: rootAuthor.name }, sourceId: thread.root.id,
        ...(thread.root.date === undefined ? {} : { date: thread.root.date }), bodyNodeId: this.content.comment(thread.sheetId, textHtml(thread.root.text)),
        anchor: { kind: "cells", sheetId: thread.sheetId, ranges: [{ from: { ...cell }, to: { ...cell } }] } };
      this.account(annotation, false); if (replies.length) annotation.replies = replies; annotations.push(annotation);
      this.account({ annotationId, format: "modern", authorId: rootAuthor.authorId }, true);
      comments.push({ annotationId, format: "modern", authorId: rootAuthor.authorId, replies: bindings });
    }
  }
  classic(native: NativeClassicNotes, boxes: ReadonlyMap<string, WorkbookNoteBox>): void {
    const people = new Map(native.authors.map(author => [author.authorId, author]));
    for (const author of native.authors) { this.account(author, true); this.records.authors.push({ ...author }); }
    for (const note of native.notes) {
      if (note.placeholderId !== undefined) fail(note.source, "A legacy thread placeholder must be reconciled before annotation mapping.");
      const author = people.get(note.authorId)!;
      const key = `${note.sheetId}:${note.cell.row}:${note.cell.column}`, annotationId = `xlsx-note-${key}`, box = boxes.get(key);
      const annotation: Annotation = { annotationId, documentId: this.documentId, baseRevision: 0, status: "open", author: { name: author.name },
        sourceId: note.guid ?? key, bodyNodeId: this.content.comment(note.sheetId, note.html), anchor: { kind: "cells", sheetId: note.sheetId, ranges: [{ from: { ...note.cell }, to: { ...note.cell } }] } };
      const comment: WorkbookComment = { annotationId, format: "classic", authorId: author.authorId, ...(box === undefined ? {} : { box }), ...(note.shapeId === undefined ? {} : { nativeShapeId: note.shapeId }) };
      this.account(annotation, false); this.account(comment, true); this.records.annotations.push(annotation); this.records.comments.push(comment);
    }
  }
}
export function mapThreadedComments(native: NativeThreads, threads: readonly NativeCommentThread[], content: WorkbookContent, documentId: string): WorkbookCommentRecords {
  const mapper = new CommentMapper(content, documentId); mapper.modern(native, threads); return mapper.records;
}
export function mapWorkbookComments(native: NativeThreads, threads: readonly NativeCommentThread[], notes: NativeClassicNotes, boxes: ReadonlyMap<string, WorkbookNoteBox>, content: WorkbookContent, documentId: string): WorkbookCommentRecords {
  const mapper = new CommentMapper(content, documentId); mapper.modern(native, threads); mapper.classic(notes, boxes); return mapper.records;
}
