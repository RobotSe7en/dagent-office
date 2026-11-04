import { DocumentError } from "@dagent-office/core";
import type { WorkbookCoordinate } from "@dagent-office/core";
import type { OpcPackage } from "../ooxml/package.js";
import { check, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { address, attributes, boolean, one, spreadsheetText, value } from "./values.js";
import { spreadsheetStringText } from "./strings.js";
import type { NativeSheet } from "./workbook.js";

export const THREADED_COMMENTS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
export interface NativeCommentPerson {
  readonly id: string;
  readonly name: string;
  readonly userId?: string;
  readonly providerId?: string;
}
export interface NativeThreadedComment {
  readonly id: string;
  readonly personId: string;
  readonly parentId?: string;
  readonly cell?: WorkbookCoordinate;
  readonly date?: string;
  readonly resolved: boolean;
  readonly text: string;
  readonly source: XmlNode;
}
export interface NativeCommentThread {
  readonly sheetId: string;
  readonly root: NativeThreadedComment;
  /** Original XML order; parent identities retain nested reply relationships. */
  readonly replies: readonly NativeThreadedComment[];
}
export interface NativeThreads {
  readonly people: readonly NativeCommentPerson[];
  readonly threads: readonly NativeCommentThread[];
}
export function commentGuid(node: XmlNode, name: string): string {
  const id = value(node, name);
  if (!id || !/^\{[\da-f]{8}-(?:[\da-f]{4}-){3}[\da-f]{12}\}$/i.test(id)) fail(node, "A spreadsheet comment GUID is invalid.");
  return id.toUpperCase();
}
/** Preserve native timestamps, validating the calendar rather than relying on
 * Date.parse, which silently normalizes dates such as February 30. */
export function commentDate(node: XmlNode, name: string): string | undefined {
  const source = value(node, name); if (source === undefined) return undefined;
  const match = /^(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)(?:\.(\d+))?(Z|([+-])(\d\d):(\d\d))?$/.exec(source);
  if (!match) fail(node, "A spreadsheet comment timestamp has an unsupported dateTime representation.");
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]), hour = Number(match[4]), minute = Number(match[5]), second = Number(match[6]);
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0), days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  const zoneHour = Number(match[10] ?? 0), zoneMinute = Number(match[11] ?? 0);
  if (year < 1 || month < 1 || month > 12 || day < 1 || day > days[month - 1]! || hour > 24 || minute > 59 || second > 59
    || hour === 24 && (minute !== 0 || second !== 0 || /[1-9]/.test(match[7] ?? ""))
    || zoneHour > 14 || zoneMinute > 59 || zoneHour === 14 && zoneMinute !== 0) fail(node, "A spreadsheet comment timestamp is not a valid calendar dateTime.");
  return source;
}
function emptyExtension(node: XmlNode | undefined): void {
  if (!node) return; attributes(node, []); check(node, [], THREADED_COMMENTS);
}
function root(node: XmlNode, name: string): void {
  if (node.uri !== THREADED_COMMENTS || node.name !== name) fail(node, "A threaded comment part has an invalid root.");
  attributes(node, []);
}

/** Parse native identities and reply topology before compatibility-placeholder
 * reconciliation. In particular, never delete threads solely because a legacy
 * comments part has not yet been read. */
export function readThreadedComments(opc: OpcPackage, workbookPart: string, sheets: readonly NativeSheet[], maxItems: number): NativeThreads {
  const people = new Map<string, NativeCommentPerson>(), identities = new Set<string>(), parts = new Set<string>(), threads: NativeCommentThread[] = [];
  let items = 0;
  const count = (): void => { if (++items > maxItems) throw new DocumentError("ResourceLimitExceeded", "Workbook comment records exceed their cumulative budget."); };
  const personPart = opc.related(workbookPart, "person");
  if (personPart) {
    if (opc.contentType(personPart) !== "application/vnd.ms-excel.person+xml") fail({ part: personPart, name: "personList" }, "A spreadsheet person part has the wrong content type.");
    const source = opc.read(personPart); root(source, "personList"); check(source, ["person", "extLst"], THREADED_COMMENTS); emptyExtension(one(source, "extLst", THREADED_COMMENTS));
    for (const node of elements(source, "person", THREADED_COMMENTS)) {
      count(); attributes(node, ["displayName", "id", "userId", "providerId"]); check(node, ["extLst"], THREADED_COMMENTS); emptyExtension(one(node, "extLst", THREADED_COMMENTS));
      const id = commentGuid(node, "id"), name = value(node, "displayName");
      if (people.has(id) || name === undefined) fail(node, "A spreadsheet person is duplicate or missing its display name.");
      const userId = value(node, "userId"), providerId = value(node, "providerId");
      people.set(id, { id, name: spreadsheetText(name, node), ...(userId === undefined ? {} : { userId: spreadsheetText(userId, node) }), ...(providerId === undefined ? {} : { providerId: spreadsheetText(providerId, node) }) });
    }
  }
  for (const sheet of sheets) {
    const part = opc.related(sheet.part, "threadedComment"); if (!part) continue;
    if (parts.has(part)) fail({ part, name: "ThreadedComments" }, "A threaded comment part is shared between worksheets."); parts.add(part);
    if (opc.contentType(part) !== "application/vnd.ms-excel.threadedcomments+xml") fail({ part, name: "ThreadedComments" }, "A threaded comment part has the wrong content type.");
    const source = opc.read(part); root(source, "ThreadedComments"); check(source, ["threadedComment", "extLst"], THREADED_COMMENTS); emptyExtension(one(source, "extLst", THREADED_COMMENTS));
    const comments = new Map<string, NativeThreadedComment>();
    for (const node of elements(source, "threadedComment", THREADED_COMMENTS)) {
      count(); attributes(node, ["ref", "dT", "personId", "id", "parentId", "done"]); check(node, ["text", "mentions", "extLst"], THREADED_COMMENTS); emptyExtension(one(node, "extLst", THREADED_COMMENTS));
      const mentions = one(node, "mentions", THREADED_COMMENTS);
      if (mentions) { attributes(mentions, []); if (mentions.content.some(item => typeof item !== "string" || item.trim())) unsupported(mentions, "Threaded comment mentions are outside the current reading boundary."); }
      const id = commentGuid(node, "id"), personId = commentGuid(node, "personId"), parentId = value(node, "parentId") === undefined ? undefined : commentGuid(node, "parentId"), ref = value(node, "ref");
      if (identities.has(id) || !people.has(personId)) fail(node, "A threaded comment identity is duplicate or its author is undefined."); identities.add(id);
      const text = one(node, "text", THREADED_COMMENTS), date = commentDate(node, "dT");
      comments.set(id, { id, personId, ...(parentId === undefined ? {} : { parentId }), ...(ref === undefined ? {} : { cell: address(ref, node) }),
        ...(date === undefined ? {} : { date }), resolved: boolean(node, "done"), text: text ? spreadsheetStringText(text) : "", source: node });
    }
    if (!comments.size) fail(source, "A threaded comment part must contain at least one comment.");
    // Path compression validates arbitrary XML order and nested parent chains
    // in linear time without recursive traversal or quadratic ancestry walks.
    const roots = new Map<string, NativeThreadedComment>();
    for (const comment of comments.values()) {
      let cursor = comment; const path: NativeThreadedComment[] = [], visiting = new Set<string>();
      while (!roots.has(cursor.id)) {
        if (visiting.has(cursor.id)) fail(cursor.source, "A threaded comment reply graph contains a cycle.");
        visiting.add(cursor.id); path.push(cursor);
        if (cursor.parentId === undefined) { roots.set(cursor.id, cursor); break; }
        const parent = comments.get(cursor.parentId); if (!parent) fail(cursor.source, "A threaded comment reply has no parent on its worksheet."); cursor = parent;
      }
      const parent = roots.get(cursor.id)!;
      for (const item of path) roots.set(item.id, parent);
    }
    const groups = new Map<string, NativeThreadedComment[]>(), cells = new Set<number>();
    for (const comment of comments.values()) {
      const parent = roots.get(comment.id)!;
      if (comment === parent) {
        if (comment.cell) {
          const key = comment.cell.row * 16384 + comment.cell.column;
          if (cells.has(key)) fail(comment.source, "Multiple top-level threaded comments target the same cell."); cells.add(key);
        }
        continue;
      }
      if (comment.cell && parent.cell && (comment.cell.row !== parent.cell.row || comment.cell.column !== parent.cell.column)) fail(comment.source, "A threaded reply targets a different cell from its thread.");
      const group = groups.get(parent.id) ?? []; group.push(comment); groups.set(parent.id, group);
    }
    for (const comment of comments.values()) if (comment.parentId === undefined) threads.push({ sheetId: sheet.sheetId, root: comment, replies: groups.get(comment.id) ?? [] });
  }
  return { people: [...people.values()], threads };
}
