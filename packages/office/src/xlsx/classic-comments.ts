import { DocumentError } from "@dagent-office/core";
import type { WorkbookCommentAuthor, WorkbookCoordinate } from "@dagent-office/core";
import type { OpcPackage } from "../ooxml/package.js";
import { check, elements, fail, S } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { address, attributes, integer, one, value } from "./values.js";
import { richString, spreadsheetStringText } from "./strings.js";
import type { RichTextStyles } from "./strings.js";
import type { NativeSheet } from "./workbook.js";
import { commentGuid } from "./threaded-comments.js";
import { utf8Bytes } from "./chunks.js";

export interface NativeClassicNote {
  readonly sheetId: string;
  readonly cell: WorkbookCoordinate;
  readonly authorId: string;
  readonly html: string;
  readonly guid?: string;
  readonly shapeId?: number;
  readonly placeholderId?: string;
  /** Still requires the object-anchor adapter before complete import. */
  readonly properties?: XmlNode;
  readonly source: XmlNode;
}
export interface NativeClassicNotes {
  readonly authors: readonly WorkbookCommentAuthor[];
  readonly notes: readonly NativeClassicNote[];
  /** Preserve an empty compatibility part too: missing placeholders can have
   * different recovery implications from no legacy format being present. */
  readonly parts: ReadonlyMap<string, string>;
}
/** Rich note content and compatibility identities are read separately from
 * VML/object-anchor geometry. Every remaining property must be consumed by the
 * complete importer; this internal record stage does not discard it. */
export function readClassicComments(opc: OpcPackage, sheets: readonly NativeSheet[], styles: RichTextStyles, maxItems: number, maxBodyBytes: number): NativeClassicNotes {
  const authors: WorkbookCommentAuthor[] = [], notes: NativeClassicNote[] = [], parts = new Set<string>(), sheetParts = new Map<string, string>(), guids = new Set<string>();
  let items = 0, bytes = 0, runs = 0;
  const count = (): void => { if (++items > maxItems) throw new DocumentError("ResourceLimitExceeded", "Classic spreadsheet comments exceed their cumulative record budget."); };
  for (const sheet of sheets) {
    const part = opc.related(sheet.part, "comments"); if (!part) continue;
    if (parts.has(part)) fail({ part, name: "comments" }, "A classic comment part is shared between worksheets."); parts.add(part); sheetParts.set(sheet.sheetId, part);
    if (opc.contentType(part) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml") fail({ part, name: "comments" }, "A classic comment part has the wrong content type.");
    const source = opc.read(part);
    if (source.uri !== S || source.name !== "comments") fail(source, "A classic comment part has an invalid root.");
    attributes(source, []); check(source, ["authors", "commentList", "extLst"], S);
    const extension = one(source, "extLst"); if (extension) { attributes(extension, []); check(extension, [], S); }
    const nativeAuthors = one(source, "authors"), list = one(source, "commentList");
    if (!nativeAuthors || !list) fail(source, "A classic comment part is missing its author or comment list.");
    attributes(nativeAuthors, []); attributes(list, []); check(nativeAuthors, ["author"], S); check(list, ["comment"], S);
    const sheetAuthors = elements(nativeAuthors, "author", S).map((node, index): WorkbookCommentAuthor => {
      count(); return { authorId: `xlsx-note-author-${authors.length + index}`, format: "classic", nativeId: `${part}:${index}`, name: spreadsheetStringText(node) };
    });
    for (const author of sheetAuthors) authors.push(author);
    const cells = new Set<number>();
    for (const node of elements(list, "comment", S)) {
      count(); attributes(node, ["ref", "authorId", "guid", "shapeId"]); check(node, ["text", "commentPr"], S);
      const cell = address(value(node, "ref") ?? "", node), key = cell.row * 16384 + cell.column, author = sheetAuthors[integer(node, "authorId")];
      if (!author || cells.has(key)) fail(node, "A classic note author is missing or the cell already has a note."); cells.add(key);
      const content = one(node, "text"); if (!content) fail(node, "A classic note has no text container.");
      const html = richString(content, styles, () => { count(); return `xlsx-note-run-${++runs}`; }).html;
      bytes += utf8Bytes(html);
      if (bytes > maxBodyBytes) throw new DocumentError("ResourceLimitExceeded", "Classic spreadsheet comment bodies exceed their cumulative source byte budget.");
      const guid = value(node, "guid") === undefined ? undefined : commentGuid(node, "guid");
      if (guid !== undefined) { if (guids.has(guid)) fail(node, "Classic note GUIDs are duplicated."); guids.add(guid); }
      const matches = [...author.name.matchAll(/tc=(\{[\da-f]{8}-(?:[\da-f]{4}-){3}[\da-f]{12}\})/gi)];
      if (matches.length > 1) fail(node, "A legacy threaded comment placeholder names multiple thread identities.");
      const placeholderId = matches[0]?.[1]?.toUpperCase(), properties = one(node, "commentPr");
      notes.push({ sheetId: sheet.sheetId, cell, authorId: author.authorId, html, source: node,
        ...(guid === undefined ? {} : { guid }), ...(value(node, "shapeId") === undefined ? {} : { shapeId: integer(node, "shapeId") }),
        ...(placeholderId === undefined ? {} : { placeholderId }), ...(properties === undefined ? {} : { properties }) });
    }
  }
  return { authors, notes, parts: sheetParts };
}
