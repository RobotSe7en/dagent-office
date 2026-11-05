import { createWorkbook, openDocument } from "@dagent-office/core";
import type { XlsxRecords } from "../../src/xlsx/read.js";
import { WorkbookContent } from "../../src/xlsx/content.js";
import { WorkbookDrawings } from "../../src/xlsx/drawing.js";
import { mapWorkbookComments } from "../../src/xlsx/comment-mapping.js";
import { auditXlsxPartCoverage } from "../../src/xlsx/package.js";

/** Test projection of the completed record adapters. Full native-part coverage,
 * appearance/format rendering and full read acceptance still gate
 * the public importer; this helper does not expose a partial SDK import API. */
export function projectXlsxRecords(records: XlsxRecords) {
  const encoder = new TextEncoder(), decoder = new TextDecoder(), documentId = "native-records";
  const files = new Map([...createWorkbook({ documentId }), ...records.chunks.files]);
  const manifest = JSON.parse(decoder.decode(files.get("manifest.json")!)); manifest.chunks = records.chunks.descriptors;
  const content = new WorkbookContent(records.limits.source), comments = mapWorkbookComments(records.threads, records.threads.threads, records.notes, records.noteBoxes, content, documentId);
  const drawings = new WorkbookDrawings(records.opc, records.styles, content);
  for (const sheet of records.sheets) drawings.sheet(sheet.source.sheetId, sheet.root);
  auditXlsxPartCoverage(records, drawings);
  manifest.assets = drawings.assets; for (const [path, bytes] of drawings.files) files.set(path, bytes);
  if (content.html()) { manifest.content = "content.html"; files.set("content.html", encoder.encode(content.html()!)); }
  files.set("annotations.json", encoder.encode(JSON.stringify(comments.annotations)));
  const metadata = { ...records.properties, workbook: { dateSystem: records.native.dateSystem, calculationMode: records.native.calculationMode, activeSheetId: records.native.activeSheetId,
    sheets: records.sheets.map(sheet => sheet.metadata), names: records.native.names, themeCssClass: "xlsx-theme", numberFormats: records.styles.numberFormats(),
    columnWidthFontCssClass: records.styles.columnWidthFontCssClass(),
    cellStyles: records.styles.cellStyles(),
    stories: content.stories, comments: comments.comments, commentAuthors: comments.authors,
    objects: drawings.objects, objectHyperlinks: drawings.hyperlinks, charts: drawings.charts.definitions,
    tableStyles: records.tableStyles.definitions, ...(records.tableStyles.defaultStyle ? { defaultTableStyle: records.tableStyles.defaultStyle } : {}) } };
  files.set("manifest.json", encoder.encode(JSON.stringify(manifest))); files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  files.set("styles.css", encoder.encode(records.styles.stylesheet()));
  return { model: openDocument(files, { kind: "workbook", limits: records.limits.source, workbookLimits: records.limits.workbook }), drawings };
}
