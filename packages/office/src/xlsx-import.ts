import { createWorkbook, DocumentError, openDocument } from "@doffi/core";
import type { DocumentPackage, WorkbookDocumentMetadata, WorkbookManifest } from "@doffi/core";
import { ImportControl } from "./import-control.js";
import { utf8Bytes } from "./xlsx/chunks.js";
import { mapWorkbookComments } from "./xlsx/comment-mapping.js";
import { WorkbookContent } from "./xlsx/content.js";
import { WorkbookDrawings } from "./xlsx/drawing.js";
import { xlsxLimits } from "./xlsx/options.js";
import type { XlsxReadOptions } from "./xlsx/options.js";
import { auditXlsxPartCoverage } from "./xlsx/package.js";
import { readXlsxRecords } from "./xlsx/read.js";
import { encode } from "./xlsx/values.js";

export interface XlsxImportOptions extends XlsxReadOptions { readonly documentId: string }

/** Own native records only during projection. The returned package retains
 * generated source and declared assets, not the native archive or XML trees. */
async function project(bytes: Uint8Array, manifest: WorkbookManifest, options: XlsxReadOptions): Promise<DocumentPackage> {
  const records = await readXlsxRecords(bytes, options), { control, limits } = records;
  const total = records.sheets.length + 1;
  await control.checkpoint({ phase: "content", completed: 0, total });
  const content = new WorkbookContent(limits.source);
  const comments = mapWorkbookComments(records.threads, records.threads.threads, records.notes, records.noteBoxes, content, manifest.documentId);
  const drawings = new WorkbookDrawings(records.opc, records.styles, content);
  await control.checkpoint({ phase: "content", completed: 1, total });
  for (const [index, sheet] of records.sheets.entries()) {
    drawings.sheet(sheet.source.sheetId, sheet.root);
    await control.checkpoint({ phase: "content", completed: index + 2, total, sheetId: sheet.source.sheetId });
  }
  auditXlsxPartCoverage(records, drawings);
  const put = (path: string, value: string | Uint8Array): void => {
    if (typeof value === "string" && utf8Bytes(value) > limits.source.maxFileBytes) throw new DocumentError("ResourceLimitExceeded", "Generated workbook source exceeds its file budget.");
    records.chunks.addFile(path, typeof value === "string" ? encode(value) : value);
  };
  manifest.chunks = records.chunks.descriptors; manifest.assets = drawings.assets;
  for (const [path, data] of drawings.files) put(path, data);
  const html = content.html();
  if (html !== undefined) { manifest.content = "content.html"; put(manifest.content, html); }
  const metadata: WorkbookDocumentMetadata = { ...records.properties, workbook: {
    dateSystem: records.native.dateSystem, calculationMode: records.native.calculationMode, activeSheetId: records.native.activeSheetId,
    sheets: records.sheets.map(sheet => sheet.metadata), names: records.native.names, themeCssClass: "xlsx-theme", numberFormats: records.styles.numberFormats(),
    columnWidthFontCssClass: records.styles.columnWidthFontCssClass(), cellStyles: records.styles.cellStyles(),
    stories: content.stories, comments: comments.comments, commentAuthors: comments.authors,
    objects: drawings.objects, objectHyperlinks: drawings.hyperlinks, charts: drawings.charts.definitions,
    tableStyles: records.tableStyles.definitions, ...(records.tableStyles.defaultStyle ? { defaultTableStyle: records.tableStyles.defaultStyle } : {}),
  } };
  put("manifest.json", JSON.stringify(manifest)); put("metadata.json", JSON.stringify(metadata));
  put("annotations.json", JSON.stringify(comments.annotations)); put("styles.css", records.styles.stylesheet());
  control.check();
  return records.chunks.files;
}

/** Map authorized native bytes to the sole HTML/CSS source. Cancellation or any
 * unsupported/unmapped part fails the whole import, with no partial package.
 * ZIP/chunk work yields cooperatively; final core validation is synchronous in
 * the calling thread and should run in the workbook's owning Worker for UI use. */
export async function importXlsx(bytes: Uint8Array, input: XlsxImportOptions): Promise<DocumentPackage> {
  if (!input || typeof input !== "object" || Array.isArray(input) || typeof input.documentId !== "string" || !input.documentId)
    throw new DocumentError("InvalidDocument", "XLSX import requires a document identity supplied by its host.");
  if (Object.keys(input).some(key => !["documentId", "limits", "workbookLimits", "nativeLimits", "signal", "onProgress"].includes(key)))
    throw new DocumentError("InvalidOperation", "Unsupported XLSX import options.");
  const limits = xlsxLimits(input), control = new ImportControl(input); control.check();
  const decoder = new (globalThis as unknown as { TextDecoder: new () => { decode(value: Uint8Array): string } }).TextDecoder();
  const manifest = JSON.parse(decoder.decode(createWorkbook({ documentId: input.documentId }).get("manifest.json")!)) as WorkbookManifest;
  // Freeze limit copies and capture control hooks before native ZIP yields.
  const options: XlsxReadOptions = { limits: limits.source, workbookLimits: limits.workbook, nativeLimits: limits.native, ...control.options };
  try {
    const files = await project(bytes, manifest, options);
    await control.checkpoint({ phase: "validate", completed: 0, total: 1 });
    const result = openDocument(files, { kind: "workbook", limits: limits.source, workbookLimits: limits.workbook }).serialize();
    await control.checkpoint({ phase: "validate", completed: 1, total: 1 });
    return result;
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    throw new DocumentError("InvalidDocument", "XLSX could not be mapped to the supported workbook model.");
  }
}
