import { z } from "zod";
import { IdentifierSchema } from "./contracts/primitives.js";
import { WorkbookSheetSchema } from "./contracts/workbook.js";
import type { WorkbookDocumentMetadata } from "./contracts/workbook.js";
import type { DocumentPackage, OpenDocumentOptions } from "./runtime.js";
import { openDocument } from "./document-model.js";
import { encode } from "./document/package.js";
import { DocumentError } from "./errors.js";
import { checkLimit, readLimits } from "./document/validation.js";

const optionsSchema = z.strictObject({ documentId: IdentifierSchema, title: z.string().optional(), locale: IdentifierSchema.optional(),
  sheetName: WorkbookSheetSchema.shape.name.optional(), dateSystem: z.enum(["1900", "1904"]).optional() });
export interface CreateWorkbookOptions {
  readonly documentId: string; readonly title?: string; readonly locale?: string; readonly sheetName?: string; readonly dateSystem?: "1900" | "1904";
}
/** Create an empty, validated workbook package. No full-sheet HTML is allocated. */
export function createWorkbook(input: CreateWorkbookOptions, readOptions: OpenDocumentOptions = {}): DocumentPackage {
  const parsed = optionsSchema.safeParse(input);
  if (!parsed.success) throw new DocumentError("InvalidDocument", "Invalid new workbook options.");
  const options = parsed.data;
  const texts = [options.documentId, options.title, options.locale, options.sheetName];
  checkLimit(texts.reduce((sum, text) => sum + (text?.length ?? 0) * 6, 0), readLimits(readOptions.limits).maxFileBytes);
  for (const text of texts) if (text !== undefined) for (const character of text) {
    const point = character.codePointAt(0)!;
    if (point >= 0xd800 && point <= 0xdfff || point === 0) throw new DocumentError("InvalidDocument", "Workbook text must be well-formed Unicode without null characters.");
  }
  const metadata: WorkbookDocumentMetadata = { title: options.title ?? "未命名工作簿", locale: options.locale ?? "zh-CN", workbook: {
    dateSystem: options.dateSystem ?? "1900", calculationMode: "auto", activeSheetId: "sheet-1",
    sheets: [{ sheetId: "sheet-1", name: options.sheetName ?? "Sheet1", visibility: "visible", defaultRowHeightPt: 15, defaultColumnWidth: 8.43,
      defaultCssClass: "cell-default", rowDimensions: [], columnDimensions: [], merges: [], frozen: { rows: 0, columns: 0 }, showGridLines: true, rightToLeft: false }],
    numberFormats: [{ formatId: "general", code: "General" }],
  } };
  const files: DocumentPackage = new Map([
    ["manifest.json", encode(JSON.stringify({ format: "doffi", formatVersion: 1, documentId: options.documentId, kind: "workbook", revision: 0, chunks: [], styles: ["styles.css"], metadata: "metadata.json", annotations: "annotations.json", assets: [] }))],
    ["metadata.json", encode(JSON.stringify(metadata))], ["annotations.json", encode("[]")],
    ["styles.css", encode('.cell-default{--office-font-ascii:"Liberation Sans";--office-font-east-asia:"Noto Sans CJK SC";font-family:var(--office-font-ascii),var(--office-font-east-asia);font-size:11pt;color:#000000;background-color:#ffffff}')],
  ]);
  return openDocument(files, { ...readOptions, kind: "workbook" }).serialize();
}
