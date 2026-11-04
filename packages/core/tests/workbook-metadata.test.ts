import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument, WorkbookPrintSchema } from "../src/index.js";
import type { WorkbookMetadata, WorkbookPrint } from "../src/index.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
function model(edit: (book: WorkbookMetadata) => void, maxNodes?: number) {
  const files = new Map(createWorkbook({ documentId: "metadata" })), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!));
  edit(metadata.workbook); files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  return openDocument(files, { kind: "workbook", ...(maxNodes === undefined ? {} : { limits: { maxNodes } }) });
}
function print(): WorkbookPrint {
  return { areas: [], paper: { widthPt: 612, heightPt: 792 }, orientation: "portrait", pageOrder: "downThenOver", marginsPt: { left: 36, right: 36, top: 36, bottom: 36, header: 18, footer: 18 }, scale: { kind: "percent", percent: 100 },
    horizontalCentered: false, verticalCentered: false, gridLines: false, headings: false, blackAndWhite: false, rowBreaks: [], columnBreaks: [],
    headerFooter: { differentFirst: false, differentOddEven: false, scaleWithDocument: true, alignWithMargins: true, oddHeader: { left: [], center: [], right: [] }, oddFooter: { left: [], center: [{ kind: "field", field: "page" }], right: [] } } };
}
const range = { from: { row: 0, column: 0 }, to: { row: 2, column: 2 } };
describe("workbook semantic metadata validation", () => {
  it("preserves scoped names, dimension number formats, links and typed print fields through reopening", () => {
    const source = model(book => {
      book.names = [
        { nameId: "global", name: "Rate", origin: { row: 0, column: 0 }, expression: "=0.2" },
        { nameId: "local", name: "rate", sheetId: "sheet-1", origin: { row: 1, column: 1 }, expression: "=Sheet1!C3", hidden: true },
      ];
      const sheet = book.sheets[0]!;
      sheet.defaultNumberFormatId = "general"; sheet.columnDimensions = [{ from: 2, to: 3, unit: "characters", numberFormatId: "general" }];
      sheet.hyperlinks = [{ range, url: "https://example.com/report", location: "Sheet1!A1", tooltip: "来源链接" }]; sheet.print = print();
    });
    expect(openDocument(source.serialize(), { kind: "workbook" }).snapshot()).toEqual(source.snapshot());
  });
  it.each(["A1", "XFD1048576", "R1C1", "R", "c", "Name Space", "SUM(A1)", "_xlnm.Print_Area", "Sheet1!Value"])("rejects reserved or non-identifier name %s", name => {
    expect(() => model(book => { book.names = [{ nameId: "n", name, origin: { row: 0, column: 0 }, expression: "=1" }]; })).toThrow();
  });
  it.each([
    (book: WorkbookMetadata) => { book.sheets[0]!.defaultNumberFormatId = "missing"; },
    (book: WorkbookMetadata) => { book.sheets[0]!.rowDimensions = [{ from: 0, to: 5, numberFormatId: "missing" }]; },
    (book: WorkbookMetadata) => { book.names = [{ nameId: "n", name: "Value", sheetId: "missing", origin: { row: 0, column: 0 }, expression: "=1" }]; },
    (book: WorkbookMetadata) => { book.names = ["Value", "value"].map((name, i) => ({ nameId: `n${i}`, name, origin: { row: 0, column: 0 }, expression: "=1" })); },
    (book: WorkbookMetadata) => { book.sheets[0]!.hyperlinks = [{ range, url: "javascript:alert(1)" }]; },
    (book: WorkbookMetadata) => { book.sheets[0]!.hyperlinks = [{ range, location: "SUM(A1:A2)" }]; },
    (book: WorkbookMetadata) => { book.sheets[0]!.hyperlinks = [{ range }]; },
    (book: WorkbookMetadata) => { book.sheets[0]!.hyperlinks = [{ range, url: "https://example.com/1" }, { range, url: "https://example.com/2" }]; },
    (book: WorkbookMetadata) => { const p = print(); p.repeatRows = { from: 2, to: 1 }; book.sheets[0]!.print = p; },
    (book: WorkbookMetadata) => { const p = print(); p.marginsPt!.left = 600; book.sheets[0]!.print = p; },
    (book: WorkbookMetadata) => { const p = print(); p.rowBreaks = [{ before: 5, from: 0, to: 1 }, { before: 4, from: 0, to: 1 }]; book.sheets[0]!.print = p; },
    (book: WorkbookMetadata) => { const p = print(); p.headerFooter.oddHeader.left = [{ kind: "text", text: "Title", cssClass: "missing" }]; book.sheets[0]!.print = p; },
  ])("rejects broken references and ambiguous source metadata", edit => expect(() => model(edit)).toThrow());
  it("counts print runs against the document metadata budget and keeps its contract strict", () => {
    expect(WorkbookPrintSchema.safeParse({ ...print(), nativeXml: "opaque" }).success).toBe(false);
    expect(() => model(book => { const p = print(); p.headerFooter.oddHeader.left = Array.from({ length: 1000 }, () => ({ kind: "text", text: "x" })); book.sheets[0]!.print = p; }, 100)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
});
