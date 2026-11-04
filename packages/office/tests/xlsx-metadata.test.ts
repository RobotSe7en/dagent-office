import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, DEFAULT_DOCUMENT_READ_LIMITS, openDocument, WorkbookPrintSchema } from "@dagent-office/core";
import { openOpc, openOpcFiles } from "../src/ooxml/package.js";
import { S, xml } from "../src/ooxml/xml.js";
import { readWorkbook } from "../src/xlsx/workbook.js";
import { printBand, readPrint } from "../src/xlsx/print.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const options = { limits: { maxPackageBytes: 128 * 1024 * 1024, maxFileBytes: 32 * 1024 * 1024 } };
function document(source: string) { return xml(encoder.encode(source), "test.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 }); }
function workbookVariant(edit: (text: string) => string) {
  const opc = openOpc(readFileSync("examples/xlsx/04-formulas-names.xlsx"), options), files = new Map(opc.files);
  files.set("xl/workbook.xml", encoder.encode(edit(decoder.decode(files.get("xl/workbook.xml")!))));
  return openOpcFiles(files, options);
}

describe("native workbook names and print metadata", () => {
  it("reads all independent and pinned Calc workbook/print parts into strict package contracts", async () => {
    // Worksheet cells are tested by the separate streaming and import suites.
    const { readdirSync } = await import("node:fs");
    let filesChecked = 0;
    for (const directory of ["examples/xlsx", "examples/xlsx/producer"]) for (const filename of readdirSync(directory).filter(name => name.endsWith(".xlsx"))) {
      const opc = openOpc(readFileSync(`${directory}/${filename}`), options), source = readWorkbook(opc);
      const files = new Map(createWorkbook({ documentId: "native-metadata" })), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!)), base = metadata.workbook.sheets[0];
      const sheets = source.sheets.map(sheet => {
        const header = decoder.decode(opc.files.get(sheet.part)!).replace(/<sheetData(?:\s[^>]*)?>[\s\S]*?<\/sheetData>/, "<sheetData/>");
        const print = readPrint(document(header), sheet.printNames);
        expect(WorkbookPrintSchema.safeParse(print).success, `${directory}/${filename}`).toBe(true);
        return { ...base, sheetId: sheet.sheetId, name: sheet.name, visibility: sheet.visibility, print };
      });
      Object.assign(metadata.workbook, { sheets, names: source.names, activeSheetId: source.activeSheetId, dateSystem: source.dateSystem, calculationMode: source.calculationMode });
      files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
      const model = openDocument(files, { kind: "workbook" });
      expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(model.snapshot());
      filesChecked++;
    }
    expect(filesChecked).toBe(32);
  });
  it("preserves scoped names and treats equivalent reordered Calc print titles as one definition", () => {
    const source = readWorkbook(openOpc(readFileSync("examples/xlsx/04-formulas-names.xlsx"), options));
    expect(source.names).toEqual([
      { nameId: "name-1", name: "BaseRate", expression: "='基础 2'!$B$2", origin: { row: 0, column: 0 }, hidden: false },
      { nameId: "name-2", name: "LocalNumbers", expression: "=Data!$A$2:$A$5", origin: { row: 0, column: 0 }, sheetId: "sheet-1", hidden: false },
    ]);
    const opc = openOpc(readFileSync("examples/xlsx/producer/11-print-layout.xlsx"), options), book = readWorkbook(opc), print = readPrint(opc.read(book.sheets[0]!.part), book.sheets[0]!.printNames);
    expect(print.repeatRows).toEqual({ from: 0, to: 0 }); expect(print.repeatColumns).toEqual({ from: 0, to: 0 });
    expect(print.areas).toEqual([{ from: { row: 0, column: 0 }, to: { row: 44, column: 5 } }]);
    expect(print.paper?.widthPt).toBeCloseTo(595.27559); expect(print.orientation).toBe("landscape");
    expect(print.scale).toEqual({ kind: "fit", width: 1, height: 0 }); expect(print.rowBreaks).toEqual([{ before: 23, from: 0, to: 16383 }]);
    expect(print.headerFooter.oddFooter.center).toEqual([{ kind: "text", text: "第 " }, { kind: "field", field: "page" }, { kind: "text", text: " 页 / " }, { kind: "field", field: "pages" }, { kind: "text", text: " 页" }]);
  });
  it("rejects duplicate names, macro names, bad scopes, protection, conflicting titles and unknown extensions", () => {
    const edits = [
      (s: string) => s.replace('</definedNames>', '<definedName name="baserate">1</definedName></definedNames>'),
      (s: string) => s.replace('name="BaseRate"', 'name="BaseRate" function="1"'),
      (s: string) => s.replace('localSheetId="0"', 'localSheetId="99"'),
      (s: string) => s.replace('<sheets>', '<workbookProtection lockStructure="1"/><sheets>'),
      (s: string) => s.replace('</workbook>', '<extLst><ext uri="unknown"/></extLst></workbook>'),
      (s: string) => s.replace('</definedNames>', '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$2:$2</definedName></definedNames>'),
      (s: string) => s.replace('</definedNames>', '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$L$30</definedName></definedNames>'),
    ];
    for (const edit of edits) expect(() => readWorkbook(workbookVariant(edit))).toThrow();
  });
  it("parses literal ampersands and basic fields, rejects unsupported controls, and keeps unspecified paper unspecified", () => {
    const node = document(`<worksheet xmlns="${S}"/>`);
    expect(printBand("&L && &A &C&P&R&F", node)).toEqual({ left: [{ kind: "text", text: " & " }, { kind: "field", field: "sheet" }, { kind: "text", text: " " }], center: [{ kind: "field", field: "page" }], right: [{ kind: "field", field: "file" }] });
    for (const source of ["&G", "&Bbold", "&", "&P+1"]) {
      expect(() => printBand(source, node)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    }
    const print = readPrint(node, { areas: [] }); expect(print.paper).toBeUndefined(); expect(print.marginsPt).toBeUndefined(); expect(print.orientation).toBe("default");
  });
});
