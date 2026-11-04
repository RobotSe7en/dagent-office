import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, DEFAULT_DOCUMENT_READ_LIMITS, openDocument } from "@dagent-office/core";
import type { WorkbookTable, WorkbookMetadata } from "@dagent-office/core";
import { openOpc, openOpcFiles } from "../src/ooxml/package.js";
import type { OpcPackage } from "../src/ooxml/package.js";
import { S, xml } from "../src/ooxml/xml.js";
import { readTables, readTableStyles } from "../src/xlsx/tables.js";
import { readAutoFilter } from "../src/xlsx/filter.js";
import { XlsxStyles } from "../src/xlsx/styles.js";
import { WorkbookChunks } from "../src/xlsx/chunks.js";
import { textHtml } from "../src/xlsx/values.js";
import { writeOpcZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import { projectXlsxRecords } from "./helpers/xlsx-record-model.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const source = (producer = false) => openOpc(readFileSync(`examples/xlsx/${producer ? "producer/" : ""}05-native-table.xlsx`), {});
function variant(edits: Record<string, (s: string) => string>) {
  const files = new Map(source().files); for (const [path, edit] of Object.entries(edits)) files.set(path, encoder.encode(edit(decoder.decode(files.get(path)!))));
  return openOpcFiles(files, {});
}
function parsed(opc: OpcPackage) {
  const styles = new XlsxStyles(opc, "xl/workbook.xml", 1_000_000), tableStyles = readTableStyles(styles);
  return { styles, tableStyles, tables: readTables(opc.read("xl/worksheets/sheet1.xml"), opc, styles, tableStyles) };
}
function packageWithTables(input: ReturnType<typeof parsed>) {
  const files = new Map(createWorkbook({ documentId: "tables" })), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!));
  metadata.workbook.sheets[0].tables = input.tables; metadata.workbook.tableStyles = input.tableStyles.definitions;
  metadata.workbook.defaultTableStyle = input.tableStyles.defaultStyle; metadata.workbook.numberFormats.push(...input.styles.numberFormats());
  const chunks = new WorkbookChunks(DEFAULT_DOCUMENT_READ_LIMITS); chunks.startSheet('sheet-1');
  for (const table of input.tables) if (table.header) for (const [index, column] of table.columns.entries()) chunks.append({ row: table.range.from.row, column: table.range.from.column + index }, `<td data-column="${table.range.from.column + index}" data-value-type="text">${textHtml(column.name)}</td>`);
  chunks.finish(); for (const [path, bytes] of chunks.files) files.set(path, bytes);
  const manifest = JSON.parse(decoder.decode(files.get('manifest.json')!)); manifest.chunks = chunks.descriptors; files.set('manifest.json', encoder.encode(JSON.stringify(manifest)));
  files.set("metadata.json", encoder.encode(JSON.stringify(metadata))); files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")!) + "\n" + input.styles.stylesheet())); return files;
}
function filter(sourceText: string) { const opc = source(); return readAutoFilter(xml(encoder.encode(`<autoFilter xmlns="${S}" ref="A1:C20">${sourceText}</autoFilter>`), "filter.xml", opc.limits, { nodes: 0 }))!; }
describe("native tables and auto-filters", () => {
  it("preserves source table identity, calculated columns, total labels/functions and builtin style links", () => {
    const input = parsed(source()), table = input.tables[0]!;
    expect(table).toMatchObject({ tableId: "table-1", name: "SalesTable", header: true, totals: true, style: { kind: "builtin", name: "TableStyleMedium2" }, rowStripes: true });
    expect(table.columns[0]).toMatchObject({ name: "产品", totalLabel: "合计" });
    expect(table.columns[3]).toMatchObject({ name: "金额", calculatedFormula: "=SalesTable[[#This Row],[数量]]*SalesTable[[#This Row],[单价]]", totalFunction: "sum" });
    expect(table.autoFilter?.range.to.row).toBe(4);
    const model = openDocument(packageWithTables(input), { kind: "workbook" }); expect(openDocument(model.serialize()).snapshot()).toEqual(model.snapshot());
    const absentCount = parsed(variant({ "xl/tables/table1.xml": s => s.replace('<tableColumns count="4">', '<tableColumns>') })); expect(absentCount.tables[0]?.columns).toHaveLength(4);
    // External producer drops these definitions; the reader must not invent
    // them from cached cells or conflate a producer loss with SDK preservation.
    const calc = parsed(source(true)); expect(calc.tables[0]?.columns[3]?.calculatedFormula).toBeUndefined(); expect(calc.tables[0]?.style).toBeUndefined();
    expect(calc.tables[0]?.autoFilter?.range.to.row).toBe(5); openDocument(packageWithTables(calc));
  });
  it("maps custom style bands and differential formats without flattening their table association", () => {
    const input = parsed(variant({
      "xl/styles.xml": s => s.replace(/<tableStyles[^>]+\/>/, '<tableStyles count="1" defaultTableStyle="CustomBand"><tableStyle name="CustomBand" table="1" pivot="0" count="2"><tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="firstRowStripe" size="2" dxfId="1"/></tableStyle></tableStyles>'),
      "xl/tables/table1.xml": s => s.replace('name="TableStyleMedium2"', 'name="CustomBand"').replace('name="SalesTable" displayName', 'name="InternalObject" displayName'),
    }));
    expect(input.tableStyles.definitions[0]).toMatchObject({ name: "CustomBand", pivotCompatible: false, elements: [{ region: "wholeTable", size: 1 }, { region: "firstRowStripe", size: 2 }] });
    expect(input.tables[0]).toMatchObject({ objectName: "InternalObject", name: "SalesTable", style: { kind: "custom", styleId: "table-style-1" } });
    openDocument(packageWithTables(input));
  });
  it("reads explicit selected values/dates, comparisons, top/relative criteria and current sort definitions", () => {
    expect(filter('<filterColumn colId="0"><filters blank="1"><filter val="甲"/><dateGroupItem dateTimeGrouping="month" year="2026" month="11"/></filters></filterColumn>').columns[0]?.criterion)
      .toEqual({ kind: "values", blank: true, values: ["甲"], dates: [{ grouping: "month", year: 2026, month: 11 }] });
    expect(filter('<filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="5"/><customFilter operator="lessThan" val="100"/></customFilters></filterColumn>').columns[0])
      .toMatchObject({ column: 1, criterion: { kind: "custom", and: true, tests: [{ operator: "greaterThan", value: "5" }, { operator: "lessThan", value: "100" }] } });
    expect(filter('<filterColumn colId="2"><top10 top="0" percent="1" val="20" filterVal="100"/></filterColumn>').columns[0]?.criterion).toEqual({ kind: "top", top: false, percent: true, count: 20 });
    expect(filter('<filterColumn colId="0"><dynamicFilter type="thisMonth" val="46000" maxVal="46030"/></filterColumn>').columns[0]?.criterion).toEqual({ kind: "dynamic", type: "thisMonth" });
    expect(filter('<sortState ref="A2:C20" caseSensitive="1"><sortCondition ref="B2:B20" descending="1"/></sortState>').sort)
      .toMatchObject({ caseSensitive: true, method: "none", keys: [{ column: 1, descending: true }] });
  });
  it("rejects invalid table parts, formulas and native criteria instead of importing a partial table", () => {
    for (const edit of [
      (s: string) => s.replace('count="4"', 'count="3"'),
      (s: string) => s.replace('<table ', '<table connectionId="1" '),
      (s: string) => s.replace('<calculatedColumnFormula>', '<calculatedColumnFormula array="1">'),
      (s: string) => s.replace('name="TableStyleMedium2"', 'name="MissingStyle"'),
      (s: string) => s.replace('ref="A1:D6"', 'ref="A1048576:D1048576"'),
    ]) expect(() => parsed(variant({ "xl/tables/table1.xml": edit }))).toThrow();
    for (const invalid of [
      '<filterColumn colId="3"/>', '<filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn>',
      '<filterColumn colId="0"><top10 percent="1" val="101"/></filterColumn>',
      '<filterColumn colId="0"><filters/><customFilters/></filterColumn>',
      '<sortState ref="A2:C20"><sortCondition ref="B2:C20"/></sortState>',
    ]) expect(() => filter(invalid)).toThrow();
  });
  it("validates global names, table/merge/filter overlap, column identities, style links and filter coordinates", () => {
    const changes: ((table: WorkbookTable, workbook: WorkbookMetadata) => void)[] = [
      table => { table.columns[1]!.columnId = table.columns[0]!.columnId; },
      table => { table.columns[1]!.name = table.columns[0]!.name; },
      table => { table.range.to.column = 4; },
      table => { table.name = "SalesTable "; },
      table => { table.columns[0]!.name = "Different header"; },
      table => { table.columns[0]!.totalFunction = "custom"; },
      table => { table.style = { kind: "custom", styleId: "missing" }; },
      table => { table.autoFilter!.columns = [{ column: 99, hiddenButton: false, showButton: true }]; },
      (table, workbook) => { workbook.names = [{ nameId: "conflict", name: table.name, expression: "=1", origin: { row: 0, column: 0 } }]; },
      (_table, workbook) => { workbook.sheets[0]!.merges = [{ from: { row: 0, column: 0 }, to: { row: 0, column: 1 } }]; },
      (table, workbook) => { workbook.sheets[0]!.autoFilter = { range: table.range, columns: [] }; },
    ];
    for (const change of changes) {
      const files = packageWithTables(parsed(source())), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!)) as { workbook: WorkbookMetadata };
      change(metadata.workbook.sheets[0]!.tables![0]!, metadata.workbook); files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
      expect(() => openDocument(files)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    }
  });
  it("rejects native column names that disagree with the actual table header cells", async () => {
    const opc = variant({ 'xl/tables/table1.xml': xml => xml.replace('name="产品"', 'name="另一个名称"') });
    const records = await readXlsxRecords(writeOpcZip(opc.files));
    expect(() => projectXlsxRecords(records)).toThrowError(expect.objectContaining({ code: 'InvalidDocument', message: 'A table column name must match its constant text header cell.' }));
  });
});
