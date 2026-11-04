import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createSharedWorkbookFormulas, DEFAULT_DOCUMENT_READ_LIMITS } from "@dagent-office/core";
import type { WorkbookCoordinate } from "@dagent-office/core";
import { decodeWorkbookCell } from "../../core/src/workbook/chunk-index.js";
import { openOpc } from "../src/ooxml/package.js";
import { S, xml } from "../src/ooxml/xml.js";
import { ImportControl } from "../src/import-control.js";
import { richString, readSharedStrings } from "../src/xlsx/strings.js";
import { scanWorksheet } from "../src/xlsx/scan.js";
import { cellFormula, readCell } from "../src/xlsx/cell.js";

const encoder = new TextEncoder(), limits = { ...DEFAULT_DOCUMENT_READ_LIMITS, maxNodes: 2_000_000 };
const styles = { run: () => "rich-run", cell: (index: number) => ({ cssClass: `cell-${index}`, numberFormatId: "general" }), tabStyle: () => "cell-0" };
const control = new ImportControl(), stringsOptions = () => ({ limits, budget: { nodes: 0 }, maxRetainedNodes: 1000, maxStrings: 10000, maxCodeUnits: 1_000_000, control });
const document = (source: string) => xml(encoder.encode(source), "cell.xml", limits, { nodes: 0 });
async function source(filename: string) {
  const opc = openOpc(readFileSync(filename), { limits }), part = "xl/worksheets/sheet1.xml", shared = createSharedWorkbookFormulas();
  const strings = await readSharedStrings(opc.files.get("xl/sharedStrings.xml"), "xl/sharedStrings.xml", styles, stringsOptions());
  const cells = new Map<string, ReturnType<typeof decodeWorkbookCell>>();
  await scanWorksheet(opc.files.get(part)!, part, { ...stringsOptions(), sheetId: "s", pass: 0, cell(node, coordinate) { cellFormula(node, coordinate, shared, 0); } });
  await scanWorksheet(opc.files.get(part)!, part, { ...stringsOptions(), sheetId: "s", pass: 1, cell(node, coordinate) { cells.set(`${coordinate.row},${coordinate.column}`, decodeWorkbookCell(readCell(node, coordinate, styles, strings, shared).html)); } });
  // This helper visits only sheet1; its fixtures have all shared-string uses here.
  strings.validateReferences(); return cells;
}
describe("native cell types, rich strings and formula caches", () => {
  it.each(["examples/xlsx/01-values-text.xlsx", "examples/xlsx/producer/01-values-text.xlsx"])("retains Unicode, literal strings, whitespace and scalar types from %s", async filename => {
    const cells = await source(filename);
    expect(cells.get("0,0")?.value).toEqual({ type: "text", value: "类型与中文😀禰󠄀" });
    expect(cells.get("1,1")?.value).toEqual({ type: "text", value: "001234567890123456789" });
    expect(cells.get("2,0")?.value).toEqual({ type: "number", value: 12.5 });
    expect(cells.get("2,2")?.value).toEqual({ type: "boolean", value: true });
    expect([...cells.values()].some(cell => cell.value.type === "text" && cell.value.value === "中文 Latin 😀禰󠄀\n下一行")).toBe(true);
    expect([...cells.values()].some(cell => cell.value.type === "text" && cell.value.value === "共享字符串 & 保留空格 ")).toBe(true);
  });
  it("expands shared formula references and separates missing caches from imported values", async () => {
    const cells = await source("examples/xlsx/04-formulas-names.xlsx");
    expect(cells.get("1,1")?.formula?.expression).toBe("=A2*2"); expect(cells.get("4,1")?.formula?.expression).toBe("=(A5*2)");
    const unknown = await source("examples/xlsx/14-preserved-formulas.xlsx");
    expect(unknown.get("1,1")?.formula).toEqual({ expression: "=FACT(A2)", cacheState: "imported" });
    expect(unknown.get("1,1")?.value).toEqual({ type: "number", value: 24 });
    expect(unknown.get("2,1")?.formula).toEqual({ expression: "=FACT(A2)", cacheState: "unavailable" });
    expect(unknown.get("2,1")?.value).toEqual({ type: "blank" });
  });
  it("handles a shared origin after its follower without retaining cells and rejects orphans", async () => {
    const table = createSharedWorkbookFormulas(), origin = { row: 1, column: 0 }, follower = { row: 0, column: 0 };
    const base = document(`<c xmlns="${S}"><f t="shared" si="7" ref="A1:A2">B2</f></c>`), derived = document(`<c xmlns="${S}"><f t="shared" si="7"/></c>`);
    cellFormula(derived, follower, table, 0); cellFormula(base, origin, table, 0);
    expect(cellFormula(derived, follower, table, 1)).toBe("=B1");
    expect(() => cellFormula(derived, follower, createSharedWorkbookFormulas(), 1)).toThrow();
  });
  it("bounds the complete rich string and shared pool and rejects malformed or executable structures", async () => {
    const native = (content: string) => document(`<si xmlns="${S}">${content}</si>`);
    expect(richString(native('<t xml:space="preserve">a_x000D_&amp;&lt;b&gt;_x005F_x0041_</t>'), styles)).toEqual({ html: "a&#13;&amp;&lt;b&gt;_x0041_", codeUnits: 13 });
    for (const content of ['<t>x</t><r><t>y</t></r>', '<r><rPr/></r>', '<t xml:space="unknown">x</t>', '<rPh><t>x</t></rPh>', '<t><script/></t>', '<r><t>' + "x".repeat(20000) + '</t></r><r><t>' + "x".repeat(20000) + '</t></r>']) expect(() => richString(native(content), styles)).toThrow();
    const encoded = encoder.encode(`<sst xmlns="${S}"><si><t>one</t></si><si><t>two</t></si></sst>`);
    await expect(readSharedStrings(encoded, "strings.xml", styles, { ...stringsOptions(), maxStrings: 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    await expect(readSharedStrings(encoded, "strings.xml", styles, { ...stringsOptions(), maxCodeUnits: 5 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    const empty = await readSharedStrings(undefined, "strings.xml", styles, stringsOptions());
    const coordinate: WorkbookCoordinate = { row: 0, column: 0 };
    for (const content of ['t="s"><v>0</v>', 't="b"><v>TRUE</v>', 't="n"><v>NaN</v>', 't="n"><v>1e999</v>', 't="n"><is><t>x</t></is>', 't="n"><v>1</v><v>2</v>', 't="inlineStr"><f>1</f><is><t>x</t></is>', 't="n" cm="1"><v>1</v>']) expect(() => readCell(document(`<c xmlns="${S}" ${content}</c>`), coordinate, styles, empty, createSharedWorkbookFormulas())).toThrow();
  });
});
