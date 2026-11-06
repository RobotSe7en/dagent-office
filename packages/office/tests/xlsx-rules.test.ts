import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "@dagent-office/core";
import type { WorkbookConditionalRule, WorkbookDataValidation } from "@dagent-office/core";
import { openOpc, openOpcFiles } from "../src/ooxml/package.js";
import type { OpcPackage } from "../src/ooxml/package.js";
import { readConditionalRules, readDataValidations } from "../src/xlsx/rules.js";
import { XlsxStyles } from "../src/xlsx/styles.js";

const encode = (s: string) => new TextEncoder().encode(s), decode = (s: Uint8Array) => new TextDecoder().decode(s);
const sheetPart = "xl/worksheets/sheet1.xml";
function source(producer = false) { return openOpc(readFileSync(`examples/xlsx/${producer ? "producer/" : ""}06-rules-validation.xlsx`), {}); }
function variant(edit: (source: string) => string) {
  const files = new Map(source().files); files.set(sheetPart, encode(edit(decode(files.get(sheetPart)!)))); return openOpcFiles(files, {});
}
function rules(opc: OpcPackage) {
  const root = opc.read(sheetPart), styles = new XlsxStyles(opc, "xl/workbook.xml", 1_000_000);
  return { styles, conditionalRules: readConditionalRules(root, "sheet-1", styles), ...readDataValidations(root, "sheet-1") };
}
function packageWithRules(input = rules(source())) {
  const files = new Map(createWorkbook({ documentId: "rules" })), metadata = JSON.parse(decode(files.get("metadata.json")!));
  metadata.workbook.numberFormats.push(...input.styles.numberFormats()); metadata.workbook.themeCssClass = "xlsx-theme";
  Object.assign(metadata.workbook.sheets[0], { conditionalRules: input.conditionalRules, dataValidations: input.dataValidations, disableValidationPrompts: input.disableValidationPrompts });
  files.set("metadata.json", encode(JSON.stringify(metadata))); files.set("styles.css", encode(decode(files.get("styles.css")!) + "\n" + input.styles.stylesheet()));
  return files;
}
describe("native conditional formatting and data validation source", () => {
  it("preserves all seven validation types and four supported conditional cases from independent and Calc sources", () => {
    for (const producer of [false, true]) {
      const result = rules(source(producer));
      expect(result.dataValidations?.map(rule => rule.kind)).toEqual(["list", "whole", "decimal", "date", "time", "textLength", "custom"]);
      expect(result.dataValidations?.[0]).toMatchObject({ kind: "list", expression: '="甲,乙,丙"', dropdown: true, errorStyle: "stop", allowBlank: true });
      expect(result.dataValidations?.[2]).toMatchObject({ kind: "decimal", operator: "between", formulas: ["=0", "=1"], errorStyle: "warning" });
      expect(result.conditionalRules.map(rule => rule.kind)).toEqual(["cellIs", "expression", "colorScale", "colorScale"]);
      expect(result.conditionalRules[0]).toMatchObject({ operator: "greaterThan", formulas: ["=100"], cssClass: "xlsx-dxf-0" });
      expect(result.conditionalRules[1]).toMatchObject({ expression: "=AND(B2>=50,B2<=150)", origin: { row: 1, column: 1 } });
      expect(result.conditionalRules[3]).toMatchObject({ stops: [{ kind: "min" }, { kind: "percentile", value: 50 }, { kind: "max" }] });
      const book = openDocument(packageWithRules(result), { kind: "workbook" });
      expect(openDocument(book.serialize(), { kind: "workbook" }).snapshot()).toEqual(book.snapshot());
    }
  });
  it("anchors reordered regions at their leftmost/topmost actual start and inverts native hide-dropdown semantics", () => {
    const result = rules(variant(s => s.replace('sqref="B2:B6"', 'sqref="D5:D6 B2:B3"').replace('sqref="E8"', 'sqref="H9:H10 E8"').replace('type="list"', 'type="list" showDropDown="1"')));
    expect(result.conditionalRules[1]).toMatchObject({ origin: { row: 1, column: 1 }, expression: "=AND(B2>=50,B2<=150)" });
    expect(result.dataValidations?.[6]).toMatchObject({ origin: { row: 7, column: 4 }, expression: "=E8>0" });
    expect(result.dataValidations?.[0]).toMatchObject({ dropdown: false });
    expect(openDocument(packageWithRules(result), { kind: "workbook" }).snapshot().metadata.workbook.sheets[0]?.dataValidations).toHaveLength(7);
    const crossed = rules(variant(s => s.replace('sqref="B2:B6"', 'sqref="D2 B5"').replace('sqref="E8"', 'sqref="H8 E9"')));
    expect(crossed.conditionalRules[1]).toMatchObject({ origin: { row: 4, column: 1 }, expression: "=AND(B2>=50,B2<=150)" });
    expect(crossed.dataValidations?.[6]).toMatchObject({ origin: { row: 8, column: 4 }, expression: "=E8>0" });
  });
  it("retains safe unknown functions and formula thresholds without claiming evaluation", () => {
    const result = rules(variant(s => s.replace('E8&gt;0', 'FACT(E8)&gt;0').replace('type="percentile" val="50"', 'type="formula" val="AVERAGE($D$2:$D$6)"')));
    expect(result.dataValidations?.[6]).toMatchObject({ expression: "=FACT(E8)>0" });
    expect(result.conditionalRules[3]).toMatchObject({ stops: [{ kind: "min" }, { kind: "formula", expression: "=AVERAGE($D$2:$D$6)" }, { kind: "max" }] });
    openDocument(packageWithRules(result), { kind: "workbook" });
  });
  it.each([
    (s: string) => s.replace('type="cellIs"', 'type="dataBar"'),
    (s: string) => s.replace('operator="greaterThan"', 'operator="between"'),
    (s: string) => s.replace('val="50"', 'val="101"'),
    (s: string) => s.replace('dxfId="0"', 'dxfId="999"'),
    (s: string) => s.replace('sqref="E3"', 'sqref="XFE1"'),
    (s: string) => s.replace('<dataValidations count="7">', '<dataValidations count="8">'),
    (s: string) => s.replace('type="custom"', 'type="custom" imeMode="on"'),
    (s: string) => s.replace('E8&gt;0', '[External.xlsx]Sheet1!A1&gt;0'),
  ])("rejects invalid, unsupported or unbound native rule structures", edit => { expect(() => rules(variant(edit))).toThrow(); });
  it("rejects package priority/id/style/formula-arity collisions and ambiguous overlapping validations", () => {
    const invalidCases = [
      (cf: WorkbookConditionalRule[]) => { cf[1]!.priority = cf[0]!.priority; },
      (cf: WorkbookConditionalRule[]) => { cf[1]!.ruleId = cf[0]!.ruleId; },
      (cf: WorkbookConditionalRule[]) => { if (cf[0]?.kind === "cellIs") cf[0].cssClass = "missing"; },
      (cf: WorkbookConditionalRule[]) => { if (cf[0]?.kind === "cellIs") cf[0].formulas.push("=1"); },
      (_: WorkbookConditionalRule[], dv: WorkbookDataValidation[]) => { dv[1]!.ranges = dv[0]!.ranges; },
      (cf: WorkbookConditionalRule[]) => { cf[3]!.stopIfTrue = true; },
    ];
    for (const edit of invalidCases) {
      const input = rules(source()); edit(input.conditionalRules, input.dataValidations!);
      expect(() => openDocument(packageWithRules(input))).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    }
  });
});
