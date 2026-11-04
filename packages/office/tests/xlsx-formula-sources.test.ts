import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { unzipSync } from "fflate";
import { SaxesParser } from "saxes";
import { expect, it } from "vitest";
import { parseFormula } from "../../core/src/workbook/formula.js";
import { printFormula } from "../../core/src/workbook/formula-print.js";
import { createSharedFormulaTable } from "../../core/src/workbook/shared-formula.js";

const root = resolve(import.meta.dirname, "../../../examples/xlsx");
const expected = JSON.parse(readFileSync(resolve(root, "expected.json"), "utf8")) as { cases: { file: string }[] };
const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
function coordinate(address: string): { row: number; column: number } {
  const match = /^\$?([A-Z]+)\$?(\d+)$/.exec(address)!; let column = 0;
  for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

it("parses independent native and Calc formula sources, including shared followers and print names", () => {
  let parsed = 0, shared = 0;
  for (const directory of [root, resolve(root, "producer")]) for (const file of expected.cases) {
    const parts = unzipSync(readFileSync(resolve(directory, file.file)));
    for (const [path, bytes] of Object.entries(parts)) {
      if (!path.endsWith(".xml")) continue;
      const parser = new SaxesParser({ xmlns: true }), formulas: { value: string; address: string; id?: number; range?: string }[] = [];
      let address = "A1", active: typeof formulas[number] | undefined;
      parser.on("opentag", node => {
        const attribute = (name: string): string | undefined => Object.values(node.attributes).find(value => value.local === name)?.value;
        if (node.uri === S && node.local === "c") address = attribute("r") ?? "A1";
        if (node.uri === C && node.local === "f" || node.uri === S && ["f", "formula", "formula1", "formula2", "definedName", "calculatedColumnFormula", "totalsRowFormula"].includes(node.local)) {
          active = { value: "", address, ...(attribute("si") === undefined ? {} : { id: Number(attribute("si")) }), ...(attribute("ref") === undefined ? {} : { range: attribute("ref")! }) };
        }
      });
      parser.on("text", text => { if (active) active.value += text; });
      parser.on("closetag", node => {
        if (active && (node.uri === C && node.local === "f" || node.uri === S && ["f", "formula", "formula1", "formula2", "definedName", "calculatedColumnFormula", "totalsRowFormula"].includes(node.local))) {
          formulas.push(active); active = undefined;
        }
      });
      parser.write(new TextDecoder().decode(bytes)).close();
      const groups = createSharedFormulaTable();
      for (const formula of formulas) if (formula.id !== undefined && formula.value) {
        expect(formula.range, file.file + ": " + path).toBeDefined();
        const [start, end = start] = formula.range!.split(":");
        groups.define(formula.id, coordinate(formula.address), { from: coordinate(start!), to: coordinate(end!) }, "=" + formula.value);
      }
      for (const formula of formulas) {
        const origin = coordinate(formula.address);
        const source = formula.id === undefined ? "=" + formula.value : groups.resolve(formula.id, origin);
        const ast = parseFormula(source, origin);
        expect(parseFormula(printFormula(ast, origin), origin), file.file + ": " + path + " " + source).toEqual(ast);
        parsed++; if (formula.id !== undefined) shared++;
      }
    }
  }
  expect(parsed).toBeGreaterThan(40_000); expect(shared).toBeGreaterThanOrEqual(4);
}, 30_000);
