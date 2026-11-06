import { expect, it } from "vitest";
import { parseStructuredReference } from "../src/workbook/structured-reference.js";
import { bindStructuredReference, type FormulaTableDefinition } from "../src/workbook/formula-table-binding.js";
import { rewriteFormula } from "../src/workbook/formula-rewrite.js";

const table: FormulaTableDefinition = { name: "Sales", sheet: "Data", range: { from: { row: 0, column: 0 }, to: { row: 5, column: 3 } }, header: true, totals: true, columns: ["Product", "Quantity", "Price", "Amount"] };
const origin = { sheet: "Data", row: 2, column: 3 };
it("binds structured references to source rectangles, including empty sections and current rows", () => {
  const bind = (source: string) => bindStructuredReference(parseStructuredReference(source, "Sales"), origin, [table]);
  expect(bind("[Amount]")).toEqual({ sheet: "Data", range: { from: { row: 1, column: 3 }, to: { row: 4, column: 3 } } });
  expect(bind("[[#Headers],[#Data],[Quantity]:[Price]]")).toEqual({ sheet: "Data", range: { from: { row: 0, column: 1 }, to: { row: 4, column: 2 } } });
  expect(bind("[@Quantity]")).toEqual({ sheet: "Data", range: { from: { row: 2, column: 1 }, to: { row: 2, column: 1 } } });
  expect(bind("[#Totals]")).toEqual({ sheet: "Data", range: { from: { row: 5, column: 0 }, to: { row: 5, column: 3 } } });
  expect(bind("[Missing]")).toEqual({ error: "#REF!" });
  expect(bindStructuredReference(parseStructuredReference("[@Amount]", "Sales"), { ...origin, row: 8 }, [table])).toEqual({ error: "#VALUE!" });
  expect(bindStructuredReference(parseStructuredReference("[Amount]", "Sales"), origin, [{ ...table, totals: false, range: { ...table.range, to: { row: 0, column: 3 } } }])).toEqual({ sheet: "Data" });
});

it("preserves the referenced source row when its formula alone is cut outside the table", () => {
  const operation = { kind: "move", sheet: "Data", range: { from: { row: 2, column: 3 }, to: { row: 2, column: 3 } }, target: { row: 9, column: 8, sheet: "Other" } } as const;
  const bindings = { table: "Sales", tables: [table] };
  expect(rewriteFormula("=[@Quantity]*[@Price]", origin, operation.target, operation, bindings)).toBe("=('Data'!$B3:$B3*'Data'!$C3:$C3)");
  expect(rewriteFormula("=SUM([Amount])", origin, operation.target, operation, bindings)).toBe("=SUM(Sales[Amount])");
  expect(() => rewriteFormula("=[@Quantity]", origin, operation.target, operation)).toThrow();
});
