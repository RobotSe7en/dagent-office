import { describe, expect, it } from "vitest";
import { parseFormula } from "../src/workbook/formula.js";
import { printFormula } from "../src/workbook/formula-print.js";

const origin = { row: 4, column: 2 };
describe("first-party native formula references", () => {
  it("distinguishes whole axes, relative offsets and absolute endpoints", () => {
    expect(parseFormula("=SUM('基础 2'!$A:C,$2:5)", origin)).toMatchObject({ kind: "call", arguments: [
      { kind: "axisRange", sheet: "基础 2", axis: "column", from: { index: 0, absolute: true }, to: { index: 0, absolute: false } },
      { kind: "axisRange", axis: "row", from: { index: 1, absolute: true }, to: { index: 0, absolute: false } },
    ] });
    expect(printFormula(parseFormula("=SUM(A:C,2:5,$A:$C,$2:$5)", origin), { row: 5, column: 3 })).toBe("=SUM(B:D,3:6,$A:$C,$2:$5)");
  });
  it("separates union from function argument commas and preserves intersections", () => {
    expect(parseFormula("=SUM((A1:A5,C1:C5),B1:B5 D1:D5)", origin)).toMatchObject({ arguments: [
      { kind: "binary", operator: "," }, { kind: "binary", operator: " " },
    ] });
    expect(parseFormula("=SUM( A1 , , 4 )", origin)).toMatchObject({ arguments: [{ kind: "cell" }, null, { value: 4 }] });
    expect(parseFormula("=SUM(#REF!:A5)", origin)).toMatchObject({ arguments: [{ operator: ":", left: { kind: "error", value: "#REF!" } }] });
  });
  it("parses selected native table specifiers without changing column text", () => {
    expect(parseFormula("=SalesTable[[#This Row],[数量]]*SalesTable[@[单价]]", origin)).toMatchObject({ left: {
      kind: "structured", table: "SalesTable", selectors: ["thisRow"], columns: ["数量"],
    }, right: { selectors: ["thisRow"], columns: ["单价"] } });
    expect(parseFormula("=SUM(T[[#Headers],[#Data],[Start]:[End]])", origin)).toMatchObject({ arguments: [{ selectors: ["headers", "data"], columns: ["Start", "End"] }] });
    expect(parseFormula("=T[Cost']'#'@'[]", origin)).toMatchObject({ columns: ["Cost]#@["] });
    expect(parseFormula("=[@Amount]", origin)).toMatchObject({ selectors: ["thisRow"], columns: ["Amount"] });
    expect(parseFormula("=T[]", origin)).toMatchObject({ selectors: [], columns: [] });
  });
  it.each([
    '=SUM(A1:$B$3,\'Q1\'\'预算\'!A:A,2:$5)', '=IF(TRUE,"a""b",#N/A)', '=SUM((A1:A5,C1:C5),A1:B5 B1:C5)',
    '=SalesTable[[#This Row],[数量]]*SalesTable[@[单价]]', '=T[Cost\']\'#\'@\'[]', '=SUM(T[[#Headers],[#Data],[Start]:[End]])',
    '=-2^2^3+50%', '=-(50%)', '=-(-2)%', '=SUM(#REF!:A5)', '=SUM(MyRange:OtherRange)',
  ])("round-trips the meaning of %s", source => {
    const ast = parseFormula(source, origin);
    expect(parseFormula(printFormula(ast, origin), origin)).toEqual(ast);
  });
  it("does not double the parser depth when printing a valid unary chain", () => {
    const ast = parseFormula("=" + "-".repeat(100) + "A1", origin);
    expect(parseFormula(printFormula(ast, origin), origin)).toEqual(ast);
  });
  it.each(['=$A', '=$2', '=1e0:3', '=1:3.0', '=0:2', '=$XFE:$XFE', '=A:2', '=1,2', '=A1 2', '=T[[Cost],[Price]]', '=T[[#All],[#Data]]', '=T[[#Totals],[#Headers]]', '=T[Unknown\'x]', '=T[[[Cost]]]', '=T[Cost', '=T[#Unknown]', '=T[[#Data],[Cost]:[#Totals]]'])('rejects malformed reference syntax: %s', source => {
    expect(() => parseFormula(source, origin)).toThrow();
  });
});
