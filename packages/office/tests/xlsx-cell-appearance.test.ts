import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { DEFAULT_DOCUMENT_READ_LIMITS, openDocument } from "@dagent-office/core";
import { S, xml } from "../src/ooxml/xml.js";
import { alignmentDeclarations, borderDeclarations, fillDeclarations } from "../src/xlsx/appearance.js";
import { SpreadsheetColors } from "../src/xlsx/colors.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import { projectXlsxRecords } from "./helpers/xlsx-record-model.js";

const node = (source: string) => xml(new TextEncoder().encode(source), "styles.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 });
const colors = new SpreadsheetColors(undefined, new Set(["accent1"]));
const fill = (body: string) => fillDeclarations(node(`<fill xmlns="${S}">${body}</fill>`), colors);
const border = (attributes: string, body: string) => borderDeclarations(node(`<border xmlns="${S}" ${attributes}>${body}</border>`), colors, true);
type Case = { kind: string; pattern?: string; type?: string; degree?: number; left?: number; right?: number; top?: number; bottom?: number; rotation?: number; style?: string; up?: boolean; down?: boolean };
describe("native advanced cell appearance sources", () => {
  it("keeps all 42 independent appearances through native reading and package reopen", async () => {
    const bytes = readFileSync("packages/office/tests/fixtures/xlsx-cell-appearance.xlsx");
    const regenerated = execFileSync("python3", ["-c", "import importlib.util,sys;s=importlib.util.spec_from_file_location('appearance','scripts/check-xlsx-cell-appearance.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"]);
    expect(regenerated).toEqual(bytes);
    const evidence = JSON.parse(readFileSync("docs/evidence/xlsx-cell-appearance.json", "utf8")) as { sourceSha256: string; cases: Case[] };
    expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256); expect(evidence.cases).toHaveLength(42);
    const { model } = projectXlsxRecords(await readXlsxRecords(bytes)), snapshot = model.snapshot();
    for (const [index, input] of evidence.cases.entries()) {
      const cell = model.getCell(`sheet-${index + 1}`, 1, 1)!; expect(cell.cssClass).toBe(`xlsx-cell-${index + 1}`);
      const css = snapshot.styles[0]!.css.split(`.${cell.cssClass}{`)[1]!.split("}")[0]!;
      if (input.kind === "pattern") { expect(css).toContain(`--office-cell-fill-kind:${input.pattern};`); expect(css).toContain("--office-cell-fill-color:#204080;"); expect(css).toContain("--office-cell-fill-background-color:#ffe0a0;"); }
      if (input.kind === "gradient") {
        expect(css).toContain(`--office-cell-gradient-kind:${input.type};`); expect(css).toContain(`--office-cell-gradient-degree:${input.degree ?? 0};`);
        for (const side of ["left", "right", "top", "bottom"] as const) expect(css).toContain(`--office-cell-gradient-${side}:${input[side] ?? 0};`);
        expect(css).toContain("--office-cell-gradient-stops:3;");
        for (const [i, color] of ["#ff0000", "#00ff00", "#0000ff"].entries()) { expect(css).toContain(`--office-cell-gradient-stop-${i}-position:${i / 2};`); expect(css).toContain(`--office-cell-gradient-stop-${i}-color:${color};`); }
      }
      if (input.kind === "diagonal") { expect(css).toContain(`--office-cell-border-diagonal-kind:${input.style};`); expect(css).toContain(`--office-cell-border-diagonal-up:${input.up};`); expect(css).toContain(`--office-cell-border-diagonal-down:${input.down};`); }
      if (input.kind === "rotation") { expect(css).toContain(`--office-cell-text-rotation:${input.rotation};`); expect(cell.value).toEqual({ type: "text", value: "预算 Ab" }); }
    }
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(snapshot);
  });
  it("preserves gradient stop order, hard transitions, theme tints, native alpha and inactive diagonal styling", () => {
    const css = fill('<gradientFill><stop position="1"><color rgb="80204060"/></stop><stop position="0.5"><color theme="4" tint="0.2"/></stop><stop position="0.5"><color indexed="3"/></stop></gradientFill>');
    expect(css).toMatchObject({ "--office-cell-gradient-kind": "linear", "--office-cell-gradient-degree": "0", "--office-cell-gradient-stop-0-position": "1", "--office-cell-gradient-stop-0-native-alpha": "128",
      "--office-cell-gradient-stop-1-position": "0.5", "--office-cell-gradient-stop-1-color": "var(--office-theme-accent1)", "--office-cell-gradient-stop-1-tint": "0.2", "--office-cell-gradient-stop-2-position": "0.5" });
    expect(fill('<gradientFill/>')["--office-cell-gradient-stops"]).toBe("0");
    expect(border('', '<diagonal style="double"><color rgb="FF204060"/></diagonal>')).toMatchObject({ "--office-cell-border-diagonal-up": "false", "--office-cell-border-diagonal-down": "false", "--office-cell-border-diagonal-kind": "double", "--office-cell-border-diagonal-color": "#204060" });
  });
  it("rejects malformed gradients, unknown patterns and invalid rotations without discarding supported effects", () => {
    for (const body of ['<gradientFill type="unknown"/>', '<gradientFill degree="Infinity"/>', '<gradientFill left="-0.1"/>', '<gradientFill right="1.1"/>', '<gradientFill><stop position="-0.1"><color rgb="FF000000"/></stop></gradientFill>', '<gradientFill><stop><color rgb="FF000000"/></stop></gradientFill>', '<gradientFill><stop position="0"/></gradientFill>', '<gradientFill><stop position="0"><color rgb="FF000000"/><color rgb="FFFFFFFF"/></stop></gradientFill>', '<patternFill patternType="unknown"/>', '<patternFill/><gradientFill/>']) expect(() => fill(body)).toThrow();
    for (const rotation of [-1, 181, 254, 256, 1.5]) expect(() => alignmentDeclarations(node(`<alignment xmlns="${S}" textRotation="${rotation}"/>`), true)).toThrow();
    for (const rotation of [0, 1, 89, 90, 91, 179, 180, 255]) expect(alignmentDeclarations(node(`<alignment xmlns="${S}" textRotation="${rotation}"/>`), true)["--office-cell-text-rotation"]).toBe(String(rotation));
    expect(alignmentDeclarations(undefined, false)).toEqual({}); expect(alignmentDeclarations(undefined, true)["--office-cell-text-rotation"]).toBe("0");
  });
});
