import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, DEFAULT_DOCUMENT_READ_LIMITS, openDocument } from "@dagent-office/core";
import { openOpc, openOpcFiles } from "../src/ooxml/package.js";
import { S, xml } from "../src/ooxml/xml.js";
import { XlsxStyles } from "../src/xlsx/styles.js";

const encoder = new TextEncoder(), decoder = new TextDecoder(), options = { limits: { maxFileBytes: 32 * 1024 * 1024 } };
const dummy = xml(encoder.encode(`<c xmlns="${S}"/>`), "cell.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 });
function variant(edit: (source: string) => string) {
  const source = openOpc(readFileSync("examples/xlsx/02-styles-dates.xlsx"), options), files = new Map(source.files);
  files.set("xl/styles.xml", encoder.encode(edit(decoder.decode(files.get("xl/styles.xml")!)))); return openOpcFiles(files, options);
}
describe("native workbook style tables", () => {
  it("maps styles from all 32 independent/producer workbooks into validated CSS and number format definitions", () => {
    let checked = 0;
    for (const directory of ["examples/xlsx", "examples/xlsx/producer"]) for (const filename of readdirSync(directory).filter(name => name.endsWith(".xlsx"))) {
      const styles = new XlsxStyles(openOpc(readFileSync(`${directory}/${filename}`), options), "xl/workbook.xml", 4 * 1024 * 1024);
      const files = new Map(createWorkbook({ documentId: "style-tables" })), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!));
      metadata.workbook.themeCssClass = "xlsx-theme"; metadata.workbook.numberFormats = styles.numberFormats();
      metadata.workbook.columnWidthFontCssClass = styles.columnWidthFontCssClass();
      metadata.workbook.sheets[0].defaultCssClass = styles.cell(0, dummy).cssClass; metadata.workbook.sheets[0].defaultNumberFormatId = styles.cell(0, dummy).numberFormatId;
      files.set("metadata.json", encoder.encode(JSON.stringify(metadata))); files.set("styles.css", encoder.encode(styles.stylesheet()));
      const model = openDocument(files, { kind: "workbook" }); expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(model.snapshot()); checked++;
    }
    expect(checked).toBe(32);
  });
  it("keeps effective cell formats independent of future named-style update flags", () => {
    const styles = new XlsxStyles(variant(source => source.replace('numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"', 'numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="0"')), "xl/workbook.xml", 1_000_000);
    expect(styles.cell(2, dummy).numberFormatId).toBe("num-164"); expect(styles.cellFormats[2]?.apply.NumberFormat).toBe(false); expect(styles.cellFormats[2]?.parent).toBe(0);
    expect(styles.namedStyles).toEqual([{ name: "Normal", format: 0, builtinId: 0, hidden: false, customBuiltin: false }]);
    expect(styles.numberFormats()).toContainEqual({ formatId: "num-164", code: "#,##0.00;[Red]-#,##0.00" });
    expect(styles.stylesheet()).toContain("--office-cell-horizontal:center"); expect(styles.stylesheet()).toContain("--office-cell-border-left-kind:thin");
  });
  it("shares repeated rich-run styles and leaves unspecified differential properties inherited", () => {
    const styles = new XlsxStyles(openOpc(readFileSync("examples/xlsx/02-styles-dates.xlsx"), options), "xl/workbook.xml", 1_000_000);
    const source = xml(encoder.encode(`<rPr xmlns="${S}"><b/></rPr>`), "run.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 });
    const first = styles.run(source), size = styles.stylesheet().length; expect(styles.run(source)).toBe(first); expect(styles.stylesheet().length).toBe(size);
    const differential = styles.differential(0, dummy); expect(differential.numberFormatId).toBeUndefined();
    const css = styles.stylesheet().split(`.${differential.cssClass}{`)[1]!.split("}")[0]!;
    expect(css).toContain("#9c0006"); expect(css).toContain("#ffc7ce"); expect(css).not.toContain("font-weight");
  });
  it("binds column metrics to the built-in Normal style's font, independently of font zero and cell format zero", () => {
    const styles = new XlsxStyles(variant(source => source
      .replace(/<cellStyleXfs[\s\S]*?<\/cellStyleXfs>/, '<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellStyleXfs>')
      .replace('name="Normal" xfId="0" builtinId="0"', 'name="常规" xfId="1" builtinId="0" customBuiltin="1"')), "xl/workbook.xml", 1_000_000);
    expect(styles.columnWidthFontCssClass()).toBe('xlsx-named-1'); expect(styles.cell(0, dummy).cssClass).toBe('xlsx-cell-0');
    const normal = styles.stylesheet().split('.xlsx-named-1{')[1]!.split('}')[0]!, cell = styles.stylesheet().split('.xlsx-cell-0{')[1]!.split('}')[0]!;
    expect(normal).toContain('font-size:12pt'); expect(normal).toContain('--office-font-source-name:"Noto Sans CJK SC"'); expect(cell).toContain('font-size:11pt');
    const missing = new XlsxStyles(variant(source => source.replace(/<cellStyles[\s\S]*?<\/cellStyles>/, '')), 'xl/workbook.xml', 1_000_000);
    expect(() => missing.columnWidthFontCssClass()).toThrowError(expect.objectContaining({ code: 'UnsupportedFeature' }));
    expect(() => new XlsxStyles(variant(source => source.replace('<cellStyles count="1">', '<cellStyles count="2"><cellStyle name="Other" xfId="0" builtinId="0"/>')), 'xl/workbook.xml', 1_000_000)).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
  });
  it("rejects broken component IDs, duplicate definitions, active unsupported formats and style source over-budget", () => {
    const edits = [
      (s: string) => s.replace('fontId="1"', 'fontId="999"'),
      (s: string) => s.replace('numFmtId="164" fontId', 'numFmtId="999" fontId'),
      (s: string) => s.replace('<fonts count="3">', '<fonts count="4">'),
      (s: string) => s.replace('<alignment horizontal="center"', '<alignment textRotation="45" horizontal="center"'),
      (s: string) => s.replace('patternType="solid"', 'patternType="darkGrid"'),
      (s: string) => s.replace('<border><left/>', '<border diagonalUp="1"><left/>'),
    ];
    for (const edit of edits) expect(() => new XlsxStyles(variant(edit), "xl/workbook.xml", 1_000_000)).toThrow();
    expect(() => new XlsxStyles(variant(s => s), "xl/workbook.xml", 100)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
});
