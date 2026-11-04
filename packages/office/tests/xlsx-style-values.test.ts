import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, DEFAULT_DOCUMENT_READ_LIMITS, openDocument, WorkbookNumberFormatSchema } from "@dagent-office/core";
import { declarations } from "../src/ooxml/styles.js";
import { openOpc } from "../src/ooxml/package.js";
import { S, xml } from "../src/ooxml/xml.js";
import { SpreadsheetColors, colorDeclarations, themeSlots } from "../src/xlsx/colors.js";
import { readWorkbookTheme } from "../src/xlsx/theme.js";
import { fontDeclarations } from "../src/xlsx/font.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const node = (source: string) => xml(encoder.encode(source), "styles.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 });
const colors = () => new SpreadsheetColors(undefined, new Set(themeSlots));
describe("SpreadsheetML appearance source values", () => {
  it("keeps indexed/theme references and signed luminance tint, including native ARGB bytes", () => {
    const palette = colors(), read = (attributes: string) => palette.read(node(`<color xmlns="${S}" ${attributes}/>`));
    expect(palette.css["--office-indexed-0"]).toBe("#000000"); expect(palette.css["--office-indexed-9"]).toBe("#ffffff"); expect(palette.css["--office-indexed-63"]).toBe("#333333");
    expect(read('theme="1" tint="-0.2"')).toEqual({ css: "var(--office-theme-dk1)", tint: -0.2 });
    expect(read('theme="0"')).toEqual({ css: "var(--office-theme-lt1)", tint: 0 });
    expect(read('indexed="64"')).toEqual({ css: "var(--office-system-foreground)", tint: 0 });
    expect(read('auto="1" indexed="2" rgb="80FF0000" theme="4"')).toEqual({ css: "var(--office-theme-accent1)", tint: 0 });
    expect(read('rgb="00FF0000"')).toEqual({ css: "#ff0000", tint: 0, nativeAlpha: 0 });
    expect(colorDeclarations("--test", read('rgb="804472C4" tint="0.4"'))).toEqual({ "--test-color": "#4472c4", "--test-tint": "0.4", "--test-native-alpha": "128" });
    for (const attrs of ['theme="20"', 'indexed="66"', 'tint="1.1" rgb="FFFFFFFF"', 'rgb="ffffff"', 'theme="4" rgb="invalid"', 'auto="false"']) expect(() => read(attrs)).toThrow();
  });
  it("preserves custom indexed palettes and rejects dangling color indexes", () => {
    const palette = new SpreadsheetColors(node(`<colors xmlns="${S}"><indexedColors><rgbColor rgb="FF123456"/></indexedColors></colors>`), new Set());
    expect(palette.css["--office-indexed-0"]).toBe("#123456"); expect(palette.read(node(`<color xmlns="${S}" indexed="0"/>`)).css).toBe("var(--office-indexed-0)");
    expect(() => palette.read(node(`<color xmlns="${S}" indexed="1"/>`))).toThrow();
    expect(() => palette.read(node(`<color xmlns="${S}" theme="0"/>`))).toThrow();
  });
  it.each(["examples/xlsx/02-styles-dates.xlsx", "examples/xlsx/producer/02-styles-dates.xlsx"])("maps actual theme font roles into valid isolated CSS from %s", filename => {
    const opc = openOpc(readFileSync(filename), {}), theme = readWorkbookTheme(opc, "xl/workbook.xml");
    expect(theme.css["--office-theme-accent1"]).toBe("#4472c4"); expect(theme.css["--office-minor-ascii"]).toBe('"Liberation Sans"');
    expect(theme.css["--office-minor-east-asia"]).toContain("Noto Sans CJK SC");
    const font = fontDeclarations(node(`<font xmlns="${S}"><name val="Liberation Sans"/><scheme val="minor"/><color theme="4" tint="-0.2"/><b/><sz val="12"/></font>`), new SpreadsheetColors(undefined, theme.colors), true);
    expect(font["--office-font-source-name"]).toBe('"Liberation Sans"'); expect(font["--office-font-east-asia"]).toBe('var(--office-minor-east-asia,"Liberation Sans")');
    const files = new Map(createWorkbook({ documentId: "theme" })), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!));
    metadata.workbook.themeCssClass = "book-theme"; metadata.workbook.numberFormats.push({ formatId: "locale-date", builtinId: 14 }); metadata.workbook.sheets[0].tabCssClass = "cell-default";
    files.set("metadata.json", encoder.encode(JSON.stringify(metadata))); files.set("styles.css", encoder.encode(`.book-theme{${declarations({ ...theme.css, ...colors().css })}}.cell-default{${declarations(font)}}`));
    const model = openDocument(files, { kind: "workbook" }); expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(model.snapshot());
    expect(WorkbookNumberFormatSchema.safeParse({ formatId: "both", code: "0.00", builtinId: 2 }).success).toBe(false);
  });
  it("preserves independent rich-run overrides and safely quotes source font names", () => {
    const partial = fontDeclarations(node(`<rPr xmlns="${S}"><strike/><rFont val="A&quot;B; color:red; \\ C"/></rPr>`), colors(), false);
    expect(partial["--office-font-strike"]).toBe("true"); expect(partial["--office-font-underline"]).toBeUndefined(); expect(partial["--office-font-scheme"]).toBe("none");
    expect(partial["--office-font-source-name"]).toBe('"A\\"B; color:red; \\\\ C"');
    const underline = fontDeclarations(node(`<rPr xmlns="${S}"><u val="doubleAccounting"/><i val="0"/></rPr>`), colors(), false);
    expect(underline["--office-font-underline"]).toBe("doubleAccounting"); expect(underline["--office-font-strike"]).toBeUndefined(); expect(underline["font-style"]).toBe("normal");
    for (const source of ['<name val=""/>', '<sz val="0"/>', '<b/><b/>', '<scheme val="unknown"/>', '<shadow/>', '<color theme="4" tint="NaN"/>']) expect(() => fontDeclarations(node(`<font xmlns="${S}">${source}</font>`), colors(), true)).toThrow();
  });
});
