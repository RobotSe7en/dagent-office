import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { openOpc, openOpcFiles } from "../src/ooxml/package.js";
import { XlsxStyles } from "../src/xlsx/styles.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const placeholder = '<a:schemeClr val="phClr"></a:schemeClr>';
const solid = '<a:solidFill>' + placeholder + '</a:solidFill>';
function styles(edit: (source: string) => string = source => source, maximum = 1_000_000): XlsxStyles {
  const files = new Map(openOpc(readFileSync("examples/xlsx/08-images-shapes.xlsx"), {}).files), path = "xl/theme/theme1.xml";
  files.set(path, encoder.encode(edit(decoder.decode(files.get(path)!))));
  return new XlsxStyles(openOpcFiles(files, {}), "xl/workbook.xml", maximum);
}
describe("workbook drawing theme definitions", () => {
  it("preserves shared style lists and per-object placeholder references in CSS", () => {
    const source = styles(), css = source.stylesheet();
    expect(source.drawingTheme).toEqual({
      fill: ["xlsx-theme-fill-1", "xlsx-theme-fill-2", "xlsx-theme-fill-3"],
      background: ["xlsx-theme-background-1", "xlsx-theme-background-2", "xlsx-theme-background-3"],
      line: ["xlsx-theme-line-1", "xlsx-theme-line-2", "xlsx-theme-line-3"],
      effect: ["xlsx-theme-effect-1", "xlsx-theme-effect-2", "xlsx-theme-effect-3"],
    });
    expect(css).toContain("--office-fill-reference-prefix:--office-fill-reference");
    expect(css).toContain("--office-line-reference-prefix:--office-line-reference");
    expect(css).toContain("--office-fill-color:var(--office-background-reference-color,var(--office-theme-dk1))");
    expect(css).toContain("--office-theme-tx1:var(--office-theme-dk1)");
    expect(css).not.toContain("--office-placeholder-color");
  });
  it("maps supported gradient, outline and shadow source values without flattening theme colors", () => {
    const source = styles(xml => xml.replace(solid, '<a:gradFill flip="none" rotWithShape="0"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="80000"/><a:lumOff val="10000"/><a:alpha val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>')
      .replace('<a:effectLst/>', '<a:effectLst><a:outerShdw blurRad="25400" dist="38100" dir="2700000"><a:schemeClr val="accent2"/></a:outerShdw></a:effectLst>'));
    const css = source.stylesheet();
    expect(css).toContain("--office-fill-kind:linear"); expect(css).toContain("--office-fill-angle:90");
    expect(css).toContain("--office-fill-stop-0-lum-mod:0.8"); expect(css).toContain("--office-fill-stop-0-alpha:0.5");
    expect(css).toContain("--office-shadow-blur:2pt"); expect(css).toContain("--office-shadow-distance:3pt");
    expect(css).toContain("--office-shadow-color:var(--office-theme-accent2)");
    expect(css).toContain("--office-fill-scaled:true"); expect(css).toContain("--office-fill-rotate-with-shape:false");
  });
  it("keeps omitted native gradient attributes available for theme/style inheritance", () => {
    const source = styles(xml => xml.replace(solid, '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin/></a:gradFill>'));
    const css = source.stylesheet(); expect(css).toContain('--office-fill-kind:linear');
    for (const field of ['--office-fill-angle:', '--office-fill-scaled:', '--office-fill-rotate-with-shape:']) expect(css).not.toContain(field);
  });
  it("rejects unknown, repeated, unrepresentable or dangling native style definitions", () => {
    const edits = [
      (s: string) => s.replace('<a:fmtScheme name=', '<a:fmtScheme unknown="1" name='),
      (s: string) => s.replace('<a:ln w="12700"', '<a:ln w="12700" cap="invalid"'),
      (s: string) => s.replace('<a:ln w="12700"', '<a:ln w="99999999"'),
      (s: string) => s.replace('<a:prstDash val="solid"/>', '<a:round/><a:bevel/>'),
      (s: string) => s.replace(placeholder, '<a:schemeClr val="phClr"><a:lumOff val="10000"/><a:lumMod val="80000"/></a:schemeClr>'),
      (s: string) => s.replace(placeholder, '<a:schemeClr val="phClr"><a:tint val="10000"/><a:tint val="20000"/></a:schemeClr>'),
      (s: string) => s.replace(placeholder, '<a:schemeClr val="phClr" unknown="1"/>'),
      (s: string) => s.replace('<a:effectLst/>', '<a:effectLst><a:glow rad="12700"/></a:effectLst>'),
      (s: string) => s.replace('</a:theme>', '<a:objectDefaults><a:spDef/></a:objectDefaults></a:theme>'),
      (s: string) => s.replace('</a:theme>', '<a:extLst><a:ext uri="unknown"/></a:extLst></a:theme>'),
      (s: string) => s.replace('<a:fillStyleLst>', '<a:fillStyleLst><a:solidFill/>'),
      (s: string) => s.replace(solid, ''),
    ];
    for (const [index, edit] of edits.entries()) expect(() => styles(edit), `mutation ${index}`).toThrow();
    expect(() => styles(s => s, 3000)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
});
