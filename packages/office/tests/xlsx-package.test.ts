import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { openOpc, openOpcFiles } from "../src/ooxml/package.js";
import { auditXlsxPackage } from "../src/xlsx/package.js";
import { R } from "../src/ooxml/xml.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
function source() { return new Map(openOpc(readFileSync("examples/xlsx/01-values-text.xlsx"), {}).files); }
function update(files: Map<string, Uint8Array>, path: string, fn: (text: string) => string) { files.set(path, encoder.encode(fn(decoder.decode(files.get(path)!)))); }
describe("complete native workbook container audit", () => {
  it("rejects unsupported orphan parts, supported disconnected parts and unowned relationship files", () => {
    for (const kind of ["macro", "image", "relationship"]) {
      const files = source();
      if (kind === "macro") files.set("xl/vbaProject.bin", new Uint8Array([0]));
      if (kind === "image") files.set("xl/media/orphan.png", new Uint8Array([0]));
      if (kind === "relationship") files.set("ghost/_rels/sheet.xml.rels", files.get("_rels/.rels")!);
      expect(() => auditXlsxPackage(openOpcFiles(files, {}))).toThrow();
    }
  });
  it("does not trust relationship suffixes, foreign child namespaces or mismatched target content types", () => {
    const changes = [
      (s: string) => s.replace(R + "/worksheet", "https://attacker.invalid/worksheet"),
      (s: string) => s.replace('Type="' + R + '/worksheet"', 'Type="' + R + '/theme"'),
      (s: string) => s.replace("</Relationships>", '<x:Relationship xmlns:x="urn:unexpected"/></Relationships>'),
      (s: string) => s.replace("</Relationships>", '<Relationship Id="external" Type="' + R + '/theme" Target="https://example.invalid/theme.xml" TargetMode="External"/></Relationships>'),
    ];
    for (const change of changes) {
      const files = source(); update(files, "xl/_rels/workbook.xml.rels", change);
      expect(() => auditXlsxPackage(openOpcFiles(files, {}))).toThrow();
    }
  });
  it("audits hidden content-type children and allows the declared external hyperlink without fetching it", () => {
    auditXlsxPackage(openOpcFiles(source(), {}));
    const files = source(); update(files, "[Content_Types].xml", s => s.replace("</Types>", '<other xmlns="urn:unhandled"/></Types>'));
    expect(() => auditXlsxPackage(openOpcFiles(files, {}))).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  });
});
