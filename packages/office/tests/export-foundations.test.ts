import { expect, test } from "vitest";
import { DEFAULT_DOCUMENT_READ_LIMITS } from "@dagent-office/core";
import { ExportCss } from "../src/export/css.js";
import { parseSource } from "../src/export/source.js";
import { color, fontName, length } from "../src/export/values.js";
import { escapeXml, XmlParts } from "../src/export/xml.js";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";
const limits = DEFAULT_DOCUMENT_READ_LIMITS;
function source(body: string): ReturnType<typeof parseSource> {
  return parseSource('<!doctype html><html><head></head><body><article data-node-id="root" data-office-kind="document">' + body + '</article></body></html>', limits);
}
test("native compilation observes descendant/data selectors, important cascade and shorthand order", () => {
  const tree = source('<section data-node-id="section" class="parent"><p data-node-id="p" data-role="selected" class="special" style="color:purple;margin-left:4pt!important">TEXT</p></section>');
  const styles = new ExportCss([':root{--accent:#123456}.parent p[data-role=selected]{color:var(--accent)!important;margin:1pt 2pt}.special{color:red;margin-left:9pt!important}'], limits);
  const result = styles.get(tree.byId.get("p")!);
  expect(result.color).toBe("#123456"); expect(result["margin-left"]).toBe("4pt"); expect(result["margin-top"]).toBe("1pt"); expect(result["margin-right"]).toBe("2pt");
});
test("custom properties inherit computed values and cyclic fallback dependencies remain invalid", () => {
  const tree = source('<section data-node-id="parent" style="--a:var(--b);--b:red;color:blue"><p data-node-id="child" style="--b:green;color:var(--a)">TEXT</p><p data-node-id="cycle" style="--x:var(--b,var(--x));color:var(--x,purple)">TEXT</p><p data-node-id="invalid" style="--bad:6pt;color:var(--bad)">TEXT</p></section>');
  const styles = new ExportCss([], limits);
  expect(styles.get(tree.byId.get("child")!).color).toBe("red");
  expect(styles.get(tree.byId.get("cycle")!).color).toBe("purple");
  expect(styles.get(tree.byId.get("invalid")!).color).toBe("blue");
});
test("source font roles and explicit revision formatting survive without baking review markup into content", () => {
  const tree = source('<div data-node-id="main" class="word-main" style="--word-font-ascii:Georgia;--word-font-eastAsia:宋体"><ins data-node-id="revision"><span data-node-id="run" style="font-family:var(--word-font-eastAsia)">汉󠄀</span></ins><ins data-node-id="explicit" style="color:blue">TEXT</ins></div>');
  const styles = new ExportCss(['.word-main ins{text-decoration:underline;color:#006400}'], limits, new Set(["revision", "explicit"]));
  const run = styles.get(tree.byId.get("run")!);
  expect(fontName(run["font-family"]!)).toBe("宋体"); expect(run["--word-font-ascii"]).toBe("Georgia"); expect(run.color).toBe("#000000"); expect(run["text-decoration-line"]).toBe("none");
  expect(styles.get(tree.byId.get("explicit")!).color).toBe("blue");
});
test("CSS arithmetic uses explicit page/font context and refuses nonnative units", () => {
  expect(length("calc(100% - 2em)", { font: 10, percent: 200 })).toBe(180);
  expect(length("clamp(1cm,50%,2in)", { font: 10, percent: 200 })).toBe(100);
  expect(length("calc((12px + 2pt) * 2)", { font: 10, percent: 200 })).toBe(22);
  expect(() => length("1vw", { font: 10, percent: 200 })).toThrow(/cannot be represented/);
  expect(color("hsl(240 100% 50%)")).toBe("0000ff"); expect(color("rgb(50% 0% 100%)")).toBe("8000ff");
  expect(() => color("rgba(1,2,3,.5)")).toThrow(/cannot be represented/);
  expect(fontName('"仿宋_GB2312",serif')).toBe("仿宋_GB2312");
});
test("XML preserves astral UTF-16 and variation sequences, rejects invalid scalars and stops before oversized output", () => {
  expect(escapeXml("𠀀漢󠄀&\r")).toBe("𠀀漢󠄀&amp;&#13;");
  expect(() => escapeXml("\ud800")).toThrow(/XML 1.0/); expect(() => escapeXml("\uffff")).toThrow(/XML 1.0/);
  const parts = new XmlParts({ ...limits, maxFileBytes: 5 }), output = parts.buffer("word/document.xml");
  output.add("12345"); expect(() => output.add("6")).toThrow(/budget/); expect(parts.files.size).toBe(0);
});
test("native OPC writing shares deterministic raw ZIP rules without requiring an HTML manifest", () => {
  const files = new Map([["word/document.xml", new Uint8Array([1, 2, 3])]]), first = writeOpcZip(files), second = writeOpcZip(files);
  expect(first).toEqual(second); expect(readDocumentZip(first).get("word/document.xml")).toEqual(files.get("word/document.xml"));
  expect(() => writeOpcZip(new Map([["../escape.xml", new Uint8Array([1])]]))).toThrow();
});
test("variable expansion has a bound before a repeated dependency graph grows exponentially", () => {
  const tree = source('<p data-node-id="p">TEXT</p>');
  const declarations = ['--v0:"A"', ...Array.from({ length: 20 }, (_, index) => '--v' + (index + 1) + ':var(--v' + index + ') var(--v' + index + ')')].join(';');
  const styles = new ExportCss(['p{' + declarations + ';font-family:var(--v20)}'], { ...limits, maxFileBytes: 2000 });
  expect(() => styles.get(tree.byId.get("p")!)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
});
