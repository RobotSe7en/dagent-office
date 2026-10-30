import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { openDocument, DEFAULT_DOCUMENT_READ_LIMITS } from "@dagent-office/core";
import { exportDocx, importDocx, readDocumentZip } from "@dagent-office/office";
import { attr, child, elements, W, xml } from "../src/ooxml/xml.js";
import type { XmlNode } from "../src/ooxml/xml.js";
const encoder = new TextEncoder(), decoder = new TextDecoder();
const root = new URL("../../../examples/report/", import.meta.url);
function authored(body: string, css = "") {
  const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((name) => [name, new Uint8Array(readFileSync(new URL(name, root)))]));
  files.set("content.html", encoder.encode(decoder.decode(files.get("content.html")).replace(/<article[\s\S]*<\/article>/, '<article data-node-id="root" data-office-kind="document">' + body + '</article>')));
  files.set("annotations.json", encoder.encode("[]"));
  files.set("styles.css", encoder.encode('article{font-family:Arial;font-size:12pt}table{border-collapse:collapse}' + css));
  return openDocument(files);
}
function native(document: ReturnType<typeof openDocument>, path = "word/document.xml"): XmlNode {
  return xml(readDocumentZip(exportDocx(document).bytes).get(path)!, path, DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 });
}
const descendants = (node: XmlNode, name: string): XmlNode[] => node.content.flatMap((child) => typeof child === "string" ? [] : [...(child.uri === W && child.name === name ? [child] : []), ...descendants(child, name)]);
const firstRun = (node: XmlNode) => descendants(node, "rPr")[0];
for (const [name, body] of [
  ["cell", '<table data-node-id="t"><tr data-node-id="r"><td data-node-id="c">A<strong data-node-id="b">B</strong> C</td></tr></table>'],
  ["container", '<div data-node-id="d">A<strong data-node-id="b">B</strong> C</div>'],
  ["root", 'A<strong data-node-id="b">B</strong> C'],
  ["list", '<ul data-node-id="list"><li data-node-id="item"><strong data-node-id="b">A</strong> <em data-node-id="i">B</em></li></ul>'],
] as const) test("continuous inline source retains paragraph structure and spaces: " + name, () => {
  const model = authored(body), generated = exportDocx(model), tree = native(model);
  const paragraphs = name === "cell" ? descendants(descendants(tree, "tc")[0]!, "p") : elements(child(tree, "body"), "p");
  expect(paragraphs).toHaveLength(1);
  expect(descendants(paragraphs[0]!, "t").map((node) => node.content.join("")).join("")).toBe(name === "list" ? "A B" : "AB C");
  const reopened = openDocument(importDocx(generated.bytes, { documentId: "inline" }));
  const story = reopened.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!;
  expect(reopened.getNode(story.rootNodeId)!.text).toBe(name === "list" ? "A B" : "AB C");
});
test("pre-wrap whitespace-only source remains a native paragraph", () => {
  const model = authored('<div data-node-id="d" style="white-space:pre-wrap"> \t </div>'), tree = native(model);
  expect(descendants(tree, "tab")).toHaveLength(1);
  expect(descendants(tree, "t").map((node) => node.content.join("")).join("")).toBe("  ");
});
test("rem, inherited line height and inherited letter spacing use computed contexts", () => {
  let tree = native(authored('<p data-node-id="p" style="font-size:1rem">TEXT</p>', 'html{font-size:20pt}'));
  expect(attr(child(firstRun(tree), "sz"), "val")).toBe("40");
  tree = native(authored('<div data-node-id="d" style="font-size:10pt;line-height:200%"><p data-node-id="p" style="font-size:20pt;line-height:inherit">TEXT</p></div>'));
  expect(attr(child(descendants(tree, "pPr")[0], "spacing"), "line")).toBe("400");
  tree = native(authored('<div data-node-id="d" style="font-size:20pt;letter-spacing:.1em"><p data-node-id="p" style="font-size:10pt">TEXT</p></div>'));
  expect(attr(child(firstRun(tree), "spacing"), "val")).toBe("40");
});
test("currentColor inherits foreground and block decorations propagate through spans", () => {
  let tree = native(authored('<div data-node-id="d" style="color:red"><p data-node-id="p" style="color:currentColor">TEXT</p></div>'));
  expect(attr(child(firstRun(tree), "color"), "val")).toBe("ff0000");
  tree = native(authored('<p data-node-id="p" style="text-decoration:underline">A<span data-node-id="s" style="text-decoration:none">B</span>C</p>'));
  expect(descendants(tree, "r").filter((run) => child(run, "t")).map((run) => attr(child(child(run, "rPr"), "u"), "val"))).toEqual(["single", "single", "single"]);
});
test("new CSS style properties and relative paragraph dimensions reach native styles", () => {
  const source = openDocument(importDocx(new Uint8Array(readFileSync(new URL("../../../examples/docx/04-styles-sections.docx", import.meta.url))), { documentId: "styles" }));
  const derived = source.snapshot().metadata.word!.styles.find((style) => style.styleId === "Derived")!, files = new Map(source.serialize());
  files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")) + '\n' + ('.' + derived.cssClass).repeat(5) + '{font-size:30pt;margin-top:2em;color:#00ff00}'));
  const tree = native(openDocument(files), "word/styles.xml"), style = elements(tree, "style").find((style) => attr(style, "styleId") === "Derived")!;
  expect(attr(child(child(style, "rPr"), "sz"), "val")).toBe("60");
  expect(attr(child(child(style, "rPr"), "color"), "val")).toBe("00ff00");
  expect(attr(child(child(style, "pPr"), "spacing"), "before")).toBe("1200");
  const base = source.snapshot().metadata.word!.styles.find((style) => style.styleId === "Base")!;
  files.set("styles.css", encoder.encode(decoder.decode(source.serialize().get("styles.css")) + '\n' + ('.' + base.cssClass).repeat(5) + '{font-size:20pt;margin-left:2em}'));
  const baseStyle = elements(native(openDocument(files), "word/styles.xml"), "style").find((style) => attr(style, "styleId") === "Base")!;
  expect(attr(child(child(baseStyle, "pPr"), "ind"), "left")).toBe("800");
});
test("complex-script font overrides preserve the other source roles", () => {
  const tree = native(authored('<p data-node-id="p" style="--word-font-ascii:Arial;--word-font-hAnsi:Georgia;--word-font-eastAsia:SimSun;--word-font-cs:Amiri"><span data-node-id="s" data-word-script="cs" style="font-family:CustomCS">م</span></p>'));
  const fonts = child(firstRun(tree), "rFonts"); expect(attr(fonts, "cs")).toBe("CustomCS"); expect(attr(fonts, "ascii")).toBe("Arial"); expect(attr(fonts, "eastAsia")).toBe("SimSun");
});
test.each([
  '<p data-node-id="p" style="white-space:nowrap">TEXT</p>',
  '<p data-node-id="p"><span data-node-id="s" style="border:1pt solid red">TEXT</span></p>',
  '<table data-node-id="t"><tr data-node-id="r"><td data-node-id="c" style="height:100pt">TEXT</td></tr></table>',
])("unmapped visible formatting is explicit: %s", (body) => {
  expect(() => exportDocx(authored(body))).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
});
test.each([['A', '', 'upperLetter'], ['i', '', 'lowerRoman'], ['A', 'ol{list-style-type:lower-alpha}', 'lowerLetter']] as const)("HTML numbering type respects authored CSS: %s / %s", (type, css, expected) => {
  const tree = native(authored('<ol data-node-id="ol" type="' + type + '"><li data-node-id="li">ITEM</li></ol>', css), "word/numbering.xml");
  expect(attr(child(descendants(tree, "lvl")[0], "numFmt"), "val")).toBe(expected);
});
test("vertical merge expansion checks the budget before allocating its matrix", () => {
  const count = 900;
  const body = '<table data-node-id="t"><tr data-node-id="r0">' + Array.from({ length: count }, (_, index) => '<td data-node-id="c' + index + '" rowspan="' + count + '">X</td>').join('') + '</tr>' + Array.from({ length: count - 1 }, (_, index) => '<tr data-node-id="r' + (index + 1) + '"></tr>').join('') + '</table>';
  expect(() => exportDocx(authored(body), { limits: { maxNodes: 3000 } })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded", message: expect.stringContaining("table expansion") }));
});
test("inline spaces participate in exact comment mapping and NBSP stays visible", () => {
  const model = authored('<div data-node-id="d"><span data-node-id="a">A</span> <span data-node-id="b">B</span></div>'), files = new Map(model.serialize()), manifest = model.snapshot().manifest;
  const annotation = { annotationId: "range", documentId: manifest.documentId, baseRevision: manifest.revision, instruction: "范围", status: "open", anchor: { kind: "text-set", ranges: [{ nodeId: "a", from: 0, to: 1, quote: "A" }, { nodeId: "b", from: 0, to: 1, quote: "B" }] } };
  files.set("annotations.json", encoder.encode(JSON.stringify([annotation])));
  expect(exportDocx(openDocument(files)).report.unexportedAnnotations).toEqual([{ annotationId: "range", reason: "disjoint" }]);
  const continuous = { ...annotation, anchor: { kind: "text", nodeId: "d", from: 0, to: 3, quote: "A B" } };
  files.set("annotations.json", encoder.encode(JSON.stringify([continuous])));
  const reopened = openDocument(importDocx(exportDocx(openDocument(files)).bytes, { documentId: "range" })), anchor = reopened.snapshot().annotations[0]!.anchor;
  expect(anchor.kind === "text-set" ? anchor.ranges.map((range) => range.quote).join("") : anchor.kind === "text" ? anchor.quote : "").toBe("A B");
  expect(descendants(native(authored('\u00a0')), "t").map((node) => node.content.join("")).join("")).toBe("\u00a0");
});
test("an ancestor underline keeps its color across a differently colored descendant", () => {
  const tree = native(authored('<p data-node-id="p" style="color:red;text-decoration:underline">A<span data-node-id="s" style="color:blue;text-decoration:none">B</span></p>'));
  const run = descendants(tree, "r").find((run) => child(run, "t")?.content.join("") === "B")!;
  expect(attr(child(child(run, "rPr"), "color"), "val")).toBe("0000ff");
  expect(attr(child(child(run, "rPr"), "u"), "color")).toBe("ff0000");
});
test("unused source class names cannot supply an implicit native baseline", () => {
  const document = authored('<p data-node-id="p" style="margin-top:10pt">TEXT</p>', '.word-native-no-style{margin-top:10pt}.word-main{font-size:40pt}');
  expect(attr(child(descendants(native(document), "pPr")[0], "spacing"), "before")).toBe("200");
  const defaults = child(child(child(native(document, "word/styles.xml"), "docDefaults"), "rPrDefault"), "rPr");
  expect(attr(child(defaults, "sz"), "val")).toBe("24");
});
