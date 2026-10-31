import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, DocumentError, openDocument, readPresentationGeometry } from "@dagent-office/core";
import { importPptx, readDocumentZip, writeDocumentZip } from "@dagent-office/office";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const directory = new URL("../../../examples/pptx/", import.meta.url);
const expected = JSON.parse(readFileSync(new URL("expected.json", directory), "utf8")) as {
  cases: { file: string; text: string[]; slides: number; objects?: number }[];
  negativeCases: { file: string; error: string }[];
};
const fixture = (name = "01-rich-text.pptx") => new Uint8Array(readFileSync(new URL(name, directory)));
const load = (name = "01-rich-text.pptx") => openDocument(importPptx(fixture(name), { documentId: "ppt-test" }));
function mutate(name: string, part: string, change: (source: string) => string): Uint8Array {
  const files = unzipSync(fixture(name)); files[part] = strToU8(change(strFromU8(files[part]!))); return zipSync(files);
}
const readingLater = /^(09|10|11|12|13)-/;
for (const item of expected.cases.filter(c => !readingLater.test(c.file))) for (const producer of ["", "producer/"]) {
  test("imports independent PPTX base objects and reopens the HTML package: " + producer + item.file, () => {
    const bytes = fixture(producer + item.file), before = new Uint8Array(bytes);
    const files = importPptx(bytes, { documentId: "ppt-test" }), model = openDocument(files), snapshot = model.snapshot();
    expect(bytes).toEqual(before);
    expect(snapshot.manifest).toMatchObject({ documentId: "ppt-test", revision: 0, kind: "presentation" });
    expect(snapshot.metadata.presentation!.slides).toHaveLength(item.slides);
    const slideText = snapshot.metadata.presentation!.slides.map(slide => model.getNode(slide.rootNodeId)!.text).join("\n");
    for (const text of item.text) expect(slideText).toContain(text);
    if (item.objects !== undefined) expect(snapshot.metadata.presentation!.objects).toHaveLength(item.objects);
    expect(openDocument(model.serialize()).snapshot()).toEqual(snapshot);
    expect(openDocument(readDocumentZip(writeDocumentZip(files))).snapshot()).toEqual(snapshot);
    expect(importPptx(bytes, { documentId: "ppt-test" })).toEqual(files);
    expect([...files.keys()].some(name => name.endsWith(".pptx"))).toBe(false);
  });
}

test("keeps source font roles, Unicode variation sequences, paragraphs, lists and links", () => {
  const model = load(), snapshot = model.snapshot(), p = snapshot.metadata.presentation!;
  const slide = model.getNode(p.slides[0]!.rootNodeId)!;
  expect(slide.text).toContain("扩展字𠀀与变体葛󠄀");
  expect(slide.text).toContain("换行前\n换行后");
  expect(snapshot.html).toContain('data-office-numbering="arabicPeriod"');
  expect(snapshot.html).toContain('data-office-bullet="•"');
  expect(snapshot.html).toContain("font-weight:bold");
  expect(snapshot.html).toContain("--office-font-east-asia");
  expect(snapshot.html).toContain("Noto Sans CJK SC");
  expect(p.hyperlinks).toEqual([expect.objectContaining({ target: "https://example.com/pptx" })]);
  expect(snapshot.html).not.toContain('href="https://example.com');
});

test("preserves placeholder inheritance without materializing a second transform", () => {
  const model = load("02-theme-placeholders.pptx"), p = model.snapshot().metadata.presentation!;
  expect(p.masters).toHaveLength(2);
  const slide = model.getNode(p.slides[0]!.rootNodeId)!, object = model.getNode(slide.childNodeIds[0]!)!;
  expect(readPresentationGeometry(object.attributes)).toBeUndefined();
  const layout = p.layouts.find(l => l.layoutId === p.slides[0]!.layoutId)!;
  const inherited = model.getNode(model.getNode(layout.rootNodeId)!.childNodeIds[0]!)!;
  expect(readPresentationGeometry(inherited.attributes)).toMatchObject({ x: 60, y: 40, width: 840, height: 100 });
  expect(model.snapshot().styles[0]!.css).toContain("--office-palette-accent1:#c0504d");
});

test("resolves percent-encoded part names without relying on the reference application's path support", () => {
  const files = unzipSync(fixture("17-renamed-parts.pptx"));
  for (const name of Object.keys(files)) {
    const target = name.replace("page-one.xml", "page one.xml");
    const bytes = files[name]!;
    if (target !== name) delete files[name];
    files[target] = name.endsWith(".xml") || name.endsWith(".rels") ? strToU8(strFromU8(bytes).replaceAll("page-one.xml", "page%20one.xml")) : bytes;
  }
  const model = openDocument(importPptx(zipSync(files), { documentId: "encoded" }));
  const p = model.snapshot().metadata.presentation!;
  expect(model.getNode(p.slides[0]!.rootNodeId)!.text).toBe("命名空间不依赖前缀");
});

test("keeps nested local transforms, negative coordinates, rotation and flips", () => {
  const model = load("05-groups.pptx"), p = model.snapshot().metadata.presentation!;
  const group = p.objects.find(o => o.kind === "group" && o.nativeId === "2")!;
  const node = model.getNode(group.nodeId)!;
  expect(node.attributes).toMatchObject({ "data-office-child-x": "-50", "data-office-child-y": "-25", "data-office-child-width": "900", "data-office-child-height": "450" });
  expect(readPresentationGeometry(node.attributes)).toMatchObject({ rotation: 10, flipH: true });
  expect(readPresentationGeometry(model.getNode(p.objects.find(o => o.nativeId === "5")!.nodeId)!.attributes)).toMatchObject({ x: -20, y: 0 });
});

test("resolves connector identities after all objects are mapped and retains native elbow adjustments", () => {
  for (const prefix of ["", "producer/"]) {
    const model = load(prefix + "06-connectors.pptx"), p = model.snapshot().metadata.presentation!;
    const a = p.objects.find(o => o.nativeId === "2")!, b = p.objects.find(o => o.nativeId === "3")!;
    for (const connector of p.objects.filter(o => o.kind === "connector")) expect(connector).toMatchObject({ startNodeId: a.nodeId, endNodeId: b.nodeId, startSite: 1, endSite: 3 });
    if (prefix) expect(model.snapshot().html).toContain('data-office-connector-preset="bentConnector4"');
  }
});

test("retains image bytes, crop and accessible description without fetching external resources", () => {
  const bytes = fixture("07-images.pptx"), native = unzipSync(bytes), files = importPptx(bytes, { documentId: "images" });
  const model = openDocument(files), snapshot = model.snapshot();
  expect(snapshot.manifest.assets.map(a => a.mimeType).sort()).toEqual(["image/jpeg", "image/png", "image/webp"]);
  for (const asset of snapshot.manifest.assets) expect(files.get(asset.path)).toEqual(native[asset.path.slice("assets/".length)]);
  expect(snapshot.html).toContain('data-office-crop-left="0.1"');
  expect(snapshot.html).toContain('alt="公开合成图片"');
});

test("maps native merged table cells to one HTML cell per merge origin", () => {
  for (const prefix of ["", "producer/"]) {
    const model = load(prefix + "08-table-merges.pptx"), p = model.snapshot().metadata.presentation!;
    const table = p.objects.find(o => o.kind === "table")!;
    if (table.kind !== "table") throw new Error("Expected table");
    // LibreOffice quantizes these native grid widths on resave; retain its actual
    // coordinates rather than snapping them back to the synthetic source.
    expect(table.grid).toEqual(Array(3).fill(prefix ? 239.98110236220472 : 240));
    expect(model.getNode(table.nodeId)!.html).toContain('colspan="2"');
    expect(model.getNode(table.nodeId)!.html).toContain('rowspan="2"');
    expect(model.getNode(table.nodeId)!.html.match(/<td /g)).toHaveLength(7);
  }
});

for (const item of expected.negativeCases) test("rejects malformed or unsupported native PPTX: " + item.file, () => {
  try { importPptx(fixture(item.file), { documentId: "invalid" }); throw new Error("Accepted invalid source"); }
  catch (error) { expect(error).toBeInstanceOf(DocumentError); expect((error as DocumentError).code).toBe(item.error); }
});

test("rejects wrong types, cyclic ownership, unsupported geometry and broken connection targets", () => {
  const cases: [string, string, (s: string) => string, string][] = [
    ["01-rich-text.pptx", "[Content_Types].xml", s => s.replace("presentationml.slide+xml", "presentationml.slideLayout+xml"), "InvalidDocument"],
    ["01-rich-text.pptx", "ppt/slideLayouts/_rels/slideLayout1.xml.rels", s => s.replace("../slideMasters/slideMaster1.xml", "../slideLayouts/slideLayout1.xml"), "InvalidDocument"],
    ["01-rich-text.pptx", "ppt/slides/slide1.xml", s => s.replace('prst="rect"', 'prst="cloud"'), "UnsupportedFeature"],
    ["06-connectors.pptx", "ppt/slides/slide1.xml", s => s.replace('<a:stCxn id="2"', '<a:stCxn id="999"'), "InvalidDocument"],
    ["08-table-merges.pptx", "ppt/slides/slide1.xml", s => s.replace('gridSpan="2"', 'gridSpan="99999999"'), "InvalidDocument"],
  ];
  for (const [name, part, change, code] of cases) expect(() => importPptx(mutate(name, part, change), { documentId: "bad" })).toThrowError(expect.objectContaining({ code }));
});

test("enforces byte/node/ZIP budgets, identities, format separation and the editing gate", () => {
  expect(() => importPptx(fixture(), { documentId: "" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  expect(() => importPptx(fixture(), { documentId: "small", maxZipBytes: 10 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => importPptx(fixture(), { documentId: "small", limits: { maxNodes: 30 } })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  const docx = readFileSync(new URL("../../../examples/docx/01-text-numbering.docx", import.meta.url));
  expect(() => importPptx(docx, { documentId: "wrong" })).toThrowError(expect.objectContaining({ code: "UnsupportedFormat" }));
  expect(() => createDocumentSession(load())).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
});

test("compares native numeric identities by value, including leading-zero aliases and unsigned bounds", () => {
  const duplicate = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => {
    const shape = source.match(/<p:sp>.*?<\/p:sp>/s)![0];
    return source.replace("</p:spTree>", shape.replace('id="2"', 'id="02"') + "</p:spTree>");
  });
  const outside = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => source.replace('id="2"', 'id="4294967296"'));
  const masters = mutate("02-theme-placeholders.pptx", "ppt/presentation.xml", source => source.replace('id="2147483649"', 'id="2147483648"'));
  for (const bytes of [duplicate, outside, masters]) expect(() => importPptx(bytes, { documentId: "identity" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});

test("rejects a table continuation whose direction disagrees with its merge origin", () => {
  const bytes = mutate("08-table-merges.pptx", "ppt/slides/slide1.xml", source => source.replace('vMerge="1"', 'hMerge="1"'));
  expect(() => importPptx(bytes, { documentId: "merge" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});

test("preserves the reference color and transforms for a theme background", () => {
  const bytes = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => source.replace('<p:spTree>', '<p:bg><p:bgRef idx="1001"><a:srgbClr val="123456"><a:alpha val="50000"/></a:srgbClr></p:bgRef></p:bg><p:spTree>'));
  const snapshot = openDocument(importPptx(bytes, { documentId: "background" })).snapshot();
  expect(snapshot.html).toContain("--office-background-reference-color:#123456");
  expect(snapshot.html).toContain("--office-background-reference-alpha:0.5");
  expect(snapshot.styles[0]!.css).toContain("var(--office-background-reference-color");
});

test("preserves absent text-body overrides for placeholder inheritance instead of manufacturing defaults", () => {
  const input = mutate("02-theme-placeholders.pptx", "ppt/slides/slide1.xml", source => source.replace(/<a:bodyPr[^>]*>[\s\S]*?<\/a:bodyPr>/g, "<a:bodyPr/>"));
  const model = openDocument(importPptx(input, { documentId: "body-inheritance" }));
  const object = model.snapshot().metadata.presentation!.objects.find(o => o.kind === "text" && model.getNode(o.nodeId)!.text.includes("继承主题标题"))!;
  expect(object).toMatchObject({ text: {} });
  const node = model.getNode(object.nodeId)!; expect(node.attributes.style ?? "").not.toMatch(/padding-|white-space|text-anchor|tab-size/);
});
