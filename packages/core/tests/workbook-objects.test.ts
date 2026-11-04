import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createAssetDescriptor, createWorkbook, openDocument } from "../src/index.js";
import type { WorkbookDocumentMetadata, WorkbookManifest, WorkbookObject } from "../src/index.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
function seeded(edit?: (object: WorkbookObject, metadata: WorkbookDocumentMetadata) => void, markup?: (source: string) => string) {
  const files = new Map(createWorkbook({ documentId: "objects" })), manifest = JSON.parse(decoder.decode(files.get("manifest.json")!)) as WorkbookManifest;
  const metadata = JSON.parse(decoder.decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  const bytes = new Uint8Array(readFileSync("examples/presentation/assets/sample.png")), asset = createAssetDescriptor("assets/sample.png", "image/png", bytes);
  manifest.content = "content.html"; manifest.assets.push(asset); files.set(asset.path, bytes);
  const object: WorkbookObject = { nodeId: "shape", sheetId: "sheet-1", kind: "geometry", preset: "roundRect", locked: true, printable: true, nativeId: "1",
    anchor: { kind: "two-cell", from: { row: 2, column: 3, offsetXPt: -1, offsetYPt: 30 }, to: { row: 8, column: 7, offsetXPt: 10, offsetYPt: 0 }, moveWithCells: true, sizeWithCells: false }, text: { autoFit: "none", wrap: true } };
  metadata.workbook.objects = [object, { nodeId: "image", sheetId: "sheet-1", kind: "image", locked: false, printable: false, nativeId: "2", anchor: { kind: "absolute", xPt: 0, yPt: 0, widthPt: 100, heightPt: 75 } }];
  edit?.(object, metadata);
  const content = '<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body><section data-node-id="content" data-office-kind="workbook-content"><div data-node-id="shape" data-office-kind="geometry" data-office-rotation="-10" data-office-flip-h="true" data-office-adjust-adj="16667" class="cell-default"><p data-node-id="paragraph"><span data-node-id="run">中文😀禰󠄀</span></p></div><div data-node-id="image" data-office-kind="image" data-office-crop-left="0.1" data-office-crop-right="0.1"><img data-node-id="bitmap" src="assets/sample.png" alt="图片"></div></section></body></html>';
  files.set("content.html", encoder.encode(markup ? markup(content) : content));
  files.set("manifest.json", encoder.encode(JSON.stringify(manifest))); files.set("metadata.json", encoder.encode(JSON.stringify(metadata))); return files;
}
describe("workbook auxiliary drawing source", () => {
  it("keeps anchored objects, exact text, orientation and authorized image bytes through a package roundtrip", () => {
    const model = openDocument(seeded(), { kind: "workbook" }), snapshot = model.snapshot();
    expect(snapshot.manifest.chunks).toEqual([]); expect(snapshot.metadata.workbook.objects).toHaveLength(2);
    expect(model.getNode("shape")?.text).toBe("中文😀禰󠄀"); expect(model.getNode("shape")?.attributes["data-office-rotation"]).toBe("-10");
    expect(Object.isFrozen(snapshot.metadata.workbook.objects?.[0]?.anchor)).toBe(true);
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(snapshot);
  });
  it("rejects foreign sheets, duplicate ownership, dangling sources and duplicate native identities", () => {
    const edits: Parameters<typeof seeded>[0][] = [
      object => { object.sheetId = "missing"; }, object => { object.nodeId = "missing"; },
      (object, metadata) => { metadata.workbook.objects!.push(object); },
      (_, metadata) => { metadata.workbook.objects![1]!.nativeId = "1"; },
      object => { if (object.anchor.kind === "two-cell") object.anchor.from.row = 1048576; },
    ];
    for (const edit of edits) expect(() => openDocument(seeded(edit))).toThrow();
    expect(() => openDocument(seeded((_, metadata) => { metadata.workbook.objects = []; }))).toThrow();
  });
  it("checks crop visibility, preset adjustments, rich-body ownership and the unique anchor source", () => {
    for (const edit of [
      (s: string) => s.replace('data-office-rotation="-10"', 'data-office-rotation="NaN"'),
      (s: string) => s.replace('data-office-flip-h="true"', 'data-office-flip-h="1"'),
      (s: string) => s.replace('data-office-adjust-adj="16667"', 'data-office-adjust-surprise="1"'),
      (s: string) => s.replace('data-office-crop-right="0.1"', 'data-office-crop-right="0.9"'),
      (s: string) => s.replace('data-office-rotation="-10"', 'data-office-x="50"'),
      (s: string) => s.replace('<p data-node-id="paragraph">', 'outside<p data-node-id="paragraph">'),
      (s: string) => s.replace('<p data-node-id="paragraph">', '<p data-node-id="paragraph" data-office-kind="text">'),
      (s: string) => s.replace('src="assets/sample.png"', 'src="https://example.test/private.png"'),
      (s: string) => s.replace('class="cell-default"', 'class="missing"'),
      (s: string) => s.replace('<img data-node-id="bitmap"', '<br data-node-id="extra"><img data-node-id="bitmap"'),
    ]) expect(() => openDocument(seeded(undefined, edit))).toThrow();
  });
  it("retains zero-size source extents and accounts objects in the cumulative metadata limit", () => {
    const files = seeded(object => { object.anchor = { kind: "one-cell", from: { row: 0, column: 0, offsetXPt: 0, offsetYPt: 0 }, widthPt: 0, heightPt: 0 }; });
    expect(openDocument(files, { kind: "workbook" }).snapshot().metadata.workbook.objects?.[0]?.anchor).toMatchObject({ widthPt: 0, heightPt: 0 });
    expect(() => openDocument(files, { limits: { maxNodes: 4 } })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
  it("binds safe object hyperlinks to their owned source and validates paragraph end styles", () => {
    const files = seeded((_, metadata) => { metadata.workbook.objectHyperlinks = [{ nodeId: "run", location: "'Sheet 1'!A1", tooltip: "说明" }, { nodeId: "image", url: "https://example.test/source" }]; });
    expect(openDocument(files, { kind: "workbook" }).snapshot().metadata.workbook.objectHyperlinks).toHaveLength(2);
    for (const link of [{ nodeId: "content", url: "https://example.test" }, { nodeId: "missing", location: "A1" }, { nodeId: "run", url: "javascript:run()" }, { nodeId: "run", location: "SUM(A1:A3)" }]) {
      expect(() => openDocument(seeded((_, metadata) => { metadata.workbook.objectHyperlinks = [link]; }))).toThrow();
    }
    expect(() => openDocument(seeded((_, metadata) => { metadata.workbook.objectHyperlinks = [{ nodeId: "run", location: "A1" }, { nodeId: "run", location: "A2" }]; }))).toThrow();
    expect(() => openDocument(seeded(undefined, source => source.replace('<p data-node-id="paragraph">', '<p data-node-id="paragraph" data-office-end-style="missing">')))).toThrow();
  });
});
