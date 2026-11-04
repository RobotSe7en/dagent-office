import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "../src/index.js";
import type { Annotation, WorkbookDocumentMetadata, WorkbookManifest } from "../src/index.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
function set(files: Map<string, Uint8Array>, path: string, value: unknown) { files.set(path, encoder.encode(JSON.stringify(value))); }
function json<T>(files: Map<string, Uint8Array>, path: string): T { return JSON.parse(decoder.decode(files.get(path)!)); }
function seeded() {
  const files = new Map(createWorkbook({ documentId: "comments" })), manifest = json<WorkbookManifest>(files, "manifest.json"), metadata = json<WorkbookDocumentMetadata>(files, "metadata.json");
  manifest.content = "content.html"; metadata.workbook.stories = [{ storyId: "story-1", rootNodeId: "body-1", kind: "comment", sheetId: "sheet-1" }, { storyId: "story-2", rootNodeId: "body-2", kind: "comment", sheetId: "sheet-1" }];
  const annotation: Annotation = { annotationId: "note", documentId: "comments", baseRevision: 0, status: "resolved", bodyNodeId: "body-1", author: { name: "作者" },
    anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 1, column: 1 }, to: { row: 1, column: 1 } }, { from: { row: 8, column: 3 }, to: { row: 10, column: 4 } }] },
    replies: [{ replyId: "reply", bodyNodeId: "body-2", author: { name: "复核" } }] };
  set(files, "manifest.json", manifest); set(files, "metadata.json", metadata); set(files, "annotations.json", [annotation]);
  files.set("content.html", encoder.encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body><section data-node-id="content" data-office-kind="workbook-content"><section data-node-id="body-1" data-office-kind="comment" data-office-story="comment"><p data-node-id="p1"><span data-node-id="r1" class="cell-default">预算😀禰󠄀</span></p></section><section data-node-id="body-2" data-office-kind="comment" data-office-story="comment"><p data-node-id="p2">已复核。</p></section></section></body></html>'));
  return files;
}
describe("workbook cell annotations and bounded HTML stories", () => {
  it("opens, queries and roundtrips exact rich bodies and disjoint cell regions without synthesizing cell chunks", () => {
    const files = seeded(), model = openDocument(files, { kind: "workbook" });
    expect(model.snapshot().manifest.chunks).toEqual([]); expect(model.getCell("sheet-1", 1, 1)).toBeUndefined();
    expect(model.getNode("body-1")?.text).toBe("预算😀禰󠄀"); expect(model.getNode("body-2")?.text).toBe("已复核。");
    expect(model.snapshot().content?.path).toBe("content.html"); expect(model.snapshot().annotations[0]?.anchor).toMatchObject({ kind: "cells", ranges: [{ from: { row: 1, column: 1 } }, { from: { row: 8, column: 3 } }] });
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(model.snapshot());
    files.get("content.html")!.fill(0); model.serialize().get("content.html")!.fill(0); expect(model.getNode("body-1")?.text).toBe("预算😀禰󠄀");
    expect(Object.isFrozen(model.snapshot().annotations[0]?.anchor)).toBe(true);
  });
  it("supports instruction-only cell annotations and preserves explicitly invalidated prior targets", () => {
    const files = new Map(createWorkbook({ documentId: "instructions" }));
    const annotations: Annotation[] = [{ annotationId: "instruction", documentId: "instructions", baseRevision: 0, status: "open", instruction: "复核此处", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 9, column: 2 }, to: { row: 15, column: 4 } }] } }];
    set(files, "annotations.json", annotations); expect(openDocument(files).snapshot().annotations).toEqual(annotations);
    annotations[0]!.status = "needs_reanchor"; if (annotations[0]!.anchor.kind === "cells") annotations[0]!.anchor.sheetId = "deleted-sheet";
    set(files, "annotations.json", annotations); expect(openDocument(files).snapshot().annotations).toEqual(annotations);
  });
  it("retains rich bodies and replies when their worksheet was deleted and the annotation needs recovery", () => {
    const files = seeded(), metadata = json<WorkbookDocumentMetadata>(files, "metadata.json"), annotations = json<Annotation[]>(files, "annotations.json");
    for (const story of metadata.workbook.stories!) story.sheetId = "deleted-sheet";
    const annotation = annotations[0]!; annotation.status = "needs_reanchor";
    if (annotation.anchor.kind === "cells") annotation.anchor.sheetId = "deleted-sheet";
    set(files, "metadata.json", metadata); set(files, "annotations.json", annotations);
    const model = openDocument(files, { kind: "workbook" });
    expect(model.getNode("body-1")?.text).toBe("预算😀禰󠄀"); expect(model.getNode("body-2")?.text).toBe("已复核。");
    expect(openDocument(model.serialize()).snapshot()).toEqual(model.snapshot());
    annotation.status = "open"; set(files, "annotations.json", annotations); expect(() => openDocument(files)).toThrow();
  });
  it("rejects overlapping regions, future revisions, missing worksheets and duplicate reply/body ownership", () => {
    const edits: ((annotations: Annotation[]) => void)[] = [
      a => { a[0]!.baseRevision = 1; },
      a => { if (a[0]!.anchor.kind === "cells") a[0]!.anchor.sheetId = "missing"; },
      a => { if (a[0]!.anchor.kind === "cells") a[0]!.anchor.ranges.push(a[0]!.anchor.ranges[0]!); },
      a => { a[0]!.replies![0]!.replyId = a[0]!.annotationId; },
      a => { a[0]!.replies![0]!.bodyNodeId = a[0]!.bodyNodeId!; },
      a => { a[0]!.bodyNodeId = "r1"; },
    ];
    for (const edit of edits) { const files = seeded(), annotations = json<Annotation[]>(files, "annotations.json"); edit(annotations); set(files, "annotations.json", annotations); expect(() => openDocument(files)).toThrow(); }
  });
  it("enforces the shared scriptless boundary and every content file role/style/story reference", () => {
    for (const edit of [
      (s: string) => s.replace('<section data-node-id="body-1"', '<br><section data-node-id="body-1"'),
      (s: string) => s.replace('<section data-node-id="body-1"', '<br data-node-id="outside"><section data-node-id="body-1"'),
      (s: string) => s.replace('<p data-node-id="p1">', '<p data-node-id="p1" data-office-story="comment">'),
      (s: string) => s.replace('<p data-node-id="p1">', '<p data-node-id="p1" onclick="alert(1)">'),
      (s: string) => s.replace('class="cell-default"', 'class="missing"'),
      (s: string) => s.replace('data-node-id="p1"', 'data-node-id="p2"'),
      (s: string) => s.replace('data-office-story="comment"', 'data-office-story="unhandled"'),
    ]) { const files = seeded(); files.set("content.html", encoder.encode(edit(decoder.decode(files.get("content.html")!)))); expect(() => openDocument(files)).toThrow(); }
    const orphan = seeded(); set(orphan, "annotations.json", []); expect(() => openDocument(orphan)).toThrow();
    const roles = seeded(), manifest = json<WorkbookManifest>(roles, "manifest.json"); manifest.content = "styles.css"; set(roles, "manifest.json", manifest); expect(() => openDocument(roles)).toThrow();
    const missing = seeded(); missing.delete("content.html"); expect(() => openDocument(missing)).toThrow();
  });
  it("rejects cell anchors in existing Word documents even when marked for recovery", () => {
    const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(path => [path, new Uint8Array(readFileSync(`examples/report/${path}`))]));
    const annotations = json<Annotation[]>(files, "annotations.json"); annotations[0]!.status = "needs_reanchor";
    annotations[0]!.anchor = { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] }; set(files, "annotations.json", annotations);
    expect(() => openDocument(files)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  });
});
