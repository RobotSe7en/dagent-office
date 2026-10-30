import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@dagent-office/core";
import type { Anchor, Annotation } from "@dagent-office/core";
import { exportDocx, importDocx } from "@dagent-office/office";
import { strFromU8, unzipSync } from "fflate";
const root = new URL("../../../examples/report/", import.meta.url);
const report = () => openDocument(new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((name) => [name, new Uint8Array(readFileSync(new URL(name, root)))])));
function sourceWith(anchor: Anchor, status: Annotation["status"] = "open") {
  const model = report(), files = new Map(model.serialize()), snapshot = model.snapshot();
  const annotation: Annotation = { annotationId: "new", documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, instruction: "检查", status, anchor };
  files.set("annotations.json", new TextEncoder().encode(JSON.stringify([annotation]))); return openDocument(files);
}
function quote(exported: ReturnType<typeof exportDocx>): string {
  const doc = openDocument(importDocx(exported.bytes, { documentId: "reopen" })), anchor = doc.snapshot().annotations[0]!.anchor;
  return anchor.kind === "text" ? anchor.quote : anchor.kind === "text-set" ? anchor.ranges.map((range) => range.quote).join("") : anchor.kind === "node" ? doc.getNode(anchor.nodeId)!.text : "";
}
test("node comments, exact UTF-16 text and adjacent native paragraphs bind to the intended content", () => {
  expect(quote(exportDocx(sourceWith({ kind: "node", nodeId: "n-summary" })))).toBe("项目计划于第三季度完成。");
  expect(quote(exportDocx(sourceWith({ kind: "text", nodeId: "n-summary", from: 2, to: 6, quote: "计划于第" })))).toBe("计划于第");
});
test("region and stale comments are explicitly reported without stopping body export", () => {
  const region = sourceWith({ kind: "region", surfaceId: "root", rect: { x: 1, y: 1, width: 2, height: 2 }, nodeIds: [], layoutKey: "old" });
  const exported = exportDocx(region);
  expect(exported.report.unexportedAnnotations).toEqual([{ annotationId: "new", reason: "region" }]);
  expect(openDocument(importDocx(exported.bytes, { documentId: "body" })).snapshot().annotations).toEqual([]);
  const stale = exportDocx(sourceWith({ kind: "node", nodeId: "n-summary" }, "needs_reanchor"));
  expect(stale.report.unexportedAnnotations).toEqual([{ annotationId: "new", reason: "needs_reanchor" }]);
  expect(region.snapshot().annotations).toHaveLength(1);
});
test("disjoint ranges are reported instead of guessing a larger native comment range", () => {
  const exported = exportDocx(sourceWith({ kind: "text-set", ranges: [{ nodeId: "n-summary", from: 0, to: 2, quote: "项目" }, { nodeId: "n-summary", from: 6, to: 8, quote: "三季" }] }));
  expect(exported.report.unexportedAnnotations).toEqual([{ annotationId: "new", reason: "disjoint" }]);
});
test("HTML lists and native Word lists can coexist after a structural edit", () => {
  const input = new Uint8Array(readFileSync(new URL("../../../examples/docx/01-text-numbering.docx", import.meta.url))), session = createDocumentSession(openDocument(importDocx(input, { documentId: "lists" })));
  const paragraph = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map((match) => session.getNode(match[1]!)!).find((node) => node.tagName === "p")!;
  session.commit(session.prepare({ protocolVersion: 1, documentId: "lists", baseRevision: 0, operations: [{ op: "insertFragment", nodeId: paragraph.nodeId, position: "after", html: '<ul ><li >NEW <strong >BOLD</strong><ol  start="3"><li >NESTED</li></ol></li></ul>' }] }));
  const exported = exportDocx(session), parts = unzipSync(exported.bytes);
  expect(strFromU8(parts["word/document.xml"]!)).not.toContain('w:numId="html');
  const reopened = openDocument(importDocx(exported.bytes, { documentId: "lists-reopen" }));
  expect(reopened.snapshot().html).toContain("NESTED"); expect(reopened.snapshot().metadata.word!.numbering).toHaveLength(session.snapshot().metadata.word!.numbering.length + 2);
  session.dispose();
});
test("point comments at paragraph offsets and whole-table comments keep exact emission sites", () => {
  const point = exportDocx(sourceWith({ kind: "text", nodeId: "n-summary", from: 3, to: 3, quote: "" }));
  expect(point.report.unexportedAnnotations).toEqual([]); expect(quote(point)).toBe("");
  const table = exportDocx(sourceWith({ kind: "node", nodeId: "n-deliverables" }));
  expect(table.report.unexportedAnnotations).toEqual([]); expect(quote(table)).toBe("交付项状态业务系统实施中");
});
test.each(["word-spacing:2pt", "font-weight:500", "padding-left:4pt", "max-width:10pt"])("unrepresentable body formatting fails without returning partial content: %s", (declaration) => {
  const source = report(), files = new Map(source.serialize());
  files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")) + "\np{" + declaration + "}"));
  const model = openDocument(files), before = model.serialize();
  expect(() => exportDocx(model)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  expect(model.serialize()).toEqual(before);
});
test("collapsible HTML whitespace is explicit instead of silently changing native text", () => {
  const source = report(), files = new Map(source.serialize());
  files.set("content.html", new TextEncoder().encode(new TextDecoder().decode(files.get("content.html")).replace("项目计划于第三季度完成。", " 项目  计划。 ")));
  files.set("annotations.json", new TextEncoder().encode("[]"));
  const model = openDocument(files);
  expect(() => exportDocx(model)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature", message: expect.stringContaining("whitespace") }));
});
