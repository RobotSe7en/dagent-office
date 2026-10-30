import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@dagent-office/core";
import { importDocx, readDocumentZip, writeDocumentZip } from "@dagent-office/office";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const directory = new URL("../../../examples/docx/", import.meta.url);
const cases = JSON.parse(readFileSync(new URL("expected.json", directory), "utf8")) as { file: string }[];

for (const file of cases.flatMap((item) => [item.file, "producer/" + item.file])) {
  test("edits imported DOCX, saves/reopens and restores history: " + file, () => {
    const bytes = new Uint8Array(readFileSync(new URL(file, directory)));
    const model = openDocument(importDocx(bytes, { documentId: "editing" }));
    const original = model.snapshot();
    const session = createDocumentSession(model);
    const ids = [...original.html.matchAll(/data-node-id="([^"]+)"/g)].map((match) => match[1]!);
    const target = ids.map((id) => model.getNode(id)!).find((node) => node.tagName === "span" && node.text.trim()
      && node.childNodeIds.length === 0 && node.attributes["data-word-tab"] === undefined)!;
    expect(target).toBeDefined();
    const candidate = session.prepare({ protocolVersion: 1, documentId: "editing", baseRevision: 0, operations: [
      { op: "replaceText", nodeId: target.nodeId, from: 0, to: target.text.length, expectedText: target.text, text: "修改" + target.text },
      { op: "setStyle", nodeId: target.nodeId, property: "color", value: "#123456" },
    ] });
    expect(model.snapshot()).toEqual(original);
    session.commit(candidate);
    expect(session.getNode(target.nodeId)!.text).toBe("修改" + target.text);
    expect(session.snapshot().styles).toEqual(original.styles);
    expect(session.snapshot().metadata.word!.styles).toEqual(original.metadata.word!.styles);
    expect(session.snapshot().metadata.word!.revisions).toEqual(original.metadata.word!.revisions);
    const reopened = openDocument(readDocumentZip(writeDocumentZip(session.serialize())));
    expect(reopened.snapshot()).toEqual(session.snapshot());
    for (const asset of original.manifest.assets) expect(reopened.serialize().get(asset.path)).toEqual(model.serialize().get(asset.path));
    session.undo();
    expect(session.snapshot().html).toBe(original.html);
    expect(session.snapshot().metadata).toEqual(original.metadata);
    session.redo();
    expect(session.snapshot().html).toBe(candidate.preview.html);
  });
}

test.each(["宋体", "新宋体", "SimSun", "NSimSun", "仿宋", "仿宋_GB2312", "华文仿宋", "FangSong", "FangSong_GB2312", "STFangsong",
  "楷体", "楷体_GB2312", "华文楷体", "KaiTi", "KaiTi_GB2312", "STKaiti", "FandolSong", "FandolFang", "FandolKai", "LXGW WenKai"])(
  "preserves Chinese font %s, Latin roles and variation sequences across editing/history/reopen", (family) => {
    const source = unzipSync(new Uint8Array(readFileSync(new URL("09-chinese-fonts.docx", directory))));
    for (const path of ["word/document.xml", "word/styles.xml"]) {
      source[path] = strToU8(strFromU8(source[path]!).replace(/w:eastAsia="[^"]*"/g, 'w:eastAsia="' + family + '"'));
    }
    const model = openDocument(importDocx(zipSync(source), { documentId: "fonts" }));
    const session = createDocumentSession(model);
    const original = model.snapshot();
    const css = original.styles.map((style) => style.css).join("\n") + original.html;
    expect(css).toContain(family);
    expect(css).toContain("--word-font-ascii");
    expect(css).toContain("--word-font-eastAsia");
    const ids = [...original.html.matchAll(/data-node-id="([^"]+)"/g)].map((match) => match[1]!);
    const target = ids.map((id) => model.getNode(id)!).find((node) => node.attributes["data-word-run"] === "true" && node.text.includes("ABC"))
      ?? ids.map((id) => model.getNode(id)!).find((node) => node.attributes["data-word-run"] === "true" && node.text.trim())!;
    const candidate = session.prepare({ protocolVersion: 1, documentId: "fonts", baseRevision: 0, operations: [
      { op: "replaceText", nodeId: target.nodeId, from: 0, to: 0, expectedText: "", text: "中文ABC𠀀\u{e0100}" },
    ] });
    session.commit(candidate);
    const edited = session.getNode(target.nodeId)!;
    expect(edited.text.startsWith("中文ABC𠀀\u{e0100}")).toBe(true);
    expect(edited.html).toContain('data-word-script="eastAsia"');
    expect(edited.html).toContain("--word-font-ascii");
    expect(session.snapshot().styles).toEqual(original.styles);
    expect(openDocument(session.serialize()).snapshot()).toEqual(session.snapshot());
    session.undo();
    expect(session.snapshot().html).toBe(original.html);
  },
);

for (const file of ["08-comments-revisions.docx", "producer/08-comments-revisions.docx"]) {
  test("maps imported rich comments through editing, history and ZIP reopen: " + file, () => {
    const model = openDocument(importDocx(new Uint8Array(readFileSync(new URL(file, directory))), { documentId: "anchors" }));
    const original = model.snapshot().annotations[0]!;
    if (original.anchor.kind !== "text-set" && original.anchor.kind !== "text") throw new Error("Expected a native text comment");
    const ranges = original.anchor.kind === "text-set" ? original.anchor.ranges : [original.anchor];
    const first = ranges[0]!;
    const second = ranges[1] ?? first;
    const session = createDocumentSession(model);
    const transaction = (operations: import("@dagent-office/core").DocumentOperation[]) => session.prepare({
      protocolVersion: 1, documentId: "anchors", baseRevision: session.snapshot().manifest.revision, operations,
    });
    const candidate = transaction([
      { op: "replaceText", nodeId: first!.nodeId, from: 0, to: 0, expectedText: "", text: "前言😀" },
      { op: "replaceText", nodeId: second.nodeId, from: second === first ? 5 : 1, to: second === first ? 6 : 2, expectedText: second.quote.slice(1, 2), text: "修改" },
    ]);
    session.commit(candidate);
    const expectedRanges = ranges.map((range, index) => {
      const from = range.from + (index === 0 ? 4 : 0);
      const to = range.to + (index === 0 ? 4 : 0) + (range === second ? 1 : 0);
      const quote = range === second ? range.quote.slice(0, 1) + "修改" + range.quote.slice(2) : range.quote;
      return { nodeId: range.nodeId, from, to, quote };
    });
    const expectedAnchor = original.anchor.kind === "text-set" ? { kind: "text-set", ranges: expectedRanges }
      : { kind: "text", ...expectedRanges[0] };
    expect(session.getAnnotations()[0]).toEqual({ ...original, baseRevision: 1, anchor: expectedAnchor });
    expect(openDocument(readDocumentZip(writeDocumentZip(session.serialize()))).snapshot()).toEqual(session.snapshot());
    session.undo();
    expect(session.getAnnotations()[0]).toEqual({ ...original, baseRevision: 2 });
    session.redo();
    expect(session.getAnnotations()[0]).toEqual({ ...candidate.preview.annotations[0], baseRevision: 3 });
    session.commit(transaction([{ op: "removeNode", nodeId: second!.nodeId }]));
    expect(session.getAnnotations()[0]).toMatchObject({ status: "needs_reanchor", baseRevision: 3, anchor: candidate.preview.annotations[0]!.anchor });
    expect(session.getAnnotations()[0]!.bodyNodeId).toBe(original.bodyNodeId);
    expect(openDocument(readDocumentZip(writeDocumentZip(session.serialize()))).snapshot()).toEqual(session.snapshot());
    session.dispose();
  });
}
