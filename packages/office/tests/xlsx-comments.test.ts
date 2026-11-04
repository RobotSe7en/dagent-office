import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, DEFAULT_DOCUMENT_READ_LIMITS, openDocument } from "@dagent-office/core";
import { openOpc, openOpcFiles } from "../src/ooxml/package.js";
import type { OpcPackage } from "../src/ooxml/package.js";
import { readWorkbook } from "../src/xlsx/workbook.js";
import { readThreadedComments } from "../src/xlsx/threaded-comments.js";
import { WorkbookContent } from "../src/xlsx/content.js";
import { mapThreadedComments } from "../src/xlsx/comment-mapping.js";

const encoder = new TextEncoder(), decoder = new TextDecoder(), threadPart = "xl/threadedComments/threadedComment1.xml", personPart = "xl/persons/person.xml";
const rootId = "{33333333-3333-4333-8333-333333333333}", replyId = "{44444444-4444-4444-8444-444444444444}", otherId = "{55555555-5555-4555-8555-555555555555}";
const source = (file = "10-modern-threads") => openOpc(readFileSync(`examples/xlsx/${file}.xlsx`), {});
function variant(edits: Record<string, (source: string) => string>) {
  const files = new Map(source().files); for (const [path, edit] of Object.entries(edits)) files.set(path, encoder.encode(edit(decoder.decode(files.get(path)!))));
  return openOpcFiles(files, {});
}
function parsed(opc: OpcPackage, maxItems = 1000) { const workbook = readWorkbook(opc); return readThreadedComments(opc, workbook.part, workbook.sheets, maxItems); }
function mapped(opc = source()) {
  const native = parsed(opc), content = new WorkbookContent(DEFAULT_DOCUMENT_READ_LIMITS), mapped = mapThreadedComments(native, native.threads, content, "threads");
  const files = new Map(createWorkbook({ documentId: "threads" }));
  const manifest = JSON.parse(decoder.decode(files.get("manifest.json")!)), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!));
  manifest.content = "content.html"; metadata.workbook.stories = content.stories; metadata.workbook.comments = mapped.comments; metadata.workbook.commentAuthors = mapped.authors;
  for (const [path, value] of [["manifest.json", manifest], ["metadata.json", metadata], ["annotations.json", mapped.annotations]] as const) files.set(path, encoder.encode(JSON.stringify(value)));
  files.set("content.html", encoder.encode(content.html()!)); return { files, native, content, mapped };
}
describe("native spreadsheet threaded comments", () => {
  it("preserves author identity, Unicode, replies, timestamps and resolved state in the same workbook package", () => {
    const { files, native } = mapped(); expect(native.threads).toHaveLength(2); expect(native.people).toHaveLength(2);
    const model = openDocument(files, { kind: "workbook" }), annotations = model.snapshot().annotations;
    expect(annotations[0]).toMatchObject({ status: "resolved", sourceId: rootId, date: "2026-11-04T08:00:00Z", author: { name: "预算负责人" }, anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 1, column: 1 }, to: { row: 1, column: 1 } }] } });
    expect(model.getNode(annotations[0]!.bodyNodeId!)?.text).toBe("请复核预算😀禰󠄀");
    expect(model.getNode(annotations[0]!.replies![0]!.bodyNodeId!)?.text).toBe("已复核并补充。");
    expect(model.snapshot().metadata.workbook.commentAuthors?.[0]).toMatchObject({ userId: "opaque-owner", providerId: "None" });
    expect(openDocument(model.serialize()).snapshot()).toEqual(model.snapshot());
    // The pinned external producer removes modern threads entirely. Reading
    // its output must report the actual empty source rather than fabricate it.
    expect(parsed(source("producer/10-modern-threads"))).toEqual({ people: [], threads: [] });
  });
  it("retains nested reply parents and accepts replies preceding roots or omitting their inherited cell", () => {
    const opc = variant({ [threadPart]: s => {
      const nodes = s.match(/<threadedComment\s[\s\S]*?<\/threadedComment>/g)!;
      const nested = nodes[2]!.replace('ref="C3"', `parentId="${replyId}"`).replace('done="0"', 'done="1"');
      return s.slice(0, s.indexOf(nodes[0]!)) + nested + nodes[1]!.replace('ref="B2" ', '') + nodes[0]! + '</ThreadedComments>';
    } });
    const { files, native } = mapped(opc); expect(native.threads).toHaveLength(1); expect(native.threads[0]!.replies).toHaveLength(2);
    const model = openDocument(files, { kind: "workbook" });
    expect(model.snapshot().metadata.workbook.comments?.[0]).toMatchObject({ replies: [{ replyId: `xlsx-reply-${otherId}`, parentReplyId: `xlsx-reply-${replyId}` }, { replyId: `xlsx-reply-${replyId}` }] });
  });
  it("preserves optional empty text and escapes source markup and native UTF-16 strings exactly once", () => {
    const { files } = mapped(variant({ [threadPart]: s => s.replace('<text>待确认。</text>', '<text xml:space="preserve"> &lt;script&gt;_x005F_x0041__x000D_\n😀 </text>').replace('<text>已复核并补充。</text>', '') }));
    const model = openDocument(files, { kind: "workbook" }), annotations = model.snapshot().annotations;
    expect(model.getNode(annotations[1]!.bodyNodeId!)?.text).toBe(" <script>_x0041_\r\n😀 ");
    expect(model.getNode(annotations[0]!.replies![0]!.bodyNodeId!)?.text).toBe("");
    expect(openDocument(model.serialize()).snapshot()).toEqual(model.snapshot());
  });
  it.each([
    ["unknown author", (s: string) => s.replace('personId="{11111111-1111-4111-8111-111111111111}"', `personId="${otherId}"`)],
    ["duplicate identity", (s: string) => s.replace(`id="${otherId}"`, `id="${rootId}"`)],
    ["bad GUID", (s: string) => s.replace(`id="${otherId}"`, 'id="x"')],
    ["missing parent", (s: string) => s.replace(`parentId="${rootId}"`, 'parentId="{66666666-6666-4666-8666-666666666666}"')],
    ["self cycle", (s: string) => s.replace(`parentId="${rootId}"`, `parentId="${replyId}"`)],
    ["cross-cell reply", (s: string) => s.replace('ref="B2" dT="2026-11-04T08:01:00Z"', 'ref="B3" dT="2026-11-04T08:01:00Z"')],
    ["duplicate root cell", (s: string) => s.replace('ref="C3"', 'ref="B2"')],
    ["invalid boolean", (s: string) => s.replace('done="1"', 'done="yes"')],
    ["normalized invalid calendar", (s: string) => s.replace('2026-11-04T08:00:00Z', '2026-02-30T08:00:00Z')],
    ["invalid zone", (s: string) => s.replace('2026-11-04T08:00:00Z', '2026-11-04T08:00:00+14:01')],
  ])("rejects %s instead of returning a partial comment set", (_name, edit) => {
    expect(() => parsed(variant({ [threadPart]: edit }))).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  });
  it("rejects cyclic fixture, duplicated people, undeclared attributes/extensions and exceeds cumulative records before mapping", () => {
    expect(() => parsed(source("invalid/thread-parent-cycle"))).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    expect(() => parsed(variant({ [personPart]: s => s.replace('id="{22222222-2222-4222-8222-222222222222}"', 'id="{11111111-1111-4111-8111-111111111111}"') }))).toThrow();
    expect(() => parsed(variant({ [threadPart]: s => s.replace('done="1"', 'done="1" surprise="1"') }))).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    expect(() => parsed(variant({ [threadPart]: s => s.replace('</ThreadedComments>', '<extLst><ext uri="unknown"/></extLst></ThreadedComments>') }))).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    expect(() => parsed(source(), 4)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    const content = new WorkbookContent({ ...DEFAULT_DOCUMENT_READ_LIMITS, maxFileBytes: 300 });
    expect(() => content.comment("sheet-1", "😀".repeat(200))).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
  it("bounds repeated long author names before assembling the complete annotation JSON", () => {
    const native = parsed(variant({ [personPart]: s => s.replace('displayName="预算负责人"', `displayName="${"长".repeat(1000)}"`) }));
    const content = new WorkbookContent({ ...DEFAULT_DOCUMENT_READ_LIMITS, maxFileBytes: 5000 });
    expect(() => mapThreadedComments(native, native.threads, content, "threads")).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded", message: "Workbook comment metadata exceeds its source byte budget." }));
  });
  it("validates native author/reply bindings again when opening an edited source package", () => {
    for (const edit of [
      (m: ReturnType<typeof mapped>["mapped"]) => { m.authors[0]!.name = "changed"; },
      (m: ReturnType<typeof mapped>["mapped"]) => { m.comments[0]!.annotationId = "missing"; },
      (m: ReturnType<typeof mapped>["mapped"]) => { m.comments[0]!.recoveryStatus = "resolved"; },
      (m: ReturnType<typeof mapped>["mapped"]) => { const c = m.comments[0]!; if (c.format === "modern") c.replies[0]!.parentReplyId = c.replies[0]!.replyId; },
      (m: ReturnType<typeof mapped>["mapped"]) => { const c = m.comments[0]!; if (c.format === "modern") c.replies = []; },
    ]) {
      const data = mapped(); edit(data.mapped); const metadata = JSON.parse(decoder.decode(data.files.get("metadata.json")!));
      metadata.workbook.commentAuthors = data.mapped.authors; metadata.workbook.comments = data.mapped.comments;
      data.files.set("metadata.json", encoder.encode(JSON.stringify(metadata))); expect(() => openDocument(data.files)).toThrow();
    }
  });
});
