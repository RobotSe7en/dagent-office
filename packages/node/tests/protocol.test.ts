import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createPresentation, openDocument } from "@dagent-office/core";
import type { DocumentModel } from "@dagent-office/core";
import { writeDocumentZip, exportPptx } from "@dagent-office/office";
import { createDocumentConverter, executeConversionRequest, parseConversionRequest, serveConversionProtocol } from "@dagent-office/node";
import type { ConversionRequestV1, DocumentConverter, RenderedOutput } from "@dagent-office/node";

const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path,
  new Uint8Array(readFileSync(new URL("../../../examples/report/" + path, import.meta.url))),
]));
function request(taskId = "任务😀"): ConversionRequestV1 {
  return { protocolVersion: 1, taskId, format: "pdf", input: { format: "document-package-zip", documentId: "report-demo", bytesBase64: Buffer.from(writeDocumentZip(files)).toString("base64") } };
}
function fake(bytes = 8): DocumentConverter & { calls: string[] } {
  const calls: string[] = [];
  const render = async (model: DocumentModel): Promise<RenderedOutput> => {
    const snapshot = model.snapshot(); calls.push(snapshot.manifest.documentId);
    return { bytes: new Uint8Array(bytes), mimeType: "application/pdf", engine: { name: "chromium", version: "fixture" },
      report: { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, layout: "pages", pages: [], nodePages: {}, fields: [], diagnostics: [], fonts: [] } };
  };
  return { calls, renderPdf: render, renderPage: render, async close() {} };
}

test("conversion accepts only versioned authorized bytes and preserves asserted package identity", async () => {
  const converter = fake(), input = request(), result = await executeConversionRequest(converter, input);
  expect(result).toMatchObject({ protocolVersion: 1, taskId: "任务😀", ok: true, output: { mimeType: "application/pdf", report: { documentId: "report-demo", revision: 0 } } });
  expect(converter.calls).toEqual(["report-demo"]);
  expect(await executeConversionRequest(converter, { ...input, input: { ...input.input, documentId: "another-private-document" } })).toMatchObject({ ok: false, error: { code: "InvalidDocument" } });
  expect(converter.calls).toHaveLength(1);
});
test.each([
  [{ ...request(), protocolVersion: 2 }, "UnsupportedVersion"],
  [{ ...request(), format: "unknown" }, "UnsupportedFormat"],
  [{ ...request(), input: { format: "docx", bytesBase64: "AA==" } }, "InvalidRequest"],
  [{ ...request(), input: { ...request().input, path: "/private/file.docx" } }, "InvalidRequest"],
  [{ ...request(), input: { ...request().input, url: "https://example.com/private.docx" } }, "InvalidRequest"],
  [{ ...request(), options: { browserExecutable: "/tmp/executable" } }, "InvalidRequest"],
  [{ ...request(), options: { pageIndex: 0 } }, "InvalidRequest"],
  [{ ...request(), options: { limits: { maxPages: 0 } } }, "InvalidRequest"],
  [{ ...request(), extra: true }, "InvalidRequest"],
])("rejects an unsupported conversion request without invoking the renderer", async (input, code) => {
  const converter = fake();
  expect(await executeConversionRequest(converter, input)).toMatchObject({ taskId: "任务😀", ok: false, error: { code } });
  expect(converter.calls).toEqual([]);
});
test.each(["Z===", "Zh==", "AA A", "AAAA=", "AAéA"])("rejects noncanonical base64 %s", async (bytesBase64) => {
  expect(await executeConversionRequest(fake(), { ...request(), input: { ...request().input, bytesBase64 } })).toMatchObject({ ok: false, error: { code: "InvalidRequest" } });
});
test("conversion byte budgets are checked before decoding or rendering", async () => {
  const converter = fake();
  expect(await executeConversionRequest(converter, request(), { maxInputBytes: 1 })).toMatchObject({ ok: false, error: { code: "ResourceLimitExceeded" } });
  expect(converter.calls).toEqual([]);
});
test("NDJSON handles split UTF-8 and serializes requests behind output backpressure", async () => {
  const converter = fake(), data = new TextEncoder().encode(JSON.stringify(request("任务😀")) + "\n" + JSON.stringify(request("next")) + "\n");
  const output: string[] = [];
  async function* chunks(): AsyncGenerator<Uint8Array> { for (const byte of data) yield new Uint8Array([byte]); }
  await serveConversionProtocol(chunks(), async (line) => {
    expect(converter.calls).toHaveLength(output.length + 1); await Promise.resolve(); output.push(line);
  }, converter);
  expect(output.map((line) => JSON.parse(line).taskId)).toEqual(["任务😀", "next"]);
});
test("invalid JSON is a recoverable frame and oversized frames stop before the renderer", async () => {
  const converter = fake(), output: string[] = [], value = JSON.stringify(request());
  async function* input(): AsyncGenerator<Uint8Array> { yield new TextEncoder().encode("{bad}\n" + value + "\n"); }
  await serveConversionProtocol(input(), async (line) => { output.push(line); }, converter);
  expect(JSON.parse(output[0]!)).toMatchObject({ ok: false, taskId: null, error: { code: "InvalidRequest" } });
  expect(JSON.parse(output[1]!)).toMatchObject({ ok: true, taskId: "任务😀" });
  const rejected = fake();
  await expect(serveConversionProtocol(input(), async () => {}, rejected, { maxMessageBytes: 10 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  expect(rejected.calls).toEqual([]);
});
test("output frames are bounded and malformed UTF-8 is rejected", async () => {
  const output: string[] = [];
  async function* input(): AsyncGenerator<Uint8Array> { yield new TextEncoder().encode(JSON.stringify(request()) + "\n"); }
  await serveConversionProtocol(input(), async (line) => { output.push(line); }, fake(1024), { maxOutputMessageBytes: 512 });
  expect(JSON.parse(output[0]!)).toMatchObject({ ok: false, error: { code: "ResourceLimitExceeded" } });
  async function* invalid(): AsyncGenerator<Uint8Array> { yield new Uint8Array([0xff, 10]); }
  await expect(serveConversionProtocol(invalid(), async () => {}, fake())).rejects.toThrow();
  async function* malformed(): AsyncGenerator<Uint8Array> { yield new TextEncoder().encode("{bad}\n"); }
  const rejected: string[] = [];
  await expect(serveConversionProtocol(malformed(), async (line) => { rejected.push(line); }, fake(), { maxOutputMessageBytes: 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  expect(rejected).toEqual([]);
});
test("converter configuration, cancellation and quota failures do not require browser startup", async () => {
  expect(() => createDocumentConverter({ maxConcurrency: 0 })).toThrow();
  const model = openDocument(files), converter = createDocumentConverter({ maxInputBytes: 1 });
  await expect(converter.renderPdf(model)).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  const signal = new AbortController(); signal.abort();
  await expect(converter.renderPdf(model, { signal: signal.signal })).rejects.toMatchObject({ code: "Cancelled" });
  await expect(converter.renderPdf(model, { limits: { maxPages: 257 } })).rejects.toMatchObject({ code: "InvalidRequest" });
  await converter.close(); await converter.close();
  await expect(converter.renderPdf(model)).rejects.toMatchObject({ code: "InvalidOperation" });
  expect(() => parseConversionRequest({})).toThrow();
});

test("native PPTX bytes and PPT packages use the same bounded conversion boundary",async()=>{
  const files=createPresentation({documentId:"private-ppt"}),native=exportPptx(openDocument(files)).bytes;
  for(const[format,bytes]of[["pptx",native],["document-package-zip",writeDocumentZip(files)]]as const){
    const converter=fake(),result=await executeConversionRequest(converter,{protocolVersion:1,taskId:"ppt-conversion",format:"pdf",input:{format,documentId:"private-ppt",bytesBase64:Buffer.from(bytes).toString("base64")}});
    expect(result).toMatchObject({ok:true,output:{report:{documentId:"private-ppt"}}});expect(converter.calls).toEqual(["private-ppt"]);
  }
  expect(()=>parseConversionRequest({protocolVersion:1,taskId:"missing-id",format:"pdf",input:{format:"pptx",bytesBase64:"AA=="}})).toThrow();
});
