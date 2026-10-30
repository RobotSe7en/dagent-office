import { execFile } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import { expect, test } from "@playwright/test";
import { openDocument } from "@dagent-office/core";
import { importDocx, writeDocumentZip } from "@dagent-office/office";
import { createDocumentConverter } from "@dagent-office/node";
import { runConversionProcess } from "@dagent-office/node/dev";
import { longTable, pagedFiles } from "./fixtures/pagination-files.js";
const exec = promisify(execFile);

test("real PDF contains every source paragraph once, with physical pages and searchable text", async () => {
  const files = pagedFiles(Array.from({ length: 40 }, (_, index) => `<p data-node-id="p-${index}">UNIQUE-${String(index).padStart(2, "0")}</p>`).join(""));
  const model = openDocument(files), before = model.serialize(), converter = createDocumentConverter();
  const directory = await mkdtemp(join(tmpdir(), "dagent-pdf-"));
  try {
    const result = await converter.renderPdf(model); expect(result.report.pages).toHaveLength(4); expect(result.engine.name).toBe("chromium");
    const path = join(directory, "document.pdf"); await writeFile(path, result.bytes);
    const info = await exec("pdfinfo", [path]); expect(info.stdout).toMatch(/Pages:\s+4/); expect(info.stdout).toMatch(/300 x 240 pts/);
    const text = (await exec("pdftotext", ["-layout", path, "-"])).stdout;
    for (let index = 0; index < 40; index++) expect(text.match(new RegExp("UNIQUE-" + String(index).padStart(2, "0"), "g"))).toHaveLength(1);
    expect([...model.serialize()]).toEqual([...before]);
  } finally { await converter.close(); await rm(directory, { recursive: true, force: true }); }
});
test("PDF and PNG repeat table headers and agree on measured page geometry", async () => {
  const model = openDocument(longTable()), converter = createDocumentConverter(), directory = await mkdtemp(join(tmpdir(), "dagent-table-"));
  try {
    const pdf = await converter.renderPdf(model), image = await converter.renderPage(model, { pageIndex: 1 });
    expect(image.report.pages).toEqual(pdf.report.pages); expect(Buffer.from(image.bytes).subarray(1, 4).toString()).toBe("PNG");
    expect(Buffer.from(image.bytes).readUInt32BE(16)).toBe(400); expect(Buffer.from(image.bytes).readUInt32BE(20)).toBe(320);
    const path = join(directory, "table.pdf"); await writeFile(path, pdf.bytes);
    const text = (await exec("pdftotext", ["-layout", path, "-"])).stdout;
    expect(text.match(/REPEATED HEADER/g)).toHaveLength(pdf.report.pages.length);
    for (let index = 0; index < 32; index++) expect(text.match(new RegExp("UNIQUE-" + index + "(?:\\s|$)", "g"))).toHaveLength(1);
    await writeFile(join(directory, "page.png"), image.bytes);
    await expect(Buffer.from(image.bytes)).toMatchSnapshot("repeated-table-page.png");
  } finally { await converter.close(); await rm(directory, { recursive: true, force: true }); }
});
test("converter cancellation, quotas and a later fresh job keep private jobs isolated", async () => {
  const converter = createDocumentConverter({ maxConcurrency: 1 }), model = openDocument(pagedFiles('<p data-node-id="a">PRIVATE-A</p>'));
  const controller = new AbortController();
  try {
    const first = converter.renderPdf(model, { signal: controller.signal });
    await expect(converter.renderPdf(model)).rejects.toMatchObject({ code: "Busy" }); controller.abort();
    await expect(first).rejects.toMatchObject({ code: "Cancelled" });
  } finally { await converter.close(); }
  await expect(converter.renderPdf(model)).rejects.toMatchObject({ code: "InvalidOperation" });
  const second = createDocumentConverter({ maxPixels: 1 });
  try { await expect(second.renderPage(model)).rejects.toMatchObject({ code: "ResourceLimitExceeded" }); }
  finally { await second.close(); }
  const third = createDocumentConverter();
  try { expect((await third.renderPdf(openDocument(pagedFiles('<p data-node-id="b">PRIVATE-B</p>')))).report.documentId).toBe("report-demo"); }
  finally { await third.close(); }
});
test("the independent CLI handles authorized DOCX bytes and package bytes with task correlation", async () => {
  const docx = readFileSync(new URL("../../examples/docx/04-styles-sections.docx", import.meta.url));
  const result = await runConversionProcess({ protocolVersion: 1, taskId: "docx-task", format: "pdf", input: { format: "docx", documentId: "private-docx", bytesBase64: docx.toString("base64") } });
  expect(result).toMatchObject({ protocolVersion: 1, taskId: "docx-task", ok: true, output: { report: { documentId: "private-docx" }, mimeType: "application/pdf" } });
  if (result.ok) {
    const directory = await mkdtemp(join(tmpdir(), "dagent-cli-"));
    try {
      const path = join(directory, "source.pdf"); await writeFile(path, Buffer.from(result.output.bytesBase64, "base64"));
      const text = (await exec("pdftotext", [path, "-"])).stdout; expect(text.match(/报告页眉/g)).toHaveLength(2);
    } finally { await rm(directory, { recursive: true, force: true }); }
  }
  const packageBytes = writeDocumentZip(importDocx(docx, { documentId: "private-package" }));
  const png = await runConversionProcess({ protocolVersion: 1, taskId: "png-task", format: "png", input: { format: "document-package-zip", bytesBase64: Buffer.from(packageBytes).toString("base64") }, options: { pageIndex: 1 } });
  expect(png).toMatchObject({ taskId: "png-task", ok: true, output: { mimeType: "image/png", report: { documentId: "private-package" } } });
});
