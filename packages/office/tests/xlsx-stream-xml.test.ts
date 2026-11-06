import { dailyCi } from '../../../tests/ci-mode.js';
import { readFileSync } from "node:fs";
import { DEFAULT_DOCUMENT_READ_LIMITS } from "@dagent-office/core";
import { expect, it } from "vitest";
import { readDocumentZip } from "../src/document-zip.js";
import { ImportControl } from "../src/import-control.js";
import { streamXml } from "../src/ooxml/xml-stream.js";
import { elements, S, text, xml } from "../src/ooxml/xml.js";
import type { XmlNode } from "../src/ooxml/xml.js";

const limits = { ...DEFAULT_DOCUMENT_READ_LIMITS, maxFileBytes: 32 * 1024 * 1024, maxNodes: 2_000_000 };
const progress = (completed: number, total: number) => ({ phase: "worksheet" as const, completed, total });
const encode = (source: string) => new TextEncoder().encode(source);
it.skipIf(dailyCi)("releases 200,000 cell subtrees while retaining only bounded worksheet headers @scale", async () => {
  const source = readFileSync(new URL("../../../examples/xlsx/15-ten-thousand-rows.xlsx", import.meta.url));
  const bytes = encode(new TextDecoder().decode(readDocumentZip(source, { limits }).get("xl/worksheets/sheet1.xml")!).replaceAll("><", ">\n<"));
  let cells = 0, rows = 0, formulas = 0, last: string | undefined;
  const budget = { nodes: 0 };
  const root = await streamXml(bytes, "sheet.xml", { limits, budget, control: new ImportControl(), maxRetainedNodes: 1000, maxTextCodeUnits: 32767, progress,
    ignoreWhitespace: node => !["t", "v", "f"].includes(node.name),
    complete(node) {
      if (node.name === "c" && node.uri === S) { cells++; formulas += elements(node, "f", S).length; last = text(elements(node, "v", S)[0]); return false; }
      if (node.name === "row" && node.uri === S) { rows++; return false; }
      return true;
    },
  });
  expect({ cells, rows, formulas, last }).toEqual({ cells: 200000, rows: 10000, formulas: 20000, last: "3400152" });
  expect(budget.nodes).toBeGreaterThan(600_000); expect(JSON.stringify(root).length).toBeLessThan(10_000);
});

it("decodes text across UTF boundaries and matches namespace-qualified records", async () => {
  const value = "a".repeat(65500) + "中文😀禰󠄀", source = `<root xmlns="urn:fixture"><t>${value}</t></root>`;
  let captured: XmlNode | undefined;
  const root = await streamXml(encode(source), "unicode.xml", { limits, budget: { nodes: 0 }, control: new ImportControl(), maxRetainedNodes: 10, maxTextCodeUnits: 70000, progress,
    complete(node) { if (node.name === "t") { captured = node; return false; } return true; },
  });
  expect(captured?.uri).toBe("urn:fixture"); expect(text(captured)).toBe(value); expect(root.content).toEqual([]);
  const utf16 = new Uint8Array(Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('<root>中文😀</root>', 'utf16le')]));
  const parsed = await streamXml(utf16, "utf16.xml", { limits, budget: { nodes: 0 }, control: new ImportControl(), maxRetainedNodes: 10, maxTextCodeUnits: 100, progress, complete: () => true });
  expect(text(parsed)).toBe("中文😀");
});

it("keeps cumulative budgets despite releasing records and rejects unsafe XML", async () => {
  const parse = (source: string, maximum = 5) => streamXml(encode(source), "bad.xml", { limits: { ...limits, maxNodes: maximum }, budget: { nodes: 0 }, control: new ImportControl(), maxRetainedNodes: 100, maxTextCodeUnits: 10, progress, complete: node => node.name !== "item" });
  await expect(parse("<root>" + "<item/>".repeat(10) + "</root>")).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  await expect(parse("<!DOCTYPE root [<!ENTITY x SYSTEM 'file:///private'>]><root/>")).rejects.toMatchObject({ code: "InvalidDocument" });
  await expect(parse("<root>12345678901</root>")).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  await expect(parse("<root><other></root>")).rejects.toMatchObject({ code: "InvalidDocument" });
  await expect(parse("<root><?hidden instruction?></root>")).rejects.toMatchObject({ code: "UnsupportedFeature" });
});

it("cancels between XML slices and does not deliver the unfinished root", async () => {
  const controller = new AbortController(); let rootSeen = false;
  await expect(streamXml(encode("<root>" + "<item/>".repeat(20000) + "</root>"), "cancel.xml", { limits, budget: { nodes: 0 }, control: new ImportControl({ signal: controller.signal, onProgress() { controller.abort(); } }), maxRetainedNodes: 10, maxTextCodeUnits: 100, progress,
    complete(node) { if (node.name === "root") rootSeen = true; return node.name !== "item"; },
  })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(rootSeen).toBe(false);
});


it("counts retained text records in metadata part limits as well as the shared XML budget", () => {
  const bytes = encode("<root><item>text</item></root>");
  expect(() => xml(bytes, "metadata.xml", limits, { nodes: 0 }, 2)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(text(xml(bytes, "metadata.xml", limits, { nodes: 0 }, 3))).toBe("text");
});
