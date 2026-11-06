import { createHash } from "node:crypto";
import { expect, test } from "vitest";
import { createWorkbook, openDocument } from "@dagent-office/core";
import { readDocumentZip, writeDocumentZip } from "@dagent-office/office";

test("workbook package ZIP preserves hashed chunks and forwards workbook budgets", () => {
  const encoder = new TextEncoder(), files = new Map(createWorkbook({ documentId: "zip-book" }));
  const bytes = encoder.encode('<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="0" data-column="0"><tbody><tr data-row="0"><td data-column="0" data-value-type="number" data-value="12">12</td><td data-column="1" data-value-type="text">编号 001</td></tr></tbody></table></body></html>');
  const path = "sheets/first.html", manifest = JSON.parse(new TextDecoder().decode(files.get("manifest.json")!));
  files.set(path, bytes); manifest.chunks.push({ path, sheetId: "sheet-1", row: 0, column: 0, sha256: createHash("sha256").update(bytes).digest("hex") });
  files.set("manifest.json", encoder.encode(JSON.stringify(manifest)));
  expect(() => writeDocumentZip(files, { workbookLimits: { maxCells: 1 } })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  const reopened = openDocument(readDocumentZip(writeDocumentZip(files)), { kind: "workbook" });
  expect(reopened.getCell("sheet-1", 0, 1)?.value).toEqual({ type: "text", value: "编号 001" });
  expect(reopened.serialize().get(path)).toEqual(bytes);
});
