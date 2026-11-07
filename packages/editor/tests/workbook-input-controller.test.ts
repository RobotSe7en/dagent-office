import { expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentSelection, WorkbookSession } from "@doffi/core";
import { createWorkbookInputController } from "../src/workbook/input-controller.js";
function session(): WorkbookSession { return createDocumentSession(openDocument(createWorkbook({ documentId: "input-controller" }), { kind: "workbook" })); }
function selection(session: WorkbookSession, row = 0, column = 0): DocumentSelection { return { documentId: "input-controller", baseRevision: session.snapshot().manifest.revision,
  anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row, column }, to: { row, column } }] } }; }
test("composition is transient and publishes exactly one atomic undo step", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.compositionStart();
  controller.setText("中"); expect(() => controller.apply()).toThrow(); expect(source.history().undoSteps).toBe(0);
  controller.compositionEnd("中文😀禰󠄀\nabc"); expect(source.getCell("sheet-1", 0, 0)).toBeUndefined();
  controller.apply(); expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "中文😀禰󠄀\nabc" });
  expect(controller.state()).toMatchObject({ dirty: false, composing: false, selection: { baseRevision: 1 } }); expect(source.history().undoSteps).toBe(1);
  source.undo(); expect(source.getCell("sheet-1", 0, 0)).toBeUndefined(); controller.dispose(); expect(() => controller.state()).toThrow(); expect(source.snapshot().manifest.revision).toBe(2);
});
test("a changed source rejects a late composition while retaining recoverable input", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.compositionStart(); controller.setText("暂存");
  source.commit(source.prepare({ protocolVersion: 1, documentId: "input-controller", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "external" } } }] }] }));
  controller.compositionEnd("本地输入😀"); expect(() => controller.apply()).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  expect(controller.state()).toMatchObject({ dirty: true, text: "本地输入😀", selection: { baseRevision: 0 } }); expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "external" });
});
test("failed formulas, selection replacement and explicit conversion preserve the draft boundary", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.setText("=");
  expect(() => controller.apply()).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(controller.state().text).toBe("=");
  expect(() => controller.select(selection(source, 1))).toThrow(); controller.cancel(); controller.select(selection(source, 1));
  controller.setText("00123"); controller.apply(); expect(source.getCell("sheet-1", 1, 0)?.value).toEqual({ type: "text", value: "00123" });
  controller.setMode("number"); controller.apply(); expect(source.getCell("sheet-1", 1, 0)?.value).toEqual({ type: "number", value: 123 });
  expect(controller.state()).toMatchObject({ mode: "auto", dirty: false });
});
test("idle apply preserves explicitly typed text and empty strings without a new revision", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.setMode("text"); controller.setText("12"); controller.apply();
  const revision = source.snapshot().manifest.revision; expect(controller.apply()).toMatchObject({ changed: false, revision });
  expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "12" });
  controller.select(selection(source, 1)); controller.setMode("text"); controller.apply(); const emptyRevision = source.snapshot().manifest.revision;
  expect(controller.apply()).toMatchObject({ changed: false, revision: emptyRevision }); expect(source.getCell("sheet-1", 1, 0)?.value).toEqual({ type: "text", value: "" });
});
