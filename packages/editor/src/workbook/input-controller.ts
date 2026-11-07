import { captureWorkbookModel, DocumentError, parseContract } from "@doffi/core";
import type { DocumentSelection, TransactionPrepareOptions, TransactionReceipt, WorkbookSession } from "@doffi/core";
import { parseWorkbookTextInput } from "./text-input.js";
import type { WorkbookTextInputMode } from "./text-input.js";

export interface WorkbookInputState {
  readonly selection?: DocumentSelection;
  readonly text: string;
  readonly mode: WorkbookTextInputMode;
  readonly dirty: boolean;
  readonly composing: boolean;
}
export interface WorkbookInputController {
  state(): WorkbookInputState;
  select(selection: DocumentSelection): void;
  setText(text: string): void;
  setMode(mode: WorkbookTextInputMode): void;
  compositionStart(): void;
  compositionEnd(text: string): void;
  apply(options?: TransactionPrepareOptions): TransactionReceipt;
  cancel(): void;
  dispose(): void;
}
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
/** Transient human input owns neither source cells nor history. Every apply
 * prepares one revision-bound core candidate; failures retain the draft. */
export function createWorkbookInputController(session: WorkbookSession): WorkbookInputController {
  if (session.kind !== "workbook") invalid("Workbook input requires its owning workbook session.");
  let selection: DocumentSelection | undefined, text = "", original = "", mode: WorkbookTextInputMode = "auto", composing = false, disposed = false;
  const check = (): void => { if (disposed) invalid("This workbook input controller has been disposed."); };
  const dirty = (): boolean => text !== original || mode !== "auto";
  function read(): void {
    if (!selection || selection.anchor.kind !== "cells") invalid("Workbook input requires a cell selection.");
    const at = selection.anchor.ranges[0]!.from, model = captureWorkbookModel(session), cell = model.getCell(selection.anchor.sheetId, at.row, at.column);
    text = original = cell?.formula?.expression ?? (!cell || cell.value.type === "blank" ? "" : cell.value.type === "boolean" ? cell.value.value ? "TRUE" : "FALSE" : String(cell.value.value));
    mode = "auto"; composing = false;
  }
  function setText(value: string): void {
    check(); if (!selection || typeof value !== "string") invalid("Workbook text input requires a selected cell and string value.");
    if (value.length > 32767) throw new DocumentError("ResourceLimitExceeded", "Workbook input exceeds its text budget.");
    text = value;
  }
  return Object.freeze({
    state() { check(); return Object.freeze({ ...(selection ? { selection } : {}), text, mode, dirty: dirty(), composing }); },
    select(input: DocumentSelection) {
      check();
      if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).some(key => !["documentId", "baseRevision", "anchor"].includes(key))) invalid("Invalid workbook input selection.");
      if (!input.anchor || input.anchor.kind !== "cells" || !Array.isArray(input.anchor.ranges) || input.anchor.ranges.length !== 1) invalid("Text input requires one worksheet rectangle.");
      const anchor = parseContract("anchor", input.anchor); if (anchor.kind !== "cells") invalid("Expected a cell selection.");
      const range = anchor.ranges[0]!;
      if (range.from.row > range.to.row || range.from.column > range.to.column) invalid("Workbook selection ranges must be ordered.");
      const model = captureWorkbookModel(session), manifest = model.snapshot().manifest;
      if (input.documentId !== manifest.documentId || input.baseRevision !== manifest.revision) throw new DocumentError("RevisionConflict", "Input selection belongs to another workbook revision.");
      if (!model.snapshot().metadata.workbook.sheets.some(sheet => sheet.sheetId === anchor.sheetId)) invalid("Input selection requires an existing worksheet.");
      const next = Object.freeze({ documentId: input.documentId, baseRevision: input.baseRevision,
        anchor: Object.freeze({ kind: "cells" as const, sheetId: anchor.sheetId, ranges: Object.freeze([Object.freeze({ from: Object.freeze({ ...range.from }), to: Object.freeze({ ...range.to }) })]) }) });
      if (JSON.stringify(selection) === JSON.stringify(next)) return;
      if (composing || dirty()) invalid("Finish or cancel the current cell input before changing its selection.");
      selection = next; read();
    },
    setText,
    setMode(value: WorkbookTextInputMode) { check(); if (!selection || !["auto", "text", "number"].includes(value)) invalid("Invalid workbook input mode."); mode = value; },
    compositionStart() { check(); if (!selection) invalid("Composition requires a selected cell."); composing = true; },
    compositionEnd(value: string) { check(); composing = false; setText(value); },
    apply(options?: TransactionPrepareOptions) {
      check(); if (!selection || selection.anchor.kind !== "cells") invalid("Workbook input requires a selected cell.");
      if (composing) invalid("An active composition must finish before a document transaction.");
      const at = selection.anchor.ranges[0]!.from;
      const candidate = session.prepare({ protocolVersion: 1, documentId: selection.documentId, baseRevision: selection.baseRevision,
        operations: dirty() ? [{ op: "setWorkbookCells", sheetId: selection.anchor.sheetId, cells: [{ ...at, source: parseWorkbookTextInput(text, at, mode) }] }] : [] }, options);
      const receipt = session.commit(candidate);
      selection = Object.freeze({ ...selection, baseRevision: receipt.revision }); read(); return receipt;
    },
    cancel() { check(); text = original; mode = "auto"; composing = false; },
    dispose() { if (disposed) return; disposed = true; selection = undefined; text = original = ""; composing = false; },
  });
}
