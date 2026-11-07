import { DocumentError } from "@doffi/core";
import type { DocumentSelection, TransactionPrepareOptions, TransactionReceipt } from "@doffi/core";
import type { WorkbookInputController } from "./input-controller.js";

export type WorkbookNavigationDirection = "up" | "down" | "left" | "right";
export interface WorkbookTextBindingOptions {
  readonly prepareOptions?: () => TransactionPrepareOptions;
  readonly onApplied?: (receipt: TransactionReceipt) => void;
  readonly onCancelled?: () => void;
  readonly onNavigate?: (direction: WorkbookNavigationDirection, extend: boolean) => void;
  readonly onHistory?: (direction: "undo" | "redo") => void;
  readonly onClear?: (selection: DocumentSelection) => void;
  readonly onPaste?: (text: string, selection: DocumentSelection) => void;
  readonly onError?: (error: DocumentError, recoveryText: string) => void;
}
export interface WorkbookTextBinding {
  refresh(): void;
  focus(selectAll?: boolean): void;
  apply(): TransactionReceipt | undefined;
  dispose(): void;
}

/** Binds a host-owned native input to transient text and core transactions.
 * The host owns navigation, clipboard transactions and document history. */
export function bindWorkbookTextInput(input: HTMLTextAreaElement, controller: WorkbookInputController, options: WorkbookTextBindingOptions = {}): WorkbookTextBinding {
  const view = input?.ownerDocument?.defaultView;
  if (!view || !(input instanceof view.HTMLTextAreaElement) || !input.isConnected) throw new DocumentError("InvalidOperation", "Workbook text binding requires a connected native textarea.");
  if (!options || typeof options !== "object" || Array.isArray(options)
    || Object.keys(options).some(key => !["prepareOptions", "onApplied", "onCancelled", "onNavigate", "onHistory", "onClear", "onPaste", "onError"].includes(key))
    || Object.values(options).some(value => value !== undefined && typeof value !== "function")) throw new DocumentError("InvalidOperation", "Invalid workbook text binding options.");
  const abort = new view.AbortController(), events = { signal: abort.signal };
  let disposed = false, editing = false;
  function check(): void { if (disposed) throw new DocumentError("InvalidOperation", "This workbook text binding has been disposed."); }
  function report(error: unknown): void {
    const detail = error instanceof DocumentError ? error : new DocumentError("InvalidOperation", error instanceof Error ? error.message : "Workbook input failed.");
    try { options.onError?.(detail, input.value); } catch { /* Host reporting cannot change source or DOM event completion. */ }
  }
  function notify<T extends unknown[]>(callback: ((...args: T) => void) | undefined, ...args: T): void {
    try { callback?.(...args); } catch (error) { report(error); }
  }
  function refresh(): void {
    check(); const state = controller.state();
    if (input.value !== state.text) input.value = state.text;
    if (!state.dirty && !state.composing) editing = false;
  }
  function apply(): TransactionReceipt | undefined {
    check(); let receipt: TransactionReceipt;
    try { controller.setText(input.value); receipt = controller.apply(options.prepareOptions?.()); refresh(); }
    catch (error) { report(error); return; }
    notify(options.onApplied, receipt); return receipt;
  }
  input.addEventListener("input", () => {
    editing = true;
    try { controller.setText(input.value); } catch (error) { report(error); }
  }, events);
  input.addEventListener("pointerdown", () => { editing = true; }, events);
  input.addEventListener("compositionstart", () => {
    try { controller.compositionStart(); editing = true; } catch (error) { report(error); }
  }, events);
  input.addEventListener("compositionend", () => {
    try { controller.compositionEnd(input.value); } catch (error) { report(error); }
  }, events);
  input.addEventListener("paste", event => {
    const state = controller.state();
    if (editing || state.dirty || state.composing || input.value !== state.text || !state.selection || !options.onPaste || !event.clipboardData?.types.includes("text/plain")) return;
    event.preventDefault(); notify(options.onPaste, event.clipboardData.getData("text/plain"), state.selection);
  }, events);
  input.addEventListener("keydown", event => {
    const state = controller.state();
    if (event.isComposing || event.keyCode === 229 || state.composing) return;
    const command = event.ctrlKey || event.metaKey, key = event.key.toLowerCase();
    if (key === "escape") { event.preventDefault(); controller.cancel(); refresh(); input.select(); notify(options.onCancelled); return; }
    if (key === "f2") { event.preventDefault(); editing = true; input.setSelectionRange(input.value.length, input.value.length); return; }
    if (key === "enter" && event.altKey && !command) {
      event.preventDefault();
      try {
        const from = input.selectionStart, to = input.selectionEnd;
        controller.setText(input.value.slice(0, from) + "\n" + input.value.slice(to));
        input.setRangeText("\n", from, to, "end"); editing = true;
      } catch (error) { report(error); }
      return;
    }
    if ((key === "enter" && !event.altKey) || key === "tab") {
      event.preventDefault(); if (!apply()) return;
      if (!command) notify(options.onNavigate, key === "tab" ? event.shiftKey ? "left" : "right" : event.shiftKey ? "up" : "down", false);
      return;
    }
    if (editing || state.dirty) return;
    if (command && !event.altKey && (key === "z" || key === "y") && options.onHistory) {
      event.preventDefault(); notify(options.onHistory, key === "y" || event.shiftKey ? "redo" : "undo"); return;
    }
    if (command || event.altKey) return;
    const directions: Readonly<Record<string, WorkbookNavigationDirection>> = { arrowup: "up", arrowdown: "down", arrowleft: "left", arrowright: "right" };
    const direction = directions[key];
    if (direction && options.onNavigate) { event.preventDefault(); notify(options.onNavigate, direction, event.shiftKey); return; }
    if ((key === "delete" || key === "backspace") && state.selection && options.onClear) { event.preventDefault(); notify(options.onClear, state.selection); }
  }, events);
  refresh();
  return Object.freeze({
    refresh,
    focus(selectAll = true) {
      check(); if (typeof selectAll !== "boolean") throw new DocumentError("InvalidOperation", "Invalid workbook input focus mode.");
      // Firefox can otherwise retain keyboard focus in the preview iframe
      // even after the parent shadow textarea reports itself focused.
      view!.focus(); input.focus({ preventScroll: true }); editing = !selectAll;
      if (selectAll) input.select(); else input.setSelectionRange(input.value.length, input.value.length);
    },
    apply,
    dispose() { if (disposed) return; disposed = true; abort.abort(); },
  });
}
