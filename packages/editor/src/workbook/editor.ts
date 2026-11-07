import { captureWorkbookModel, DocumentError } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, DocumentSelection, TransactionPrepareOptions, TransactionReceipt, WorkbookRange, WorkbookSession } from "@doffi/core";
import { createWorkbookPreviewSource, mountDocumentPreview } from "@doffi/render";
import type { WorkbookPreview, WorkbookRenderReport } from "@doffi/render";
import type { WorkbookEditorOptions, WorkbookEditorView } from "./editor-contracts.js";
import { createWorkbookEditorUi } from "./editor-ui.js";
import { createWorkbookInputController } from "./input-controller.js";
import { bindWorkbookTextInput } from "./text-binding.js";
import { createWorkbookPasteTransaction } from "./text-input.js";
import type { WorkbookTextInputMode } from "./text-input.js";
import { createWorkbookFillTransaction } from "./fill.js";

/** Direct small-document host integration. The host owns the sole session;
 * source cells, candidates and history remain behind its core boundary. */
export function mountWorkbookEditor(container: HTMLElement, session: WorkbookSession, options: WorkbookEditorOptions = {}): WorkbookEditorView {
  const view = container?.ownerDocument?.defaultView;
  if (!view || !container.isConnected || session.kind !== "workbook") throw new DocumentError("InvalidOperation", "Workbook editing requires a connected container and its owning workbook session.");
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["title", "widthPt", "heightPt", "locale", "limits", "prepareOptions", "onError", "onSelectionChange"].includes(key))) throw new DocumentError("InvalidOperation", "Invalid workbook editor options.");
  const config = { ...options }, ui = createWorkbookEditorUi(container, config.title ?? "工作簿编辑"), input = createWorkbookInputController(session), events = new view.AbortController();
  let disposed = false, choosing = 0, refreshQueued = false, forceCalculation = false, latest: Promise<void> = Promise.resolve();
  let interaction: Promise<void> = Promise.resolve();
  let revision = -1, generation = 0, calculationAbort: AbortController | undefined, commandAbort: AbortController | undefined, preview: WorkbookPreview;
  const buttons: { readonly button: HTMLButtonElement; readonly available: () => boolean }[] = [];
  function check(): void { if (disposed) throw new DocumentError("InvalidOperation", "This workbook editor has been disposed."); }
  function preparation(input?: TransactionPrepareOptions): TransactionPrepareOptions {
    const supplied = input ?? config.prepareOptions?.() ?? {};
    return config.locale !== undefined && supplied.calculationContext?.locale === undefined ? { ...supplied, calculationContext: { ...supplied.calculationContext, locale: config.locale } } : supplied;
  }
  function runInteraction<T>(action: () => Promise<T>): Promise<T> { const work = interaction.then(action); interaction = work.then(() => {}, report); return work; }
  function pendingInput(): boolean { return binding.hasPendingInput(); }
  function report(error: unknown): void {
    if (disposed) return; const problem = error instanceof Error ? error : new Error("工作簿编辑失败。");
    ui.error.textContent = problem instanceof DocumentError && problem.code === "RevisionConflict" ? "文档已变化。当前输入已保留，请先复制输入或取消后重新选择。" : problem.message; ui.error.hidden = false;
    try { config.onError?.(problem); } catch { /* Host callback does not alter the edit. */ }
  }
  function controls(): void {
    if (disposed) return;
    for (const { button, available } of buttons) button.disabled = !available();
    const state = input.state(); ui.mode.value = state.mode; ui.mode.disabled = Boolean(commandAbort) || state.composing || !state.selection;
    ui.address.disabled = Boolean(commandAbort) || revision < 0;
  }
  function notify(): void { try { config.onSelectionChange?.(input.state().selection); } catch { /* Host-owned callback. */ } }
  function address(range: DeepReadonly<WorkbookRange>): string {
    const cell = (point: WorkbookRange["from"]) => { let column = point.column + 1, name = ""; while (column) { column--; name = String.fromCharCode(65 + column % 26) + name; column = Math.floor(column / 26); } return name + (point.row + 1); };
    return cell(range.from) + (range.from.row === range.to.row && range.from.column === range.to.column ? "" : ":" + cell(range.to));
  }
  function acceptSelection(selected: DocumentSelection, focus: boolean): void {
    input.select(selected); binding.refresh(); ui.input.disabled = false;
    if (selected.anchor.kind === "cells") ui.address.value = address(selected.anchor.ranges[0]!);
    controls(); notify(); if (focus) binding.focus();
  }
  function readCurrentSelection(): void {
    if (pendingInput()) return;
    const selected = input.state().selection, snapshot = session.snapshot();
    if (!selected || selected.anchor.kind !== "cells") return;
    const sheetId = selected.anchor.sheetId;
    if (!snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === sheetId)) return;
    acceptSelection({ ...selected, baseRevision: snapshot.manifest.revision }, ui.input === ui.input.ownerDocument.activeElement || ui.element === ui.element.ownerDocument.activeElement);
  }
  async function choose(sheetId: string, range: DeepReadonly<WorkbookRange>, focus = false): Promise<DocumentSelection> {
    check(); if (pendingInput()) throw new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入。");
    await sourceSettled(); check();
    choosing++;
    try {
      if (preview.layout().sheetId !== sheetId) await preview.showSheet(sheetId);
      const selected = await preview.selectRange(range); check(); acceptSelection(selected, focus); return selected;
    } finally { choosing--; }
  }
  async function restoreSelection(): Promise<void> {
    const selected = input.state().selection;
    if (selected?.anchor.kind !== "cells") return;
    const anchor = selected.anchor; await sourceSettled(); check();
    if (!preview.snapshot().metadata.workbook.sheets.some(sheet => sheet.sheetId === anchor.sheetId)) return;
    choosing++;
    try { if (preview.layout().sheetId !== anchor.sheetId) await preview.showSheet(anchor.sheetId); await preview.selectRange(anchor.ranges[0]!, false); if (pendingInput()) binding.focus(false); }
    finally { choosing--; }
  }
  function sourceSelection(selected: DocumentSelection | undefined): void {
    if (disposed || choosing) return;
    commandAbort?.abort();
    try {
      if (!selected) {
        const previous = input.state().selection, sheetId = preview.layout().sheetId;
        if (previous?.anchor.kind !== "cells" || previous.anchor.sheetId === sheetId) return;
        if (pendingInput()) { report(new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入。")); void runInteraction(() => restoreSelection()); }
        else void runInteraction(() => choose(sheetId, { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }));
        return;
      }
      if (selected.anchor.kind !== "cells") return;
      if (pendingInput()) {
        if (!binding.apply()) { void runInteraction(() => restoreSelection()); return; }
        const anchor = selected.anchor; void runInteraction(() => choose(anchor.sheetId, anchor.ranges[0]!, true));
      } else { acceptSelection(selected, true); ui.error.hidden = true; }
    } catch (error) { report(error); }
  }
  async function refresh(force: boolean): Promise<void> {
    check(); const ticket = ++generation, abort = new view!.AbortController(); calculationAbort?.abort(); calculationAbort = abort;
    const model = captureWorkbookModel(session), snapshot = model.snapshot(); ui.status.textContent = "正在更新工作簿…";
    const prepared = preparation();
    const calculation = force || snapshot.metadata.workbook.calculationMode === "auto" ? await session.calculate({ signal: abort.signal,
      yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)), ...(prepared.calculationContext ? { context: prepared.calculationContext } : {}) }) : undefined;
    if (disposed || ticket !== generation || abort.signal.aborted) return;
    preview.update(createWorkbookPreviewSource(model, calculation)); await preview.settled();
    if (disposed || ticket !== generation) return;
    revision = snapshot.manifest.revision;
    const state = input.state();
    if (!pendingInput()) {
      const anchor = state.selection?.anchor, sheetId = preview.layout().sheetId;
      const range = anchor?.kind === "cells" && anchor.sheetId === sheetId ? anchor.ranges[0]! : { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } };
      choosing++;
      try { const selected = await preview.selectRange(range, false); if (!disposed && ticket === generation && !pendingInput()) acceptSelection(selected, false); }
      finally { choosing--; }
    }
    ui.status.textContent = pendingInput() && state.selection?.baseRevision !== revision ? "文档已更新；当前输入仍绑定原修订。" : "修订 " + revision + (snapshot.metadata.workbook.calculationMode === "manual" && !force ? " · 手动计算" : ""); controls();
  }
  function scheduleRefresh(force = false): Promise<void> {
    check(); forceCalculation ||= force; calculationAbort?.abort();
    if (!refreshQueued) {
      refreshQueued = true;
      const work = Promise.resolve().then(() => { refreshQueued = false; const force = forceCalculation; forceCalculation = false; return refresh(force); });
      latest = work; void work.catch(error => { if (work === latest) report(error); });
    }
    return latest;
  }
  async function sourceSettled(): Promise<WorkbookRenderReport> {
    check();
    for (;;) { const work = latest; try { await work; } catch (error) { if (work === latest) throw error; } if (work === latest) { const result = await preview.settled(); if (work === latest) return result; } }
  }
  async function settled(): Promise<WorkbookRenderReport> {
    for (;;) { const work = interaction; await work; const result = await sourceSettled(); if (work === interaction) return result; }
  }
  function perform(operations: readonly DocumentOperation[], preparedOptions?: TransactionPrepareOptions): TransactionReceipt {
    check(); if (pendingInput()) throw new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入。");
    const manifest = session.snapshot().manifest;
    const receipt = session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [...operations] }, preparation(preparedOptions)));
    readCurrentSelection(); controls(); ui.error.hidden = true; return receipt;
  }
  function cancelInput(): void { check(); input.cancel(); binding.refresh(); readCurrentSelection(); controls(); ui.error.hidden = true; }
  function history(direction: "undo" | "redo"): void {
    check(); if (pendingInput()) throw new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入。");
    session[direction](); readCurrentSelection(); controls(); ui.error.hidden = true;
  }
  async function fill(axis: "row" | "column", series = false): Promise<void> {
    check(); if (pendingInput() || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    const selected = input.state().selection;
    if (selected?.anchor.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择填充区域。");
    const range = selected.anchor.ranges[0]!, count = range.to[axis] - range.from[axis] + 1;
    if (series && range.from.row !== range.to.row && range.from.column !== range.to.column) throw new DocumentError("InvalidOperation", "数列填充请选择单行或单列。");
    if (count < 2) throw new DocumentError("InvalidOperation", axis === "row" ? "向下填充需要至少两行。" : "向右填充需要至少两列。");
    const model = captureWorkbookModel(session), manifest = model.snapshot().manifest, sheetId = selected.anchor.sheetId;
    if (selected.documentId !== manifest.documentId || selected.baseRevision !== manifest.revision) throw new DocumentError("RevisionConflict", "Fill selection belongs to a previous source.");
    let seeds = 1;
    if (series) {
      const next = { ...range.from, [axis]: range.from[axis] + 1 }, value = model.getCell(sheetId, next.row, next.column);
      if (value?.formula || value && value.value.type !== "number" && value.value.type !== "blank") throw new DocumentError("InvalidOperation", "数列填充的种子必须是数字或空白。");
      if (value?.value.type === "number") seeds = 2;
    }
    const source = { from: { ...range.from }, to: { ...range.to, [axis]: range.from[axis] + seeds - 1 } };
    const target = { from: { ...range.from, [axis]: range.from[axis] + seeds }, to: { ...range.to } };
    if (target.from[axis] > target.to[axis]) return;
    const abort = new view!.AbortController(); commandAbort = abort; ui.input.disabled = true; controls(); ui.error.hidden = true;
    try {
      const transaction = await createWorkbookFillTransaction(session, sheetId, source, target, { mode: series ? "series" : "copy", axis, signal: abort.signal,
        yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)), onProgress: progress => { if (!disposed) ui.status.textContent = `正在填充 ${progress.completed} / ${progress.total}`; } });
      if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Workbook fill was cancelled.");
      const receipt = session.commit(session.prepare(transaction, preparation())); readCurrentSelection();
      if (!receipt.changed) ui.status.textContent = "选区没有变化。";
    } catch (error) {
      if (!(error instanceof DocumentError && error.code === "OperationCancelled" && abort.signal.aborted)) throw error;
      if (!disposed) ui.status.textContent = "操作已取消。";
    } finally { if (commandAbort === abort) commandAbort = undefined; if (!disposed) { ui.input.disabled = !input.state().selection; controls(); } }
  }
  const binding = bindWorkbookTextInput(ui.input, input, {
    prepareOptions: () => preparation(),
    onApplied() { controls(); ui.error.hidden = true; }, onCancelled: cancelInput,
    onNavigate(direction, extend) { void runInteraction(async () => {
      await sourceSettled(); if (pendingInput()) return; choosing++;
      try { const selected = await preview.moveSelection(direction, extend); check(); acceptSelection(selected, true); } finally { choosing--; }
    }); },
    onHistory: history,
    onClear(selected) {
      if (selected.anchor.kind !== "cells") return;
      const manifest = session.snapshot().manifest;
      if (selected.documentId !== manifest.documentId || selected.baseRevision !== manifest.revision) throw new DocumentError("RevisionConflict", "The selected clear range belongs to a previous revision.");
      perform([{ op: "clearWorkbookRanges", sheetId: selected.anchor.sheetId, ranges: selected.anchor.ranges.map(range => ({ from: { ...range.from }, to: { ...range.to } })), mode: "contents" }]);
    },
    onPaste(text, selected) {
      if (selected.anchor.kind !== "cells") return;
      const transaction = createWorkbookPasteTransaction(session.snapshot(), selected.anchor.sheetId, selected.anchor.ranges[0]!.from, text);
      const receipt = session.commit(session.prepare({ ...transaction, documentId: selected.documentId, baseRevision: selected.baseRevision }, preparation()));
      readCurrentSelection(); controls();
      if (receipt.changed) { const cells = transaction.operations[0]; if (cells?.op === "setWorkbookCells") {
        const last = cells.cells.at(-1)!, anchor = selected.anchor; void runInteraction(() => choose(anchor.sheetId, { from: anchor.ranges[0]!.from, to: { row: last.row, column: last.column } }, true));
      } }
    },
    onError(error) { report(error); controls(); },
  });
  try {
    preview = mountDocumentPreview(ui.preview, session, { title: "工作簿网格", retainSelection: true, ...(config.widthPt === undefined ? {} : { widthPt: config.widthPt }),
      ...(config.heightPt === undefined ? {} : { heightPt: config.heightPt }), ...(config.locale === undefined ? {} : { locale: config.locale }),
      ...(config.limits === undefined ? {} : { limits: config.limits }), onError: error => report(new DocumentError(error.code, error.message, error.details)), onSelectionChange: sourceSelection });
  } catch (error) { binding.dispose(); input.dispose(); ui.element.remove(); throw error; }
  const unsubscribe = session.subscribe(event => { if (event.type === "documentChanged" || event.type === "annotationChanged") { commandAbort?.abort(); void scheduleRefresh(); } });
  const button = (label: string, action: () => unknown, available: () => boolean): void => {
    const button = ui.element.ownerDocument.createElement("button"); button.type = "button"; button.textContent = label;
    button.addEventListener("click", () => { try { action(); } catch (error) { report(error); } }, { signal: events.signal }); ui.tools.append(button); buttons.push({ button, available });
  };
  const idle = (): boolean => !commandAbort && !pendingInput();
  button("确认输入", () => binding.apply(), () => !commandAbort && Boolean(input.state().selection) && !input.state().composing);
  button("取消输入", cancelInput, () => !commandAbort && Boolean(input.state().selection));
  button("撤销", () => history("undo"), () => idle() && session.history().undoSteps > 0);
  button("重做", () => history("redo"), () => idle() && session.history().redoSteps > 0);
  button("重新计算", () => { void scheduleRefresh(true).catch(() => {}); }, idle);
  button("向下填充", () => { void runInteraction(() => fill("row")); }, idle);
  button("向右填充", () => { void runInteraction(() => fill("column")); }, idle);
  button("数列填充", () => { void runInteraction(() => { const anchor = input.state().selection?.anchor; return fill(anchor?.kind === "cells" && anchor.ranges[0]!.from.column === anchor.ranges[0]!.to.column ? "row" : "column", true); }); }, idle);
  button("取消操作", () => commandAbort?.abort(), () => Boolean(commandAbort));
  for (const type of ["input", "keydown", "compositionstart", "compositionend"]) ui.input.addEventListener(type, controls, { signal: events.signal });
  ui.mode.addEventListener("change", () => { try { input.setMode(ui.mode.value as WorkbookTextInputMode); binding.refresh(); controls(); } catch (error) { report(error); } }, { signal: events.signal });
  ui.address.addEventListener("keydown", event => {
    if (event.key !== "Enter") return; event.preventDefault();
    try {
      const sheetId = preview.layout().sheetId, resolved = session.resolveReference("=" + ui.address.value.trim(), { sheetId, row: 0, column: 0 });
      if (resolved.kind !== "ranges" || resolved.areas.length !== 1) throw new DocumentError("InvalidOperation", "请输入一个有效的单元格或矩形区域地址。");
      const area = resolved.areas[0]!; void runInteraction(() => choose(area.sheetId, { from: area.from, to: area.to }, true));
    } catch (error) { report(error); }
  }, { signal: events.signal });
  controls(); void scheduleRefresh(); const ready = sourceSettled().then(() => {}); void ready.catch(() => {});
  return Object.freeze({ element: ui.element, ready, settled, selection() { check(); return input.state().selection; }, inputState() { check(); return input.state(); },
    selectRange: (sheetId: string, range: DeepReadonly<WorkbookRange>) => runInteraction(() => choose(sheetId, range)), focusInput(selectAll = true) { check(); binding.focus(selectAll); }, applyInput: binding.apply, cancelInput, perform,
    recalculate: async () => { await scheduleRefresh(true); return settled(); },
    dispose() { if (disposed) return; disposed = true; ++generation; calculationAbort?.abort(); commandAbort?.abort(); unsubscribe(); events.abort(); binding.dispose(); input.dispose(); preview.dispose(); ui.element.remove(); },
  });
}
