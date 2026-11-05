import { createWorkbookNumberFormatter, DocumentError } from "@dagent-office/core";
import type { DocumentSelection, WorkbookSnapshot } from "@dagent-office/core";
import { createWorkbookPreviewChrome, workbookColumnLabel } from "./preview-chrome.js";
import type { WorkbookPreview, WorkbookPreviewInput, WorkbookPreviewOptions, WorkbookRenderReport } from "./preview-contracts.js";
import { decorateWorkbookPreview } from "./preview-decoration.js";
import type { WorkbookPreviewDecoration } from "./preview-decoration.js";
import { workbookPreviewSource } from "./preview-source.js";
import { hitWorkbookCell, WorkbookMergeIndex } from "./projection.js";
import { createWorkbookSurface } from "./surface.js";
import type { WorkbookSurface } from "./surface.js";
import { WorkbookWindowController } from "./window-controller.js";
import type { WorkbookRenderedWindow } from "./window-controller.js";
import type { WorkbookViewport } from "./viewport.js";

interface PreviewWindow { readonly controls: WorkbookPreviewDecoration; commitChrome(): void }
interface PreviewContext {
  readonly snapshot: WorkbookSnapshot;
  readonly surface: WorkbookSurface;
  readonly controller: WorkbookWindowController;
  readonly windows: WeakMap<HTMLElement, PreviewWindow>;
  readonly slot: HTMLElement;
  dispose(): void;
}
function viewportOf(window: WorkbookRenderedWindow): WorkbookViewport {
  const { widthPt, heightPt, scrollLeftPt, scrollTopPt } = window.frame.layout;
  return Object.freeze({ widthPt, heightPt, scrollLeftPt, scrollTopPt });
}
function reportOf(window: WorkbookRenderedWindow): WorkbookRenderReport {
  const counts = { imported: 0, valid: 0, stale: 0, unavailable: 0 }, seen = new Set<number>();
  for (const cell of window.frame.projections) {
    const { row, column } = cell.range.from, key = row * 16_384 + column;
    if (seen.has(key)) continue; seen.add(key);
    const formula = window.frame.getCell(row, column)?.formula;
    if (formula) counts[formula.cacheState]++;
  }
  const { documentId, revision, sheetId, storedCellCount, layout } = window.frame;
  const ruleFormulaValues = { current: 0, unavailable: 0 };
  for (const cell of window.rules?.cells ?? []) if (seen.has(cell.row * 16_384 + cell.column) && window.frame.getCell(cell.row, cell.column)?.formula)
    ruleFormulaValues[cell.current.kind === "value" ? "current" : "unavailable"]++;
  return Object.freeze({ kind: "workbook", layout: "grid", documentId, revision, sheetId, viewport: viewportOf(window),
    maxScrollLeftPt: layout.maxScrollLeftPt, maxScrollTopPt: layout.maxScrollTopPt, visibleCellCount: seen.size, storedCellCount,
    objectFragmentCount: window.objects.length, formulaCaches: Object.freeze(counts), ruleFormulaValues: Object.freeze(ruleFormulaValues), diagnostics: window.diagnostics, fonts: window.fonts,
    usedFallbackLineMetrics: window.usedFallbackLineMetrics });
}

/** One preview owns its surfaces, candidates and format caches. Authorized
 * providers (including their Worker) remain owned by the host. */
export function mountWorkbookPreview(container: HTMLElement, input: WorkbookPreviewInput, options: WorkbookPreviewOptions = {}): WorkbookPreview {
  if (Object.keys(options).some(key => !["title", "widthPt", "heightPt", "sheetId", "locale", "limits", "onError", "onSelectionChange"].includes(key)))
    throw new DocumentError("InvalidOperation", "Workbook preview received an unsupported option.");
  const config = { ...options, limits: { ...options.limits, read: { ...options.limits?.read }, cells: { ...options.limits?.cells },
    annotations: { ...options.limits?.annotations }, rules: { ...options.limits?.rules }, drawings: { ...options.limits?.drawings, assets: { ...options.limits?.drawings?.assets } } } };
  let desired: WorkbookViewport = Object.freeze({ widthPt: config.widthPt ?? 600, heightPt: config.heightPt ?? 360, scrollLeftPt: 0, scrollTopPt: 0 });
  let active: PreviewContext | undefined, pending: { cancel(): void } | undefined, disposed = false, generation = 0;
  let selection: DocumentSelection | undefined, latest: Promise<void> = Promise.resolve();
  let resolveReady!: () => void, rejectReady!: (error: unknown) => void;
  const ready = new Promise<void>((resolve, reject) => { resolveReady = resolve; rejectReady = reject; });
  // Hosts may observe ready after mounting other UI. Do not create an unhandled
  // rejection in the meantime; the original promise still rejects to callers.
  void ready.catch(() => {});
  const check = (): void => { if (disposed) throw new DocumentError("InvalidOperation", "This workbook preview has been disposed."); };
  const current = (): { context: PreviewContext; window: WorkbookRenderedWindow; ui: PreviewWindow } => {
    check(); const window = active?.controller.current;
    if (!active || !window) throw new DocumentError("InvalidOperation", "The workbook preview has no accepted window yet.");
    return { context: active, window, ui: active.windows.get(window.element)! };
  };
  const errorOf = (error: unknown): DocumentError => error instanceof DocumentError ? error : new DocumentError("InvalidOperation", "The workbook preview could not complete its authorized read.");
  const failed = (error: unknown): void => {
    if (disposed) return;
    const problem = errorOf(error); chrome.status("预览未更新：" + problem.code, true);
    try { config.onError?.(problem.toJSON()); } catch { /* Host callbacks do not own the renderer lifecycle. */ }
  };
  const attempt = (action: () => void): void => { if (disposed || pending) return; try { action(); } catch (error) { failed(error); } };
  const chrome = createWorkbookPreviewChrome(container, config.title ?? "工作簿只读预览", {
    sheet: id => { void showSheet(id).catch(() => {}); },
    scroll: (left, top) => { if (!pending) void setViewport({ ...desired, scrollLeftPt: left, scrollTopPt: top }).catch(() => {}); },
    comments: () => attempt(() => {
      const { context, window, ui } = current(), sheets = new Set(context.snapshot.metadata.workbook.sheets.map(sheet => sheet.sheetId));
      ui.controls.showComments(context.snapshot.annotations.filter(annotation => annotation.anchor.kind === "cells"
        && (annotation.anchor.sheetId === window.frame.sheetId || annotation.status === "needs_reanchor" && !sheets.has(annotation.anchor.sheetId))).map(annotation => annotation.annotationId));
    }),
  });
  const clearSelection = (): void => {
    chrome.cell("", ""); chrome.validation("");
    if (!selection) return; selection = undefined;
    const ticket = generation;
    queueMicrotask(() => { if (!disposed && ticket === generation && !selection) {
      try { config.onSelectionChange?.(undefined); } catch { /* Host-owned callback. */ }
    } });
  };
  const accepted = (context: PreviewContext): void => {
    const window = context.controller.current!;
    context.windows.get(window.element)!.commitChrome(); desired = viewportOf(window); clearSelection();
    const report = reportOf(window), counts = report.formulaCaches;
    chrome.status(report.diagnostics.some(item => ["rule-current-value-unavailable", "conditional-rule-unavailable", "validation-unavailable"].includes(item.code))
      ? "部分条件格式或数据验证无法按当前值求出。"
      : report.diagnostics.some(item => item.code === "validation-invalid") ? "当前窗口含不符合数据验证的来源值，选中单元格可查看提示。"
      : report.ruleFormulaValues.current === Object.values(counts).reduce((a, b) => a + b, 0) && report.ruleFormulaValues.current > 0 ? "当前窗口的公式已按当前值求出。"
      : counts.stale || counts.unavailable ? "当前窗口含过期或不可用的公式结果。"
      : counts.imported ? "当前窗口含导入的公式缓存。" : "只读预览");
    resolveReady();
  };
  const observe = (work: Promise<void>, ticket: number): void => {
    latest = work;
    void work.catch(error => {
      if (disposed || ticket !== generation) return;
      if (active?.controller.current) desired = viewportOf(active.controller.current);
      failed(error); if (!active && ticket === generation) rejectReady(errorOf(error));
    });
  };
  async function settled(): Promise<WorkbookRenderReport> {
    check();
    for (;;) { const work = latest; try { await work; } catch (error) { if (work === latest) throw errorOf(error); } if (work === latest) return reportOf(current().window); }
  }
  function setViewport(viewport: WorkbookViewport): Promise<WorkbookRenderReport> {
    check();
    if (pending) return Promise.reject(new DocumentError("InvalidOperation", "A replacement workbook is still loading."));
    const { context, window } = current(); return render(context, window.frame.sheetId, viewport);
  }
  function showSheet(sheetId: string): Promise<WorkbookRenderReport> {
    check();
    if (pending) return Promise.reject(new DocumentError("InvalidOperation", "A replacement workbook is still loading."));
    return render(current().context, sheetId, { ...desired, scrollLeftPt: 0, scrollTopPt: 0 });
  }
  function render(context: PreviewContext, sheetId: string, viewport: WorkbookViewport): Promise<WorkbookRenderReport> {
    context.controller.cancelPending();
    const ticket = ++generation;
    desired = Object.freeze({ widthPt: viewport.widthPt, heightPt: viewport.heightPt, scrollLeftPt: viewport.scrollLeftPt, scrollTopPt: viewport.scrollTopPt });
    const work = (async () => { if (await context.controller.render(sheetId, desired) && !disposed && ticket === generation && active === context) accepted(context); })();
    observe(work, ticket); return settled();
  }
  function update(next: WorkbookPreviewInput): void {
    check();
    const providers = workbookPreviewSource(next), snapshot = providers.cells.snapshot;
    const viewport = { ...desired }, sheetId = config.sheetId ?? snapshot.metadata.workbook.activeSheetId;
    pending?.cancel(); active?.controller.cancelPending(); const ticket = ++generation, abort = new AbortController();
    let candidate: PreviewContext | undefined;
    const request = { cancel() { abort.abort(); candidate?.dispose(); } }; pending = request; chrome.busy(true);
    const work = (async () => {
      const slot = container.ownerDocument.createElement("div");
      slot.style.cssText = "position:absolute;inset:0;opacity:0;pointer-events:none;overflow:hidden";
      slot.setAttribute("aria-hidden", "true"); chrome.canvas.append(slot);
      let surface: WorkbookSurface | undefined, formatter: ReturnType<typeof createWorkbookNumberFormatter> | undefined;
      try {
        surface = await createWorkbookSurface(slot, snapshot, { ...viewport, signal: abort.signal });
        if (abort.signal.aborted) return;
        formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem,
          ...(config.locale !== undefined ? { hostLocale: config.locale } : {}), ...(snapshot.metadata.locale ? { documentLocale: snapshot.metadata.locale } : {}) });
        const windows = new WeakMap<HTMLElement, PreviewWindow>(), ownedSurface = surface, ownedFormatter = formatter;
        const controller = new WorkbookWindowController(surface, providers, formatter, config.limits, (window, content, maxNodes) => {
          const commitChrome = chrome.prepare(window, snapshot);
          const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === window.frame.sheetId)!, merges = new WorkbookMergeIndex(sheet.merges);
          const controls = decorateWorkbookPreview(ownedSurface, snapshot, window, content, maxNodes, attempt);
          const events = new ownedSurface.document.defaultView!.AbortController();
          windows.set(window.element, { controls, commitChrome });
          window.element.addEventListener("wheel", event => {
            if (event.ctrlKey || pending || active?.controller.current?.element !== window.element || (event.target as Element).closest('[data-office-view="workbook-preview-comments"]')) return;
            event.preventDefault();
            const factor = event.deltaMode === 1 ? 12 : event.deltaMode === 2 ? window.frame.layout.heightPt : .75;
            const x = event.shiftKey && !event.deltaX ? event.deltaY : event.deltaX, y = event.shiftKey && !event.deltaX ? 0 : event.deltaY;
            void setViewport({ ...desired, scrollLeftPt: desired.scrollLeftPt + x * factor, scrollTopPt: desired.scrollTopPt + y * factor }).catch(() => {});
          }, { passive: false, signal: events.signal });
          window.element.addEventListener("click", event => attempt(() => {
            if (active?.controller.current?.element !== window.element || (event.target as Element).closest('button,[data-office-view="workbook-preview-comments"],[data-office-view="workbook-drawings"]')) return;
            const rect = ownedSurface.view.getBoundingClientRect(), x = (event.clientX - rect.left) * .75, y = (event.clientY - rect.top) * .75;
            const cell = hitWorkbookCell(window.dimensions, window.frame.layout, merges, sheet.rightToLeft ? window.frame.layout.widthPt - x : x, y);
            if (!cell) return;
            controls.selectCell(cell.row, cell.column);
            const range = merges.at(cell.row, cell.column) ?? { from: cell, to: cell }, value = window.frame.getCell(cell.row, cell.column);
            const anchor = Object.freeze({ kind: "cells" as const, sheetId: sheet.sheetId, ranges: Object.freeze([Object.freeze({ from: Object.freeze({ ...range.from }), to: Object.freeze({ ...range.to }) })]) });
            selection = Object.freeze({ documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, anchor });
            chrome.cell(workbookColumnLabel(cell.column) + (cell.row + 1), value?.formula?.expression ?? (value?.value.type === "blank" || !value ? "" : String(value.value.value)));
            const validation = window.rules?.cells.find(item => item.row === cell.row && item.column === cell.column)?.validation;
            const definition = validation && sheet.dataValidations?.find(rule => rule.ruleId === validation.ruleId);
            const messages: string[] = [];
            if (definition && validation) {
              if (!sheet.disableValidationPrompts && definition.showInputMessage) messages.push(...[definition.promptTitle, definition.prompt].filter(Boolean));
              if (validation.state !== "none") messages.push(validation.state === "valid" ? "当前值符合验证规则。" : validation.state === "invalid" ? "当前值不符合验证规则。" : "当前值无法校验。" + (validation.issue?.functionName ? " 未支持函数：" + validation.issue.functionName : ""));
              if (!sheet.disableValidationPrompts && validation.state === "invalid" && definition.showErrorMessage) messages.push(...[definition.errorTitle, definition.error].filter(Boolean));
            }
            chrome.validation(messages.join("\n"));
            try { config.onSelectionChange?.(selection); } catch { /* Host-owned callback. */ }
          }), { signal: events.signal });
          return { get nodes() { return controls.nodes; }, get fonts() { return controls.fonts; }, dispose() { events.abort(); controls.dispose(); } };
        });
        let released = false;
        candidate = { snapshot, surface, controller, windows, slot, dispose() {
          if (released) return; released = true; controller.dispose(); ownedFormatter.dispose(); ownedSurface.dispose(); slot.remove();
        } };
        if (!await controller.render(sheetId, viewport) || abort.signal.aborted || disposed || ticket !== generation) return;
        const previous = active;
        // Chrome preparation and all controls already succeeded before the
        // controller published this isolated candidate's complete window.
        active = candidate; accepted(candidate); candidate = undefined;
        slot.style.opacity = "1"; slot.style.pointerEvents = "auto"; slot.removeAttribute("aria-hidden"); previous?.dispose();
        surface = undefined; formatter = undefined;
      } catch (error) { if (!abort.signal.aborted && !disposed && ticket === generation) throw error; }
      finally {
        candidate?.dispose(); surface?.dispose(); formatter?.dispose();
        if (active?.slot !== slot) slot.remove();
        if (pending === request) { pending = undefined; chrome.busy(false); }
      }
    })();
    observe(work, ticket);
  }
  try { update(input); } catch (error) { chrome.dispose(); throw error; }
  return Object.freeze({ kind: "workbook", element: chrome.element, ready, snapshot: () => current().context.snapshot,
    layout: () => reportOf(current().window), fonts: () => current().window.fonts, selection: () => { check(); return selection; }, settled, update, showSheet, setViewport,
    openComment(annotationId: string) { current().ui.controls.openComment(annotationId); }, closeComment() { current().ui.controls.closeComment(); },
    dispose() { if (disposed) return; disposed = true; ++generation; pending?.cancel(); pending = undefined; active?.dispose(); active = undefined;
      selection = undefined; chrome.dispose(); rejectReady(new DocumentError("OperationCancelled", "Workbook preview was disposed before becoming ready.")); },
  });
}
