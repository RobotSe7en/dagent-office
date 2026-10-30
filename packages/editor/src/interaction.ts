import { createOpaqueId, DocumentError } from "@dagent-office/core";
import type {
  Anchor, DeepReadonly, DocumentEditingSession, DocumentModel, DocumentOperation, DocumentSelection, Rectangle,
} from "@dagent-office/core";
import { mountPreview } from "./preview.js";
import type { DocumentPreviewOptions } from "./preview.js";
import type { DocumentRenderReport } from "@dagent-office/render";
import type { FontResolution } from "./fonts.js";
import type { PreviewSource } from "./source-index.js";

export type DocumentInteractionMode = "text" | "node" | "region";
export interface DocumentInteractionOptions extends DocumentPreviewOptions {
  readonly onSelectionChange?: (selection: DocumentSelection | undefined) => void;
  readonly onError?: (error: Error) => void;
}
export interface DocumentInteractionView {
  readonly element: HTMLElement;
  /** Resolves for the currently rendering revision. Read again after a content commit. */
  readonly ready: Promise<void>;
  selection(): DocumentSelection | undefined;
  fonts(): readonly FontResolution[];
  layout(): DocumentRenderReport | undefined;
  setMode(mode: DocumentInteractionMode): void;
  setZoom(zoom: number): void;
  focusAnnotation(annotationId: string): void;
  /** Releases only this view; the supplied session remains owned by the host. */
  dispose(): void;
}

function immutable<Value>(value: Value): Value {
  if (value && typeof value === "object") {
    for (const child of Object.values(value)) immutable(child);
    Object.freeze(value);
  }
  return value;
}
const statusLabels = { open: "待处理", resolved: "已处理", needs_reanchor: "需要重新定位" };
function anchorLabel(anchor: DeepReadonly<Anchor>, model: DocumentModel): string {
  if (anchor.kind === "text") return anchor.quote;
  if (anchor.kind === "text-set") return anchor.ranges.map((range) => range.quote).join(" … ");
  const nodeLabel = (id: string): string => {
    const node = model.getNode(id);
    if (!node) return "原对象";
    const text = node.tagName === "img" ? node.attributes.alt || "图片" : node.text.trim() || "文档对象";
    const characters = Array.from(text);
    return characters.slice(0, 80).join("") + (characters.length > 80 ? "…" : "");
  };
  if (anchor.kind === "node") return "对象：" + nodeLabel(anchor.nodeId);
  if (anchor.kind === "node-set") return "对象：" + anchor.nodeIds.map(nodeLabel).join("、");
  return anchor.nodeIds.length ? "正文区域 · 命中 " + anchor.nodeIds.length + " 个对象" : "正文空白区域";
}

/** Scriptless document preview plus revision-bound selection and annotation transactions. */
export function mountDocumentInteraction(
  container: HTMLElement,
  session: DocumentEditingSession,
  options: DocumentInteractionOptions = {},
): DocumentInteractionView {
  const document = container.ownerDocument;
  const element = document.createElement("div");
  const shadow = element.attachShadow({ mode: "open" });
  const style = document.createElement("style");
  style.textContent = `:host{display:block;color:#20252b;font:14px system-ui,sans-serif}*{box-sizing:border-box}button,select,textarea{font:inherit}button,select{padding:6px 10px}button{cursor:pointer}button:disabled{cursor:default}button[aria-pressed=true]{background:#e5efff;border-color:#245aa5}button:focus-visible,select:focus-visible,textarea:focus-visible{outline:2px solid #245aa5;outline-offset:2px}.toolbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px}.stage{overflow:auto;max-height:70vh;border:1px solid #d5dbe2}.panel{padding:12px 0}.composer{display:flex;flex-wrap:wrap;gap:8px;align-items:center}textarea{display:block;width:100%;min-height:64px;padding:8px}.selection{white-space:pre-wrap;overflow-wrap:anywhere}.card{padding:12px;margin-top:10px;border:1px solid #d5dbe2;border-radius:4px}.card p{margin:6px 0;white-space:pre-wrap;overflow-wrap:anywhere}.actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}.error{color:#a00000}.summary{margin:8px 0}h2{font-size:16px;margin:12px 0}`;
  const toolbar = document.createElement("div");
  toolbar.className = "toolbar";
  toolbar.setAttribute("role", "toolbar");
  toolbar.setAttribute("aria-label", "文档交互");
  const stage = document.createElement("div");
  stage.className = "stage";
  const plane = document.createElement("div");
  plane.style.position = "relative";
  stage.append(plane);
  const panel = document.createElement("section");
  panel.className = "panel";
  panel.setAttribute("aria-label", "文档批注");
  const heading = document.createElement("h2");
  heading.textContent = "批注";
  const summary = document.createElement("p");
  summary.className = "summary";
  summary.setAttribute("aria-label", "会话状态");
  const selectionText = document.createElement("p");
  selectionText.className = "selection";
  selectionText.setAttribute("aria-live", "polite");
  selectionText.setAttribute("aria-label", "当前选择");
  const draft = document.createElement("textarea");
  draft.setAttribute("aria-label", "新批注内容");
  draft.placeholder = "先选择文字、对象或区域，再填写批注";
  const composer = document.createElement("div");
  composer.className = "composer";
  const list = document.createElement("div");
  list.setAttribute("aria-label", "批注列表");
  const error = document.createElement("p");
  error.className = "error";
  error.setAttribute("role", "alert");
  error.hidden = true;
  panel.append(heading, summary, selectionText, draft, composer, list, error);
  shadow.append(style, toolbar, stage, panel);
  container.append(element);

  let disposed = false, mode: DocumentInteractionMode = "text", zoom = 1;
  let selection: DocumentSelection | undefined, activeAnnotation: string | undefined;
  let nativeDragging = false;
  let cancelPointerSelection = (): void => { nativeDragging = false; };
  let source: PreviewSource | undefined, layoutKey = "";
  let preview: ReturnType<typeof mountPreview> | undefined;
  let ready: Promise<void> = Promise.resolve();
  let frameCleanup: (() => void) | undefined;
  let generation = 0, updateQueued = false;
  let renderedRevision = -1, renderedContent = "";
  const controls = new AbortController();
  let cardControls = new AbortController();
  const editDrafts = new Map<string, { value: string; baseRevision: number }>();
  const modeButtons = new Map<DocumentInteractionMode, HTMLButtonElement>();

  function report(reason: unknown): void {
    if (disposed) return;
    const value = reason instanceof Error ? reason : new Error("文档交互失败。");
    error.hidden = false;
    error.textContent = value.message;
    options.onError?.(value);
  }
  function attempt(action: () => void): void {
    if (disposed) return;
    error.hidden = true;
    try { action(); } catch (reason) { report(reason); }
  }
  function button(label: string, action: () => void, parent: HTMLElement, signal = controls.signal): HTMLButtonElement {
    const control = document.createElement("button");
    control.type = "button";
    control.textContent = label;
    control.addEventListener("click", () => attempt(action), { signal });
    parent.append(control);
    return control;
  }
  function commit(operations: DocumentOperation[], baseRevision = session.snapshot().manifest.revision): void {
    const { documentId } = session.snapshot().manifest;
    const candidate = session.prepare({ protocolVersion: 1, documentId, baseRevision, operations });
    try { session.commit(candidate); } catch (reason) { session.discard(candidate); throw reason; }
  }
  function requireSelection(): DocumentSelection {
    const manifest = session.snapshot().manifest;
    if (!selection || selection.documentId !== manifest.documentId || selection.baseRevision !== manifest.revision)
      throw new DocumentError("RevisionConflict", "选择已过期，请在当前修订重新选择。");
    if (selection.anchor.kind === "region" && selection.anchor.layoutKey !== layoutKey)
      throw new DocumentError("RevisionConflict", "区域布局已变化，请重新框选。");
    return selection;
  }
  function setSelection(anchor: Anchor | undefined): void {
    const { documentId, revision } = session.snapshot().manifest;
    const next = anchor ? immutable({ documentId, baseRevision: revision, anchor }) : undefined;
    // Painting a DOM highlight can itself emit selectionchange during native dragging.
    if (selectionText.textContent && JSON.stringify(selection) === JSON.stringify(next)) return;
    selection = next;
    selectionText.textContent = selection ? anchorLabel(selection.anchor, session) : "尚未选择";
    add.disabled = !selection;
    clear.disabled = !selection;
    options.onSelectionChange?.(selection);
    highlight();
  }
  function clearSelection(): void {
    cancelPointerSelection();
    preview?.element.contentDocument?.getSelection()?.removeAllRanges();
    setSelection(undefined);
  }
  function setMode(value: DocumentInteractionMode): void {
    if (disposed) return;
    if (!["text", "node", "region"].includes(value)) throw new DocumentError("InvalidOperation", "未知选择模式。");
    mode = value;
    for (const [key, control] of modeButtons) control.setAttribute("aria-pressed", String(key === value));
    if (source) source.root.style.userSelect = value === "text" ? "text" : "none";
    clearSelection();
  }
  for (const [key, label] of [["text", "文字划选"], ["node", "对象选择"], ["region", "矩形框选"]] as const)
    modeButtons.set(key, button(label, () => setMode(key), toolbar));
  for (const [key, control] of modeButtons) control.setAttribute("aria-pressed", String(key === "text"));
  modeButtons.get("node")!.title = "Ctrl / Command 加点选可增减对象";
  const zoomSelect = document.createElement("select");
  zoomSelect.setAttribute("aria-label", "预览缩放");
  for (const value of [0.5, 0.75, 1, 1.25, 1.5, 2]) {
    const option = document.createElement("option");
    option.value = String(value);
    option.textContent = value * 100 + "%";
    zoomSelect.append(option);
  }
  zoomSelect.value = "1";
  zoomSelect.addEventListener("change", () => attempt(() => setZoom(Number(zoomSelect.value))), { signal: controls.signal });
  toolbar.append(zoomSelect);
  const undo = button("撤销", () => { session.undo(); }, toolbar);
  const redo = button("重做", () => { session.redo(); }, toolbar);
  const add = button("添加批注", () => {
    const chosen = requireSelection();
    if (!draft.value.trim()) throw new DocumentError("InvalidOperation", "请填写批注内容。");
    commit([{ op: "addAnnotation", annotation: {
      annotationId: "annotation-" + createOpaqueId(), documentId: chosen.documentId, baseRevision: chosen.baseRevision,
      anchor: structuredClone(chosen.anchor) as Anchor, instruction: draft.value, status: "open",
    } }], chosen.baseRevision);
    draft.value = "";
  }, composer);
  const clear = button("清除选择", clearSelection, composer);

  function updatePanel(): void {
    const snapshot = session.snapshot(), history = session.history();
    const focused = shadow.activeElement as HTMLTextAreaElement | null;
    const focusedId = focused?.dataset.editAnnotation;
    const caret = focusedId ? [focused!.selectionStart, focused!.selectionEnd] as const : undefined;
    summary.textContent = "修订 " + snapshot.manifest.revision + " · 批注 " + snapshot.annotations.length;
    undo.disabled = history.undoSteps === 0;
    redo.disabled = history.redoSteps === 0;
    cardControls.abort();
    cardControls = new AbortController();
    const cardButton = (label: string, action: () => void, parent: HTMLElement): HTMLButtonElement => button(label, action, parent, cardControls.signal);
    list.replaceChildren();
    for (const id of editDrafts.keys()) if (!snapshot.annotations.some((annotation) => annotation.annotationId === id)) editDrafts.delete(id);
    for (const annotation of snapshot.annotations) {
      const card = document.createElement("div");
      card.className = "card";
      card.dataset.annotationId = annotation.annotationId;
      const state = document.createElement("p");
      state.textContent = statusLabels[annotation.status];
      const quote = document.createElement("p");
      quote.textContent = anchorLabel(annotation.anchor, session);
      const content = document.createElement("p");
      content.textContent = annotation.instruction ?? session.getNode(annotation.bodyNodeId!)?.text ?? "";
      card.append(state, quote, content);
      const actions = document.createElement("div");
      actions.className = "actions";
      const editing = editDrafts.get(annotation.annotationId);
      if (editing) {
        const input = document.createElement("textarea");
        input.value = editing.value;
        input.dataset.editAnnotation = annotation.annotationId;
        input.setAttribute("aria-label", "修改批注内容");
        input.addEventListener("input", () => { editing.value = input.value; }, { signal: cardControls.signal });
        content.replaceWith(input);
        cardButton("保存批注", () => {
          if (!input.value.trim()) throw new DocumentError("InvalidOperation", "请填写批注内容。");
          commit([{ op: "updateAnnotation", annotationId: annotation.annotationId, changes: { instruction: input.value } }], editing.baseRevision);
          editDrafts.delete(annotation.annotationId);
          updatePanel();
        }, actions);
        cardButton("取消修改", () => { editDrafts.delete(annotation.annotationId); updatePanel(); }, actions);
        card.append(actions);
        list.append(card);
        if (focusedId === annotation.annotationId) { input.focus(); if (caret) input.setSelectionRange(...caret); }
        continue;
      }
      cardButton("定位", () => focusAnnotation(annotation.annotationId), actions);
      if (annotation.instruction !== undefined) {
        cardButton("修改批注", () => {
          editDrafts.set(annotation.annotationId, { value: annotation.instruction!, baseRevision: session.snapshot().manifest.revision });
          updatePanel();
          [...list.children].find((node) => (node as HTMLElement).dataset.annotationId === annotation.annotationId)?.querySelector("textarea")?.focus();
        }, actions);
      }
      if (annotation.status !== "needs_reanchor") cardButton(annotation.status === "open" ? "标为已处理" : "重新打开", () => {
        commit([{ op: "updateAnnotation", annotationId: annotation.annotationId, changes: { status: annotation.status === "open" ? "resolved" : "open" } }]);
      }, actions);
      cardButton("用当前选择重新定位", () => {
        const chosen = requireSelection();
        commit([{ op: "updateAnnotation", annotationId: annotation.annotationId,
          changes: { anchor: structuredClone(chosen.anchor) as Anchor, status: annotation.status === "resolved" ? "resolved" : "open" } }], chosen.baseRevision);
      }, actions);
      cardButton("删除批注", () => commit([{ op: "removeAnnotation", annotationId: annotation.annotationId }]), actions);
      card.append(actions);
      list.append(card);
    }
  }

  function highlight(): void {
    const frame = preview?.element.contentDocument;
    if (!frame || !source || nativeDragging) return;
    frame.querySelectorAll("[data-sdk-highlight]").forEach((node) => node.remove());
    const current = activeAnnotation ? session.getAnnotations().find((item) => item.annotationId === activeAnnotation) : undefined;
    const anchors = [selection?.anchor, current?.status !== "needs_reanchor" ? current?.anchor : undefined];
    for (const anchor of anchors) {
      if (!anchor) continue;
      if (anchor === selection?.anchor && (anchor.kind === "text" || anchor.kind === "text-set")
        && frame.hasFocus() && frame.getSelection()?.isCollapsed === false) continue;
      if (anchor.kind === "region" && anchor.layoutKey !== layoutKey) continue;
      for (const rect of source.rectangles(anchor)) {
        if (!rect.width || !rect.height) continue;
        const box = frame.createElement("div");
        box.dataset.sdkHighlight = "";
        box.setAttribute("aria-hidden", "true");
        box.style.cssText = `all:initial!important;position:fixed!important;pointer-events:none!important;box-sizing:border-box!important;z-index:2147483647!important;border:2px solid #245aa5!important;background:rgba(36,90,165,.12)!important;left:${rect.left}px!important;top:${rect.top}px!important;width:${rect.width}px!important;height:${rect.height}px!important;`;
        frame.documentElement.append(box);
      }
    }
  }
  function focusAnnotation(annotationId: string): void {
    if (disposed) return;
    const annotation = session.getAnnotations().find((item) => item.annotationId === annotationId);
    if (!annotation || annotation.status === "needs_reanchor" || annotation.anchor.kind === "region" && annotation.anchor.layoutKey !== layoutKey)
      throw new DocumentError("AnchorNotFound", "批注需要重新定位，请选择新位置后重新定位。");
    const rect = source?.rectangles(annotation.anchor)[0];
    if (!rect) throw new DocumentError("AnchorNotFound", "当前预览中无法定位该批注。");
    activeAnnotation = annotationId;
    const stageBox = stage.getBoundingClientRect();
    stage.scrollTop += preview!.element.getBoundingClientRect().top + rect.top * zoom - stageBox.top - 40;
    stage.scrollLeft += preview!.element.getBoundingClientRect().left + rect.left * zoom - stageBox.left - 40;
    highlight();
  }
  function setZoom(value: number): void {
    if (disposed) return;
    if (!Number.isFinite(value) || value < 0.5 || value > 2) throw new DocumentError("InvalidOperation", "缩放范围为 50% 至 200%。");
    zoom = value;
    zoomSelect.querySelector("[data-custom-zoom]")?.remove();
    if (![...zoomSelect.options].some((option) => option.value === String(value))) {
      const option = document.createElement("option");
      option.dataset.customZoom = "";
      option.value = String(value);
      option.textContent = Math.round(value * 100) + "%";
      zoomSelect.append(option);
    }
    zoomSelect.value = String(value);
    resizePlane();
    highlight();
  }
  function resizePlane(): void {
    if (!preview) return;
    // Transform the entire iframe: its viewport, typography and source layout stay unscaled.
    preview.element.style.width = stage.clientWidth + "px";
    preview.element.style.transformOrigin = "0 0";
    preview.element.style.transform = `scale(${zoom})`;
    plane.style.width = preview.element.offsetWidth * zoom + "px";
    plane.style.height = preview.element.offsetHeight * zoom + "px";
  }

  function invalidateRegions(): void {
    if (!source || disposed) return;
    const next = source.layoutKey();
    layoutKey = next;
    if (selection?.anchor.kind === "region" && selection.anchor.layoutKey !== next) setSelection(undefined);
    const outdated = session.getAnnotations().filter((annotation) => annotation.status !== "needs_reanchor"
      && annotation.anchor.kind === "region" && annotation.anchor.layoutKey !== next);
    if (outdated.length) commit(outdated.map((annotation) => ({ op: "updateAnnotation" as const,
      annotationId: annotation.annotationId, changes: { status: "needs_reanchor" as const } })));
    highlight();
  }

  function connect(): void {
    const frame = preview!.element.contentDocument!, view = frame.defaultView!;
    source = preview!.source()!;
    layoutKey = source.layoutKey();
    source.root.style.userSelect = mode === "text" ? "text" : "none";
    const listeners = new AbortController();
    const signal = listeners.signal;
    let drag: { id: number; start: { x: number; y: number } } | undefined;
    let layoutQueued = false;
    const layout = (): void => {
      if (layoutQueued || disposed) return;
      layoutQueued = true;
      queueMicrotask(() => {
        layoutQueued = false;
        if (!signal.aborted) attempt(invalidateRegions);
      });
    };
    const selectionChanged = (): void => {
      if (mode !== "text") return;
      const selected = view.getSelection();
      // Toolbar and panel focus must not erase the captured revision-bound selection.
      if (!selected?.rangeCount || selected.isCollapsed) {
        if (frame.hasFocus()) setSelection(undefined);
        return;
      }
      attempt(() => {
        try { setSelection(source!.textAnchor(selected.getRangeAt(0), session)); }
        catch (reason) { setSelection(undefined); throw reason; }
      });
    };
    frame.addEventListener("selectionchange", selectionChanged, { signal });
    frame.addEventListener("pointerup", () => {
      if (mode === "text") { nativeDragging = false; selectionChanged(); highlight(); }
    }, { signal });
    view.addEventListener("blur", () => { nativeDragging = false; highlight(); }, { signal });
    frame.addEventListener("click", (event) => {
      if (mode !== "node") return;
      if ((event.target as Node).nodeType !== 1) return;
      const target = (event.target as Element).closest<HTMLElement>("[data-node-id]");
      if (!target || !source!.contains(target)) return;
      event.preventDefault();
      const id = target.dataset.nodeId!;
      const previous = selection?.anchor;
      let ids = event.ctrlKey || event.metaKey ? previous?.kind === "node" ? [previous.nodeId]
        : previous?.kind === "node-set" ? [...previous.nodeIds] : [] : [];
      ids = ids.includes(id) ? ids.filter((value) => value !== id) : [...ids, id];
      ids = source!.orderedIds(ids);
      setSelection(ids.length === 1 ? { kind: "node", nodeId: ids[0]! } : ids.length ? { kind: "node-set", nodeIds: ids } : undefined);
    }, { signal });
    frame.addEventListener("pointerdown", (event) => {
      if (mode === "text") {
        activeAnnotation = undefined;
        highlight();
        nativeDragging = true;
        return;
      }
      if (mode !== "region" || event.button !== 0 || !source!.contains(event.target as Node)) return;
      event.preventDefault();
      frame.getSelection()?.removeAllRanges();
      drag = { id: event.pointerId, start: source!.point(event.clientX, event.clientY) };
      source!.root.setPointerCapture(event.pointerId);
      setSelection(undefined);
    }, { signal });
    const move = (event: PointerEvent): void => {
      if (!drag || event.pointerId !== drag.id) return;
      const end = source!.point(event.clientX, event.clientY), start = drag.start;
      const rect: Rectangle = { x: Math.min(start.x, end.x), y: Math.min(start.y, end.y), width: Math.abs(end.x - start.x), height: Math.abs(end.y - start.y) };
      if (rect.width > 1 && rect.height > 1) setSelection(source!.region(rect, layoutKey));
    };
    frame.addEventListener("pointermove", move, { signal });
    frame.addEventListener("pointerup", (event) => {
      if (!drag || event.pointerId !== drag.id) return;
      move(event);
      drag = undefined;
      source!.root.releasePointerCapture(event.pointerId);
    }, { signal });
    frame.addEventListener("pointercancel", () => { drag = undefined; nativeDragging = false; setSelection(undefined); }, { signal });
    source.root.addEventListener("lostpointercapture", () => {
      if (drag) { drag = undefined; setSelection(undefined); }
    }, { signal });
    cancelPointerSelection = () => {
      nativeDragging = false;
      const pointer = drag?.id;
      drag = undefined;
      if (pointer !== undefined && source?.root.hasPointerCapture(pointer)) source.root.releasePointerCapture(pointer);
    };
    frame.addEventListener("keydown", keydown, { signal });
    view.addEventListener("scroll", highlight, { signal });
    view.addEventListener("resize", layout, { signal });
    frame.fonts.addEventListener("loadingdone", layout, { signal });
    let resizeQueued: number | undefined;
    const observer = new ResizeObserver(() => {
      if (resizeQueued !== undefined || signal.aborted) return;
      resizeQueued = document.defaultView!.requestAnimationFrame(() => {
        resizeQueued = undefined;
        if (!signal.aborted) { resizePlane(); layout(); }
      });
    });
    observer.observe(source.root);
    observer.observe(preview!.element);
    observer.observe(stage);
    frameCleanup = () => {
      listeners.abort(); observer.disconnect();
      if (resizeQueued !== undefined) document.defaultView!.cancelAnimationFrame(resizeQueued);
    };
    setZoom(zoom);
    attempt(invalidateRegions);
  }
  function keydown(event: KeyboardEvent): void {
    const target = event.composedPath()[0] as HTMLElement;
    if (target?.matches?.("textarea,input,select")) return;
    if (event.key === "Escape") { clearSelection(); return; }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
      event.preventDefault();
      attempt(() => { if (event.shiftKey) session.redo(); else session.undo(); });
    }
  }
  shadow.addEventListener("keydown", (event) => keydown(event as KeyboardEvent), { signal: controls.signal });
  document.addEventListener("pointerup", () => { nativeDragging = false; highlight(); }, { signal: controls.signal });

  function contentKey(): string {
    const snapshot = session.snapshot();
    return JSON.stringify([snapshot.html, snapshot.styles, snapshot.metadata, snapshot.manifest.assets]);
  }
  function render(): void {
    const current = ++generation;
    frameCleanup?.();
    frameCleanup = undefined;
    preview?.dispose();
    source = undefined;
    nativeDragging = false;
    plane.replaceChildren();
    plane.style.width = "100%";
    plane.style.height = "auto";
    renderedRevision = session.snapshot().manifest.revision;
    renderedContent = contentKey();
    preview = mountPreview(plane, session, options);
    ready = preview.ready.then(() => {
      if (disposed || generation !== current) return;
      connect();
    });
    void ready.catch(report);
  }
  const unsubscribe = session.subscribe((event) => {
    if (event.type === "error" || updateQueued || disposed) return;
    updateQueued = true;
    queueMicrotask(() => {
      updateQueued = false;
      if (disposed) return;
      attempt(() => {
        if (session.snapshot().manifest.revision !== renderedRevision) {
          setSelection(undefined);
          renderedRevision = session.snapshot().manifest.revision;
          if (renderedContent !== contentKey()) render();
        }
        updatePanel();
        highlight();
      });
    });
  });
  setSelection(undefined);
  updatePanel();
  render();
  return {
    element, get ready() { return ready; }, selection: () => selection, fonts: () => preview?.fonts() ?? [], layout: () => preview?.layout(),
    setMode, setZoom, focusAnnotation,
    dispose: () => {
      if (disposed) return;
      disposed = true;
      generation++;
      unsubscribe();
      controls.abort();
      cardControls.abort();
      frameCleanup?.();
      frameCleanup = undefined;
      preview?.dispose();
      preview = undefined;
      cancelPointerSelection = () => {};
      selection = undefined;
      source = undefined;
      editDrafts.clear();
      list.replaceChildren();
      draft.value = "";
      element.remove();
    },
  };
}
