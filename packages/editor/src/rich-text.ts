import { DocumentError } from "@dagent-office/core";
import type { DocumentEditingSession, ReconciledTextEdit } from "@dagent-office/core";
import { decorate, mountDocumentPreview, packagePath } from "./preview.js";
import { resolvePreviewFonts } from "./fonts.js";
import { blockSelector, bookmark, caretAt, inputProofs, inputStep, offsetAt, restoreBookmark, sourceText } from "./rich-text/dom.js";
import type { InputStep, SelectionBookmark } from "./rich-text/dom.js";
import { SourceProjection } from "./rich-text/projection.js";

export interface DocumentTextEditorOptions { readonly onError?: (error: Error) => void }
export interface DocumentTextEditor {
  readonly element: HTMLElement;
  readonly ready: Promise<void>;
  focus(): void;
  /** Core owns persistent state and history. Disposal releases only this view. */
  dispose(): void;
}

/** Native input over the same scriptless HTML/CSS renderer used by preview. */
export function mountDocumentTextEditor(container: HTMLElement, session: DocumentEditingSession, options: DocumentTextEditorOptions = {}): DocumentTextEditor {
  const document = container.ownerDocument, element = document.createElement("div"), shadow = element.attachShadow({ mode: "open" });
  const style = document.createElement("style");
  style.textContent = ":host{display:block;font:14px system-ui}button,select{font:inherit;padding:6px 10px}.toolbar{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}.stage{max-height:70vh;overflow:auto;border:1px solid #ccd3dc}.error{color:#a00000}";
  const toolbar = document.createElement("div"); toolbar.className = "toolbar"; toolbar.setAttribute("role", "toolbar"); toolbar.setAttribute("aria-label", "正文格式");
  const stage = document.createElement("div"); stage.className = "stage";
  const error = document.createElement("p"); error.setAttribute("role", "alert"); error.className = "error"; error.hidden = true;
  shadow.append(style, toolbar, stage, error); container.append(element);
  const controls = new AbortController(), initial = session.snapshot();
  let viewControls = new AbortController();
  let disposed = false, committing = false, composing = false, compositionRevision: number | undefined;
  let root: HTMLElement | undefined, frame: Document | undefined, revision = session.snapshot().manifest.revision;
  let projection = new SourceProjection(), proofs: ReconciledTextEdit[] = [], pending: InputStep | undefined;
  let savedSelection: SelectionBookmark | undefined;
  let compositionInsertion: { from: number; to: number } | undefined;
  const typingStyles = new Map<string, string>();
  const layout = new Map<string, { property: string; original: string; rendered: string }[]>();
  const imageUrls = new Map<string, string>();
  const report = (reason: unknown): void => {
    if (disposed) return;
    const value = reason instanceof Error ? reason : new Error("正文输入失败。");
    error.textContent = value.message; error.hidden = false; options.onError?.(value);
  };
  const sourceDocument = (): Document => new DOMParser().parseFromString(session.snapshot().html, "text/html");
  const rootElement = (source: Document): HTMLElement => {
    const main = session.snapshot().metadata.word?.stories.find((story) => story.kind === "main");
    const result = main ? source.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(main.rootNodeId) + '"]')
      : source.querySelector<HTMLElement>('body > [data-office-kind="document"]');
    if (!result) throw new DocumentError("InvalidDocument", "正文内容域不存在。"); return result;
  };
  const selectedRange = (): Range | undefined => {
    const selection = frame?.getSelection();
    if (!selection?.rangeCount || !root) return;
    const range = selection.getRangeAt(0);
    return root.contains(range.startContainer) && root.contains(range.endContainer) ? range.cloneRange() : undefined;
  };
  const select = (node: Node, offset: number): void => {
    const range = frame!.createRange(); range.setStart(node, offset); range.collapse(true);
    const selection = frame!.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
  };
  function refresh(selection = savedSelection): void {
    if (!root || !frame || disposed) return;
    const source = rootElement(sourceDocument()), next = frame.importNode(source, true);
    const snapshot = session.snapshot(), files = session.serialize(), used = new Set<string>();
    const css = frame.querySelector("style[data-sdk-source-styles]");
    if (css) css.textContent = snapshot.styles.map((item) => item.css).join("\n");
    for (const node of [next, ...next.querySelectorAll<HTMLElement>("[data-node-id]")]) {
      for (const item of layout.get(node.dataset.nodeId!) ?? [])
        if (node.style.getPropertyValue(item.property) === item.original) node.style.setProperty(item.property, item.rendered);
      if (node.localName === "img") {
        const path = packagePath(node.getAttribute("src")!, snapshot.manifest.entrypoint), asset = snapshot.manifest.assets.find((item) => item.path === path)!;
        const key = path + ":" + asset.sha256; used.add(key);
        let url = imageUrls.get(key);
        if (!url) {
          const bytes = files.get(path)!, buffer = new ArrayBuffer(bytes.length); new Uint8Array(buffer).set(bytes);
          url = URL.createObjectURL(new Blob([buffer], { type: asset.mimeType })); imageUrls.set(key, url);
        }
        node.setAttribute("src", url);
      }
    }
    for (const [key, url] of imageUrls) if (!used.has(key)) { URL.revokeObjectURL(url); imageUrls.delete(key); }
    viewControls.abort(); viewControls = new AbortController();
    root.replaceWith(next); root = next; projection = new SourceProjection();
    resolvePreviewFonts(frame, "[data-node-id]"); decorate(frame, session.snapshot().metadata.word, root);
    activate(source); revision = snapshot.manifest.revision; proofs = []; pending = undefined; typingStyles.clear(); compositionInsertion = undefined;
    root.focus(); restoreBookmark(root, selection, projection.ignored); savedSelection = bookmark(root, projection.ignored);
  }
  function normalizedRoot(): HTMLElement {
    const value = projection.serialize(root!), source = sourceDocument(), original = rootElement(source);
    const sections = session.snapshot().metadata.sections.map((section) => source.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(section.rootNodeId) + '"]'))
      .filter((section): section is HTMLElement => Boolean(section && section !== original && original.contains(section)));
    for (const [index, section] of sections.entries()) {
      if (value.querySelector('[data-node-id="' + CSS.escape(section.dataset.nodeId!) + '"]')) continue;
      let reference: Element | null = null;
      for (const following of sections.slice(index + 1)) { reference = value.querySelector('[data-node-id="' + CSS.escape(following.dataset.nodeId!) + '"]'); if (reference) break; }
      while (reference?.parentElement && reference.parentElement !== value) reference = reference.parentElement;
      value.insertBefore(section.cloneNode(false), reference);
    }
    let target: Element = value;
    if (sections.length) {
      target = value.querySelector('[data-node-id="' + CSS.escape(sections[0]!.dataset.nodeId!) + '"]')!;
      for (const child of [...value.childNodes]) {
        if (child === target || child.nodeType === 1 && sections.some((section) => (child as Element).querySelector('[data-node-id="' + CSS.escape(section.dataset.nodeId!) + '"]') || (child as Element).getAttribute("data-node-id") === section.dataset.nodeId)) continue;
        if (child.nodeType !== 3 || child.textContent!.trim()) target.append(child);
      }
    }
    if ([...target.childNodes].some((child) => child.nodeType === 3 && child.textContent!.trim()) && !target.querySelector(blockSelector)) {
      const paragraph = document.createElement("p"); paragraph.append(...target.childNodes); target.append(paragraph);
    }
    const paragraph = original.querySelector("p"), run = paragraph?.querySelector('[data-word-run="true"]');
    for (const node of value.querySelectorAll("p:not([data-node-id])")) {
      if (paragraph && !node.hasAttribute("class") && !node.hasAttribute("style"))
        for (const name of ["class", "style", "data-word-style", "data-word-line-rule"]) { const attribute = paragraph.getAttribute(name); if (attribute) node.setAttribute(name, attribute); }
      if (run && !node.querySelector("[data-word-run]")) {
        const wrapper = document.createElement("span");
        for (const name of ["data-word-run", "class", "style"]) { const attribute = run.getAttribute(name); if (attribute) wrapper.setAttribute(name, attribute); }
        wrapper.append(...node.childNodes); node.append(wrapper);
      }
    }
    return value;
  }
  function flush(): void {
    if (!root || disposed || composing) return;
    const selection = bookmark(root, projection.ignored);
    try {
      const baseRevision = compositionRevision ?? revision;
      if (baseRevision !== session.snapshot().manifest.revision) throw new DocumentError("RevisionConflict", "输入期间文档已变化，请基于当前修订重新输入。");
      const normalized = normalizedRoot(), values = new Map<string, string>();
      for (const edit of proofs) {
        const previous = values.get(edit.nodeId) ?? session.getNode(edit.nodeId)!.text;
        values.set(edit.nodeId, previous.slice(0, edit.from) + edit.text + previous.slice(edit.to));
      }
      const changedByNormalization = new Set([...values].filter(([id, value]) => {
        const node = id === normalized.dataset.nodeId ? normalized : normalized.querySelector('[data-node-id="' + CSS.escape(id) + '"]');
        return node && sourceText(node, () => false) !== value;
      }).map(([id]) => id));
      const candidate = session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision,
        operations: [{ op: "reconcileFragment", nodeId: root.dataset.nodeId!, html: normalized.innerHTML, textEdits: proofs.filter((edit) => !changedByNormalization.has(edit.nodeId)) }] });
      committing = true;
      try { session.commit(candidate); } catch (reason) { session.discard(candidate); throw reason; }
      finally { committing = false; }
      error.hidden = true; compositionRevision = undefined; refresh(selection);
    } catch (reason) { compositionRevision = undefined; proofs = []; refresh(selection); report(reason); }
  }
  function mutate(range: Range, action: () => void, split = false): void {
    if (!root || composing || disposed) return;
    const step = inputStep(root, range, projection.ignored);
    const parent = range.startContainer.nodeType === 1 ? range.startContainer as Element : range.startContainer.parentElement;
    const block = parent?.closest(blockSelector);
    const moved = split && block && !block.matches("td,th,li")
      ? { ...step, to: Math.max(step.to, offsetAt(root, block, block.childNodes.length, projection.ignored)) } : undefined;
    try {
      action();
      const actual = inputProofs(root, step, projection.ignored), known = new Set(actual.map((edit) => edit.nodeId));
      proofs.push(...actual, ...(moved ? inputProofs(root, moved, projection.ignored).filter((edit) => !known.has(edit.nodeId)) : [])); flush();
    }
    catch (reason) { refresh(); report(reason); }
  }
  function insert(range: Range, content: Node): void {
    deleteRange(range); range.insertNode(content); select(content.parentNode!, [...content.parentNode!.childNodes].indexOf(content as ChildNode) + 1);
  }
  function deleteRange(range: Range): void {
    const blockAt = (node: Node): Element | null => (node.nodeType === 1 ? node as Element : node.parentElement)?.closest(blockSelector) ?? null;
    const first = blockAt(range.startContainer), last = blockAt(range.endContainer);
    const join = first && last && first !== last && first.parentElement === last.parentElement && !first.matches("td,th,li") && !last.matches("td,th,li");
    range.deleteContents();
    if (join) {
      const offset = first.childNodes.length;
      first.append(...last.childNodes); last.remove(); range.setStart(first, offset); range.collapse(true);
    }
  }
  function lineBreak(range: Range): void { insert(range, frame!.createElement("br")); }
  function paragraphBreak(range: Range): void {
    deleteRange(range);
    const parent = range.startContainer.nodeType === 1 ? range.startContainer as Element : range.startContainer.parentElement;
    const block = parent?.closest<HTMLElement>(blockSelector);
    if (!block || block === root || block.matches("td,th,li")) { lineBreak(range); return; }
    const tail = frame!.createRange(); tail.setStart(range.startContainer, range.startOffset); tail.setEnd(block, block.childNodes.length);
    const authored = projection.serialize(block), next = frame!.createElement(block.localName);
    for (const attribute of authored.attributes) if (!["id", "data-node-id"].includes(attribute.name)) next.setAttribute(attribute.name, attribute.value);
    next.append(tail.extractContents()); block.after(next); select(next, 0);
  }
  function format(property: string, value: string): void {
    if (!root || composing) return;
    root.focus(); if (!selectedRange()) restoreBookmark(root, savedSelection, projection.ignored);
    const range = selectedRange(); if (!range) return;
    if (range.collapsed) {
      typingStyles.set(property, value); return;
    }
    mutate(range, () => styleRange(range, new Map([[property, value]])));
  }
  function styleRange(range: Range, properties: ReadonlyMap<string, string>): void {
    if (!root) return;
    const walker = frame!.createTreeWalker(root!, NodeFilter.SHOW_TEXT), ranges: Range[] = [];
    let node: Node | null;
    while ((node = walker.nextNode())) {
      if (projection.ignored(node.parentNode!) || !range.intersectsNode(node)) continue;
      const part = frame!.createRange(); part.selectNodeContents(node);
      if (range.compareBoundaryPoints(Range.START_TO_START, part) > 0) part.setStart(range.startContainer, range.startOffset);
      if (range.compareBoundaryPoints(Range.END_TO_END, part) < 0) part.setEnd(range.endContainer, range.endOffset);
      if (!part.collapsed) ranges.push(part);
    }
    for (const part of ranges.reverse()) { const span = frame!.createElement("span"); for (const [property, value] of properties) span.style.setProperty(property, value); part.surroundContents(span); }
  }
  function button(label: string, action: () => void): void {
    const control = document.createElement("button"); control.type = "button"; control.textContent = label;
    control.addEventListener("mousedown", (event) => event.preventDefault(), { signal: controls.signal });
    control.addEventListener("click", () => { try { action(); } catch (reason) { report(reason); } }, { signal: controls.signal }); toolbar.append(control);
  }
  button("加粗", () => format("font-weight", "700")); button("常规字重", () => format("font-weight", "400"));
  button("斜体", () => format("font-style", "italic")); button("下划线", () => format("text-decoration", "underline")); button("深蓝色", () => format("color", "#1f4b99"));
  const size = document.createElement("select"); size.setAttribute("aria-label", "字号");
  for (const value of [9, 10, 11, 12, 14, 16, 18, 24]) { const option = document.createElement("option"); option.value = value + "pt"; option.textContent = value + " 磅"; size.append(option); }
  size.value = "12pt"; size.addEventListener("change", () => format("font-size", size.value), { signal: controls.signal }); toolbar.append(size);
  for (const [label, alignment] of [["左对齐", "left"], ["居中", "center"], ["右对齐", "right"]] as const) button(label, () => {
    if (!root || composing) return;
    const range = selectedRange(); if (!range) return;
    const nodes = [...root.querySelectorAll<HTMLElement>(blockSelector)].filter((node) => range.intersectsNode(node) && node.dataset.nodeId);
    if (!nodes.length) return;
    savedSelection = bookmark(root, projection.ignored);
    const candidate = session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: revision,
      operations: nodes.map((node) => ({ op: "setStyle" as const, nodeId: node.dataset.nodeId!, property: "text-align", value: alignment })) });
    try { session.commit(candidate); } catch (reason) { session.discard(candidate); throw reason; }
  });
  button("撤销输入", () => { session.undo(); }); button("重做输入", () => { session.redo(); });

  function paste(event: ClipboardEvent): void {
    event.preventDefault(); if (!event.clipboardData || composing || !root) return;
    try {
      const raw = event.clipboardData.getData("text/html"), text = event.clipboardData.getData("text/plain");
      if (raw.length > 1024 * 1024 || text.length > 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "粘贴内容超过大小限制。");
      const template = document.createElement("template");
      if (raw) template.innerHTML = raw;
      else text.replace(/\r\n?/g, "\n").split("\n").forEach((line, index) => { if (index) template.content.append(document.createElement("br")); template.content.append(document.createTextNode(line)); });
      for (const node of template.content.querySelectorAll("[data-node-id],[id]")) { node.removeAttribute("data-node-id"); node.removeAttribute("id"); }
      const block = [...template.content.children].some((node) => /^(?:p|h[1-6]|div|section|article|table|ul|ol)$/.test(node.localName));
      const validation = session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: revision,
        operations: [{ op: "insertFragment", nodeId: root.dataset.nodeId!, position: "end", html: block ? template.innerHTML : "<p>" + template.innerHTML + "</p>" }] });
      session.discard(validation);
      const range = selectedRange(); if (!range) return;
      mutate(range, () => {
        if (block) {
          paragraphBreak(range);
          const current = selectedRange()!, parent = current.startContainer.nodeType === 1 ? current.startContainer as Element : current.startContainer.parentElement!;
          const paragraph = parent.closest(blockSelector);
          if (!paragraph) throw new DocumentError("UnsupportedFeature", "当前位置不支持区块粘贴。");
          const last = template.content.lastChild!; paragraph.before(template.content); select(last, last.childNodes.length);
        } else {
          const last = template.content.lastChild; deleteRange(range); range.insertNode(template.content);
          if (last) select(last.parentNode!, [...last.parentNode!.childNodes].indexOf(last as ChildNode) + 1);
        }
      });
    } catch (reason) { report(reason); }
  }
  function copy(event: ClipboardEvent, cut: boolean): void {
    if (!root || !frame || !event.clipboardData) return;
    event.preventDefault(); if (composing) return;
    const range = selectedRange(); if (!range || range.collapsed) return;
    const holder = frame.createElement("div");
    let fragment: Node = range.cloneContents(), ancestor = range.commonAncestorContainer;
    if (ancestor.nodeType !== 1) ancestor = ancestor.parentNode!;
    while (ancestor !== root && root.contains(ancestor)) {
      if ((ancestor as Element).matches("td,th")) break;
      const wrapper = ancestor.cloneNode(false); wrapper.appendChild(fragment); fragment = wrapper;
      if ((ancestor as Element).matches(blockSelector)) {
        if ((ancestor as Element).matches("li") && ancestor.parentElement?.matches("ul,ol")) {
          const list = ancestor.parentElement.cloneNode(false); list.appendChild(fragment); fragment = list;
        }
        break;
      }
      ancestor = ancestor.parentNode!;
    }
    holder.append(fragment);
    const source = projection.serialize(holder);
    for (const node of source.querySelectorAll("[data-node-id],[id]")) { node.removeAttribute("data-node-id"); node.removeAttribute("id"); }
    event.clipboardData.setData("text/html", source.innerHTML);
    event.clipboardData.setData("text/plain", [...source.childNodes].map((node) => sourceText(node, () => false)).join("\n"));
    if (cut) mutate(range, () => { deleteRange(range); select(range.startContainer, range.startOffset); });
  }
  function keydown(event: KeyboardEvent): void {
    if (!root || composing) return;
    const key = event.key.toLowerCase();
    if ((event.ctrlKey || event.metaKey) && ["z", "b", "i", "u"].includes(key)) {
      event.preventDefault();
      try { if (key === "z") { if (event.shiftKey) session.redo(); else session.undo(); }
        else format(key === "b" ? "font-weight" : key === "i" ? "font-style" : "text-decoration", key === "b" ? "700" : key === "i" ? "italic" : "underline"); }
      catch (reason) { report(reason); } return;
    }
    const range = selectedRange(); if (!range) return;
    if (event.key === "Enter") { event.preventDefault(); mutate(range, () => event.shiftKey ? lineBreak(range) : paragraphBreak(range), !event.shiftKey); }
    if (event.key === "Tab") {
      const parent = range.startContainer.nodeType === 1 ? range.startContainer as Element : range.startContainer.parentElement;
      const cell = parent?.closest("td,th"), cells = [...root.querySelectorAll("td,th")];
      if (cell && cells.includes(cell)) { event.preventDefault(); const next = cells[cells.indexOf(cell) + (event.shiftKey ? -1 : 1)]; if (next) select(...caretAt(next, 0, projection.ignored)); }
    }
  }
  function beforeinput(event: InputEvent): void {
    if (!root) return;
    let range = selectedRange();
    const target = event.getTargetRanges()[0];
    if (target) { range = frame!.createRange(); range.setStart(target.startContainer, target.startOffset); range.setEnd(target.endContainer, target.endOffset); }
    if (!range || !root.contains(range.commonAncestorContainer)) { event.preventDefault(); return; }
    if (composing || event.isComposing || event.inputType === "insertCompositionText") { pending = inputStep(root, range, projection.ignored); return; }
    event.preventDefault();
    if (event.inputType === "historyUndo" || event.inputType === "historyRedo") {
      try { if (event.inputType === "historyUndo") session.undo(); else session.redo(); } catch (reason) { report(reason); } return;
    }
    if (["insertText", "insertReplacementText"].includes(event.inputType)) mutate(range, () => {
      let content: Node = frame!.createTextNode(event.data ?? "");
      if (typingStyles.size) { const span = frame!.createElement("span"); for (const [property, value] of typingStyles) span.style.setProperty(property, value); span.append(content); content = span; }
      insert(range!, content);
    });
    else if (event.inputType.startsWith("delete")) mutate(range, () => { deleteRange(range!); select(range!.startContainer, range!.startOffset); });
    else if (["insertParagraph", "insertLineBreak"].includes(event.inputType)) mutate(range, () => event.inputType === "insertParagraph" ? paragraphBreak(range!) : lineBreak(range!), event.inputType === "insertParagraph");
    else report(new DocumentError("UnsupportedFeature", "当前不支持此输入操作：" + event.inputType));
  }
  function activate(source: HTMLElement): void {
    if (!root || !frame) return;
    root.contentEditable = "true"; root.setAttribute("data-office-editor", ""); root.setAttribute("role", "textbox"); root.setAttribute("aria-label", "文档正文"); root.setAttribute("aria-multiline", "true"); root.spellcheck = false;
    for (const marker of root.querySelectorAll<HTMLElement>(".sdk-number-marker,[data-word-tab]")) marker.contentEditable = "false";
    projection.capture(root, source);
    if (!root.querySelector(blockSelector)) {
      const paragraph = frame.createElement("p"); paragraph.append(...root.childNodes); root.append(paragraph);
      if (!paragraph.childNodes.length) { const filler = frame.createElement("br"); projection.markDerived(filler); paragraph.append(filler); }
    }
    root.addEventListener("beforeinput", beforeinput as EventListener, { signal: viewControls.signal });
    root.addEventListener("input", () => {
      if (!root || disposed) return;
      if (pending) {
        const edits = inputProofs(root, pending, projection.ignored); proofs.push(...edits);
        const rootEdit = edits.find((edit) => edit.nodeId === root!.dataset.nodeId);
        if (rootEdit) compositionInsertion = { from: rootEdit.from, to: rootEdit.from + rootEdit.text.length };
      }
      pending = undefined;
      if (!composing) flush();
    }, { signal: viewControls.signal });
    root.addEventListener("keydown", keydown, { signal: viewControls.signal }); root.addEventListener("paste", paste as EventListener, { signal: viewControls.signal });
    root.addEventListener("copy", ((event: ClipboardEvent) => copy(event, false)) as EventListener, { signal: viewControls.signal });
    root.addEventListener("cut", ((event: ClipboardEvent) => copy(event, true)) as EventListener, { signal: viewControls.signal });
    root.addEventListener("drop", (event) => { event.preventDefault(); report(new DocumentError("UnsupportedFeature", "请通过粘贴或宿主资源接口插入内容。")); }, { signal: viewControls.signal });
    root.addEventListener("compositionstart", () => { composing = true; compositionRevision = revision; }, { signal: viewControls.signal });
    root.addEventListener("compositionend", () => { queueMicrotask(() => {
      if (disposed || !root || !frame) return;
      composing = false;
      if (compositionInsertion && typingStyles.size && revision === session.snapshot().manifest.revision) {
        const selected = bookmark(root, projection.ignored), range = frame.createRange();
        range.setStart(...caretAt(root, compositionInsertion.from, projection.ignored)); range.setEnd(...caretAt(root, compositionInsertion.to, projection.ignored));
        styleRange(range, typingStyles); restoreBookmark(root, selected, projection.ignored);
      }
      flush();
    }); }, { signal: viewControls.signal });
  }
  const unsubscribe = session.subscribe((event) => { if (event.type !== "error" && !committing && !disposed && !composing) refresh(); });
  const preview = mountDocumentPreview(stage, session, { title: "正文编辑" });
  const ready = preview.ready.then(() => {
    if (disposed) return;
    frame = preview.element.contentDocument!; root = rootElement(frame);
    const source = rootElement(new DOMParser().parseFromString(initial.html, "text/html"));
    for (const node of [root, ...root.querySelectorAll<HTMLElement>("[data-node-id]")]) {
      const original = node.dataset.nodeId === source.dataset.nodeId ? source : source.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(node.dataset.nodeId!) + '"]');
      if (!original) continue;
      if (session.snapshot().metadata.sections.some((section) => section.rootNodeId === node.dataset.nodeId))
        layout.set(node.dataset.nodeId!, [...node.style].filter((property) => node.style.getPropertyValue(property) !== original.style.getPropertyValue(property))
          .map((property) => ({ property, original: original.style.getPropertyValue(property), rendered: node.style.getPropertyValue(property) })));
    }
    const viewStyle = frame.createElement("style"); viewStyle.textContent = "[data-office-editor]:focus{outline:none}[data-office-editor]{min-height:100px}[data-office-editor] [data-word-tab]{user-select:all}"; frame.head.append(viewStyle);
    activate(source);
    if (revision !== session.snapshot().manifest.revision) refresh();
    frame.addEventListener("selectionchange", () => { if (root) savedSelection = bookmark(root, projection.ignored) ?? savedSelection; }, { signal: controls.signal });
  }).catch((reason: unknown) => { report(reason); throw reason; });
  return { element, ready, focus: () => { root?.focus(); }, dispose: () => {
    if (disposed) return; disposed = true; unsubscribe(); controls.abort(); viewControls.abort(); preview.dispose();
    for (const url of imageUrls.values()) URL.revokeObjectURL(url); imageUrls.clear();
    root = undefined; frame = undefined; proofs = []; pending = undefined; savedSelection = undefined; element.remove();
  } };
}
