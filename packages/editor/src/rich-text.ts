import { Editor, Extension } from "@tiptap/core";
import type { JSONContent } from "@tiptap/core";
import { baseKeymap } from "@tiptap/pm/commands";
import { keymap } from "@tiptap/pm/keymap";
import { TextSelection } from "@tiptap/pm/state";
import type { Transaction } from "@tiptap/pm/state";
import { DocumentError } from "@dagent-office/core";
import type { DocumentEditingSession, ReconciledTextEdit } from "@dagent-office/core";
import { mountDocumentPreview } from "./preview.js";
import type { DocumentPreview } from "./preview.js";
import { attributes, canonicalOffset, editorExtensions, sourceHtml, sourceJson } from "./rich-text/source.js";
import { positionAtOffset, transactionProofs } from "./rich-text/mapping.js";

export interface DocumentTextEditorOptions { readonly onError?: (error: Error) => void }
export interface DocumentTextEditor {
  readonly element: HTMLElement;
  readonly ready: Promise<void>;
  focus(): void;
  /** Core owns all persistent state and history; disposal only releases this view. */
  dispose(): void;
}

export function mountDocumentTextEditor(container: HTMLElement, session: DocumentEditingSession, options: DocumentTextEditorOptions = {}): DocumentTextEditor {
  const document = container.ownerDocument, element = document.createElement("div"), shadow = element.attachShadow({ mode: "open" });
  const style = document.createElement("style"); style.textContent = ":host{display:block;font:14px system-ui}button,select{font:inherit;padding:6px 10px}.toolbar{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}.stage{max-height:70vh;overflow:auto;border:1px solid #ccd3dc}.error{color:#a00000}";
  const toolbar = document.createElement("div"); toolbar.className = "toolbar"; toolbar.setAttribute("role", "toolbar"); toolbar.setAttribute("aria-label", "正文格式");
  const stage = document.createElement("div"); stage.className = "stage";
  const error = document.createElement("p"); error.setAttribute("role", "alert"); error.className = "error"; error.hidden = true;
  shadow.append(style, toolbar, stage, error); container.append(element);
  const controls = new AbortController();
  let disposed = false, suppressed = false, committing = false, composing = false, compositionRevision: number | undefined;
  let editor: Editor | undefined, preview: DocumentPreview | undefined, sourceRootId = "", revision = session.snapshot().manifest.revision;
  let proofs: ReconciledTextEdit[] = [];
  let renderDocument: Document | undefined;
  const report = (reason: unknown): void => {
    if (disposed) return;
    const value = reason instanceof Error ? reason : new Error("正文输入失败。"); error.textContent = value.message; error.hidden = false; options.onError?.(value);
  };
  const sourceDocument = (): Document => new DOMParser().parseFromString(session.snapshot().html, "text/html");
  const rootElement = (source: Document): HTMLElement => {
    const main = session.snapshot().metadata.word?.stories.find((story) => story.kind === "main");
    const root = main ? source.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(main.rootNodeId) + '"]')
      : source.querySelector<HTMLElement>('body > [data-office-kind="document"]');
    if (!root) throw new DocumentError("InvalidDocument", "正文内容域不存在。"); return root;
  };
  function content(): JSONContent {
    renderDocument = sourceDocument(); const root = rootElement(renderDocument); sourceRootId = root.dataset.nodeId!;
    return { type: "doc", content: [sourceJson(root)] };
  }
  function refresh(anchor?: number, head?: number): void {
    if (!editor || disposed) return;
    suppressed = true;
    try {
      editor.commands.setContent(content(), { emitUpdate: false });
      if (anchor !== undefined) {
        const from = positionAtOffset(editor.state.doc, anchor), to = positionAtOffset(editor.state.doc, head ?? anchor);
        editor.view.dispatch(editor.state.tr.setSelection(TextSelection.between(editor.state.doc.resolve(from), editor.state.doc.resolve(to))));
      }
      revision = session.snapshot().manifest.revision; proofs = [];
    } finally { suppressed = false; }
  }
  function flush(): void {
    if (!editor || disposed || suppressed || composing) return;
    const baseRevision = compositionRevision ?? revision;
    const anchor = canonicalOffset(editor.state.doc, editor.state.selection.anchor), head = canonicalOffset(editor.state.doc, editor.state.selection.head);
    try {
      if (baseRevision !== session.snapshot().manifest.revision) throw new DocumentError("RevisionConflict", "输入期间文档已变化，请基于当前修订重新输入。");
      const html = sourceHtml(document, editor.state.doc), template = document.createElement("template"); template.innerHTML = html;
      const source = sourceDocument(), originalRoot = rootElement(source);
      let root = template.content.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(sourceRootId) + '"]');
      if (!root) { root = originalRoot.cloneNode(false) as HTMLElement; root.append(template.content); }
      const sections = session.snapshot().metadata.sections.map((section) => source.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(section.rootNodeId) + '"]'))
        .filter((section): section is HTMLElement => Boolean(section && section !== originalRoot && originalRoot.contains(section)));
      for (const [index, section] of sections.entries()) {
        if (root.querySelector('[data-node-id="' + CSS.escape(section.dataset.nodeId!) + '"]')) continue;
        let reference: Element | null = null;
        for (const item of sections.slice(index + 1)) {
          const following: HTMLElement | null = root.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(item.dataset.nodeId!) + '"]');
          if (following) { reference = following; break; }
        }
        while (reference?.parentElement && reference.parentElement !== root) reference = reference.parentElement;
        root.insertBefore(section.cloneNode(false), reference);
      }
      if (sections.length) {
        const first = root.querySelector('[data-node-id="' + CSS.escape(sections[0]!.dataset.nodeId!) + '"]')!;
        for (const child of [...root.childNodes]) {
          if (child === first || child.nodeType === 1 && sections.some((section) => (child as Element).querySelector('[data-node-id="' + CSS.escape(section.dataset.nodeId!) + '"]') || (child as Element).getAttribute("data-node-id") === section.dataset.nodeId)) continue;
          if (child.nodeType !== 3 || child.textContent!.trim()) first.append(child);
        }
      }
      const paragraph = originalRoot.querySelector("p"), run = paragraph?.querySelector('[data-word-run="true"]');
      for (const node of root.querySelectorAll("p:not([data-node-id])")) {
        if (paragraph && !node.hasAttribute("class") && !node.hasAttribute("style")) {
          for (const name of ["class", "style", "data-word-style", "data-word-line-rule"]) { const value = paragraph.getAttribute(name); if (value) node.setAttribute(name, value); }
        }
        if (run && !node.querySelector("[data-word-run]")) {
          const wrapper = document.createElement("span");
          for (const name of ["data-word-run", "class", "style"]) { const value = run.getAttribute(name); if (value) wrapper.setAttribute(name, value); }
          wrapper.append(...node.childNodes); node.append(wrapper);
        }
      }
      const candidate = session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision,
        operations: [{ op: "reconcileFragment", nodeId: sourceRootId, html: root.innerHTML, textEdits: proofs }] });
      committing = true;
      try { session.commit(candidate); } catch (reason) { session.discard(candidate); throw reason; }
      finally { committing = false; }
      error.hidden = true; compositionRevision = undefined; refresh(anchor, head);
    } catch (reason) { compositionRevision = undefined; proofs = []; refresh(); report(reason); }
  }
  function format(property: string, value: string): void {
    if (!editor || disposed) return;
    editor.chain().focus().setMark("officeFormat", { style: property + ":" + value }).run();
  }
  function button(label: string, action: () => void): HTMLButtonElement {
    const control = document.createElement("button"); control.type = "button"; control.textContent = label;
    control.addEventListener("mousedown", (event) => event.preventDefault(), { signal: controls.signal });
    control.addEventListener("click", () => { try { action(); } catch (reason) { report(reason); } }, { signal: controls.signal }); toolbar.append(control); return control;
  }
  button("加粗", () => format("font-weight", "700")); button("常规字重", () => format("font-weight", "400"));
  button("斜体", () => format("font-style", "italic")); button("下划线", () => format("text-decoration", "underline"));
  button("深蓝色", () => format("color", "#1f4b99"));
  const size = document.createElement("select"); size.setAttribute("aria-label", "字号");
  for (const value of [9, 10, 11, 12, 14, 16, 18, 24]) {
    const option = document.createElement("option"); option.value = value + "pt"; option.textContent = value + " 磅"; size.append(option);
  }
  size.value = "12pt"; size.addEventListener("change", () => format("font-size", size.value), { signal: controls.signal }); toolbar.append(size);
  for (const [label, alignment] of [["左对齐", "left"], ["居中", "center"], ["右对齐", "right"]] as const) button(label, () => {
    if (!editor) return;
    const ids = new Set<string>();
    editor.state.doc.nodesBetween(editor.state.selection.from, editor.state.selection.to, (node) => {
      if (node.isTextblock) { const id = attributes(node).source["data-node-id"]; if (id) ids.add(id); }
    });
    if (!ids.size) {
      for (let depth = editor.state.selection.$from.depth; depth > 0; depth--) {
        const node = editor.state.selection.$from.node(depth); if (node.isTextblock) { ids.add(attributes(node).source["data-node-id"]!); break; }
      }
    }
    const candidate = session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: revision,
      operations: [...ids].map((nodeId) => ({ op: "setStyle" as const, nodeId, property: "text-align", value: alignment })) });
    try { session.commit(candidate); } catch (reason) { session.discard(candidate); throw reason; }
  });
  button("撤销输入", () => { session.undo(); }); button("重做输入", () => { session.redo(); });

  function paste(event: ClipboardEvent): boolean {
    if (!editor || !event.clipboardData) return false;
    event.preventDefault();
    try {
      const raw = event.clipboardData.getData("text/html"), text = event.clipboardData.getData("text/plain");
      if (raw.length > 1024 * 1024 || text.length > 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "粘贴内容超过大小限制。");
      if (!raw) { editor.commands.insertContent(text.replace(/\r\n?/g, "\n").split("\n").flatMap((line, index) => [
        ...(index ? [{ type: "officeBreak", attrs: { tag: "br" } }] : []), ...(line ? [{ type: "text", text: line }] : []),
      ])); return true; }
      const template = document.createElement("template"); template.innerHTML = raw;
      for (const node of template.content.querySelectorAll("[data-node-id],[id]")) { node.removeAttribute("data-node-id"); node.removeAttribute("id"); }
      if (![...template.content.children].some((node) => /^(?:p|h[1-6]|div|section|article|table|ul|ol)$/.test(node.localName))) {
        const paragraph = document.createElement("p"); paragraph.append(template.content); template.content.append(paragraph);
      }
      const baseline = session.getNode(sourceRootId)!;
      const validation = session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: revision,
        operations: [{ op: "insertFragment", nodeId: sourceRootId, position: "end", html: template.innerHTML }] });
      const checked = new DOMParser().parseFromString(validation.preview.html, "text/html");
      const root = checked.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(sourceRootId) + '"]')!;
      const original = new Set(baseline.childNodeIds);
      const nodes = [...root.children].filter((node) => !original.has(node.getAttribute("data-node-id")!));
      session.discard(validation);
      const json = nodes.map((node) => {
        for (const element of [node, ...node.querySelectorAll("[data-node-id]")]) element.removeAttribute("data-node-id");
        return sourceJson(node);
      });
      editor.commands.insertContent(json);
    } catch (reason) { report(reason); }
    return true;
  }

  const unsubscribe = session.subscribe((event) => {
    if (event.type !== "error" && !committing && !disposed && !composing) refresh();
  });
  preview = mountDocumentPreview(stage, session, { title: "正文编辑" });
  const ready = preview.ready.then(() => {
    if (disposed || !preview) return;
    const frame = preview.element.contentDocument!, root = rootElement(frame), source = sourceDocument();
    const imageUrls = new Map<string, string>();
    for (const image of frame.images) {
      const original = source.querySelector<HTMLImageElement>('[data-node-id="' + CSS.escape(image.dataset.nodeId ?? "") + '"]');
      if (original) imageUrls.set(original.getAttribute("src")!, image.src);
    }
    const viewStyles = new Map<string, { property: string; original: string; rendered: string }[]>();
    for (const node of frame.querySelectorAll<HTMLElement>("[data-node-id]")) {
      const original = source.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(node.dataset.nodeId!) + '"]'); if (!original) continue;
      const changed = [...node.style].filter((property) => node.style.getPropertyValue(property) !== original.style.getPropertyValue(property))
        .map((property) => ({ property, original: original.style.getPropertyValue(property), rendered: node.style.getPropertyValue(property) }));
      viewStyles.set(node.dataset.nodeId!, changed);
    }
    const host = frame.createElement("div"); root.replaceWith(host);
    const viewStyle = frame.createElement("style"); viewStyle.textContent = ".tiptap:focus{outline:none}.tiptap{min-height:100px}.tiptap [data-word-tab]{white-space:pre}.tiptap [data-word-field],.tiptap [data-word-field-end]{background:#f5f4ee}"; frame.head.append(viewStyle);
    editor = new Editor({ element: host, injectCSS: false, content: content(), parseOptions: { preserveWhitespace: "full" },
      extensions: [...editorExtensions((properties) => {
        const result = { ...properties }, probe = frame.createElement("span"); probe.setAttribute("style", properties.style ?? "");
        for (const item of viewStyles.get(properties["data-node-id"] ?? "") ?? [])
          if (probe.style.getPropertyValue(item.property) === item.original) probe.style.setProperty(item.property, item.rendered);
        if (probe.style.cssText) result.style = probe.style.cssText;
        if (properties.src && imageUrls.has(properties.src)) result.src = imageUrls.get(properties.src)!;
        return result;
      }), Extension.create({ name: "officeKeymap", addProseMirrorPlugins: () => [keymap(baseKeymap)] })],
      editorProps: {
        attributes: { role: "textbox", "aria-label": "文档正文", "aria-multiline": "true", spellcheck: "false" },
        handlePaste: (_view, event) => paste(event),
        handleKeyDown: (_view, event) => {
          if ((event.ctrlKey || event.metaKey) && ["b", "i", "u"].includes(event.key.toLowerCase())) {
            event.preventDefault(); const key = event.key.toLowerCase();
            format(key === "b" ? "font-weight" : key === "i" ? "font-style" : "text-decoration", key === "b" ? "700" : key === "i" ? "italic" : "underline"); return true;
          }
          if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
            event.preventDefault(); try { if (event.shiftKey) session.redo(); else session.undo(); } catch (reason) { report(reason); } return true;
          }
          if (event.key === "Enter" && editor) {
            const parent = editor.state.selection.$from.parent;
            if (["officeCell", "officeItem"].includes(parent.type.name) || event.shiftKey) {
              editor.commands.insertContent({ type: "officeBreak", attrs: { tag: "br" } }); return true;
            }
          }
          if (event.key === "Tab" && editor) {
            const cells: { position: number; size: number }[] = [];
            editor.state.doc.descendants((node, position) => { if (["officeCell", "officeBlockCell"].includes(node.type.name)) cells.push({ position, size: node.nodeSize }); });
            const current = cells.findIndex((cell) => cell.position < editor!.state.selection.from && editor!.state.selection.from < cell.position + cell.size);
            const next = cells[current + (event.shiftKey ? -1 : 1)];
            if (current >= 0 && next) {
              event.preventDefault(); editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(next.position + 1)))); return true;
            }
          }
          return false;
        },
        handleDOMEvents: {
          compositionstart: () => { composing = true; compositionRevision = revision; return false; },
          compositionend: () => { queueMicrotask(() => { composing = false; flush(); }); return false; },
        },
      },
      onTransaction: ({ transaction }: { transaction: Transaction }) => {
        if (suppressed || disposed || !transaction.docChanged) return;
        proofs.push(...transactionProofs(transaction)); if (!composing) flush();
      },
    });
    editor.view.dom.setAttribute("data-office-editor", "");
  }).catch((reason: unknown) => { report(reason); throw reason; });
  return { element, ready, focus: () => { editor?.commands.focus(); }, dispose: () => {
    if (disposed) return; disposed = true; unsubscribe(); controls.abort(); editor?.destroy(); editor = undefined;
    preview?.dispose(); preview = undefined; proofs = []; renderDocument = undefined; element.remove();
  } };
}
