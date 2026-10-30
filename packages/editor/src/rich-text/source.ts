import { Mark, Node as TiptapNode } from "@tiptap/core";
import type { Extensions, JSONContent } from "@tiptap/core";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { DocumentError } from "@dagent-office/core";

const blocks = new Set(["article", "section", "div", "p", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "table", "thead", "tbody", "tfoot", "tr", "td", "th"]);
const inline = new Set(["span", "strong", "em", "u", "s", "sub", "sup", "ins", "del"]);
export interface SourceAttributes {
  readonly tag: string;
  readonly source: Record<string, string>;
  readonly before: string;
  readonly beforeText: string;
  readonly tail: string;
  readonly tailText: string;
  readonly comment: string;
}
export function attributes(node: ProseMirrorNode): SourceAttributes { return node.attrs as SourceAttributes; }
export function sourceId(node: ProseMirrorNode): string | undefined { return attributes(node).source?.["data-node-id"]; }
function hasBlocks(element: Element): boolean { return [...element.children].some((child) => blocks.has(child.localName)); }
function typeOf(element: Element): string {
  const tag = element.localName;
  if (inline.has(tag)) return "officeInline";
  if (tag === "img") return "officeImage";
  if (tag === "br") return "officeBreak";
  if (tag === "table") return "officeTable";
  if (["thead", "tbody", "tfoot"].includes(tag)) return "officeTableGroup";
  if (tag === "tr") return "officeRow";
  if (["td", "th"].includes(tag)) return hasBlocks(element) ? "officeBlockCell" : "officeCell";
  if (["ul", "ol"].includes(tag)) return "officeList";
  if (tag === "li") return hasBlocks(element) ? "officeBlockItem" : "officeItem";
  if (["article", "section", "div"].includes(tag)) return hasBlocks(element) || !element.childNodes.length ? "officeContainer" : "officeParagraph";
  if (/^(?:p|h[1-6])$/.test(tag)) return "officeParagraph";
  throw new DocumentError("UnsupportedFeature", "The rich-text input adapter cannot express this element.");
}
function inlineContent(type: string): boolean { return ["officeParagraph", "officeInline", "officeCell", "officeItem"].includes(type); }

/** Derived editor JSON preserves source attributes and non-rendered block whitespace. */
export function sourceJson(element: Element): JSONContent {
  const type = typeOf(element), attrs: Record<string, unknown> = {
    tag: element.localName, source: Object.fromEntries([...element.attributes].map((attr) => [attr.name, attr.value])),
    before: "", beforeText: "", tail: "", tailText: "", comment: "",
  };
  const content: JSONContent[] = [];
  let gap = "", gapText = "";
  for (const child of element.childNodes) {
    if (child.nodeType === 3) {
      const text = child.nodeValue ?? "";
      if (inlineContent(type)) { if (text) content.push({ type: "text", text }); }
      else {
        if (text.trim()) throw new DocumentError("UnsupportedFeature", "Mixed block and inline container content cannot be edited without an explicit structure change.");
        gap += text; gapText += text;
      }
    } else if (child.nodeType === 8) {
      if (inlineContent(type)) content.push({ type: "officeComment", attrs: { comment: child.nodeValue ?? "" } });
      else gap += "<!--" + child.nodeValue + "-->";
    } else if (child.nodeType === 1) {
      const node = sourceJson(child as Element);
      if (!inlineContent(type)) { node.attrs = { ...node.attrs, before: gap, beforeText: gapText }; gap = gapText = ""; }
      content.push(node);
    }
  }
  attrs.tail = gap; attrs.tailText = gapText;
  if (type === "officeContainer" && content.length === 0) content.push({ type: "officeParagraph", attrs: { tag: "p", source: {} } });
  return { type, attrs, ...(content.length ? { content } : {}) };
}

export function editorExtensions(viewAttributes: (source: Record<string, string>) => Record<string, string>): Extensions {
  const attrs = () => ({ tag: { default: "p" }, source: { default: {} }, before: { default: "" }, beforeText: { default: "" },
    tail: { default: "" }, tailText: { default: "" }, comment: { default: "" } });
  const node = (name: string, tags: string[], group: string, content?: string, leaf = false): TiptapNode => TiptapNode.create({
    name, group, ...(content ? { content } : {}), inline: group === "inline", atom: leaf,
    ...(name === "officeContainer" ? { isolating: true, defining: true, selectable: false } : {}),
    ...(name.includes("Cell") || name === "officeCell" ? { isolating: true } : {}),
    addAttributes: attrs,
    parseHTML: () => tags.map((tag) => ({ tag, getAttrs: (element: HTMLElement) => typeOf(element) === name ? {
      tag, source: Object.fromEntries([...element.attributes].map((attr) => [attr.name, attr.value])),
    } : false })),
    renderHTML: ({ node: value }) => {
      const data = attributes(value), properties = viewAttributes(data.source);
      return leaf ? [data.tag || tags[0]!, properties] : [data.tag || tags[0]!, properties, 0];
    },
  });
  return [
    TiptapNode.create({ name: "doc", topNode: true, content: "block+" }),
    TiptapNode.create({ name: "text", group: "inline" }),
    node("officeParagraph", ["p", "h1", "h2", "h3", "h4", "h5", "h6", "div", "section", "article"], "block", "inline*"),
    node("officeContainer", ["article", "section", "div"], "block", "block*"),
    node("officeInline", [...inline], "inline", "inline*"),
    node("officeImage", ["img"], "inline", undefined, true), node("officeBreak", ["br"], "inline", undefined, true),
    node("officeTable", ["table"], "block", "(officeRow | officeTableGroup)*"),
    node("officeTableGroup", ["thead", "tbody", "tfoot"], "tableGroup", "officeRow*"),
    node("officeRow", ["tr"], "row", "cell*"),
    node("officeCell", ["td", "th"], "cell", "inline*"), node("officeBlockCell", ["td", "th"], "cell", "block*"),
    node("officeList", ["ul", "ol"], "block", "listItem*"),
    node("officeItem", ["li"], "listItem", "inline*"), node("officeBlockItem", ["li"], "listItem", "block*"),
    TiptapNode.create({ name: "officeComment", inline: true, group: "inline", atom: true, addAttributes: attrs,
      renderHTML: () => ["span", { hidden: "", "data-sdk-comment": "" }] }),
    Mark.create({ name: "officeFormat", addAttributes: () => ({ style: { default: "" } }),
      parseHTML: () => [], renderHTML: ({ mark }) => ["span", { style: mark.attrs.style as string }, 0] }),
  ];
}

export function canonicalText(node: ProseMirrorNode): string {
  if (node.isText) return node.text ?? "";
  if (node.type.name === "officeImage") return "\ufffc";
  if (node.type.name === "officeBreak") return "\n";
  if (node.type.name === "officeComment") return "";
  const parts: string[] = [];
  node.forEach((child) => { parts.push(attributes(child).beforeText ?? "", canonicalText(child)); });
  parts.push(attributes(node).tailText ?? ""); return parts.join("");
}

/** Convert ProseMirror positions (which include structural tokens) to canonical UTF-16. */
export function canonicalOffset(node: ProseMirrorNode, position: number): number {
  let total = 0, done = false;
  node.forEach((child, offset) => {
    if (done) return;
    if (position <= offset) { done = true; return; }
    total += attributes(child).beforeText?.length ?? 0;
    if (position < offset + child.nodeSize) {
      total += child.isText ? Math.max(0, Math.min(child.text!.length, position - offset))
        : child.isLeaf ? canonicalText(child).length
          : canonicalOffset(child, Math.max(0, position - offset - 1));
      done = true;
    } else total += canonicalText(child).length;
  });
  if (!done) total += attributes(node).tailText?.length ?? 0;
  return total;
}

export function sourceHtml(document: Document, node: ProseMirrorNode): string {
  const seen = new Set<string>(), domIds = new Set<string>();
  const appendGap = (parent: globalThis.Node, html: string): void => {
    if (!html) return;
    const template = document.createElement("template"); template.innerHTML = html; parent.appendChild(template.content);
  };
  const render = (value: ProseMirrorNode): globalThis.Node => {
    let result: globalThis.Node;
    if (value.isText) result = document.createTextNode(value.text ?? "");
    else if (value.type.name === "officeComment") result = document.createComment(attributes(value).comment);
    else {
      const data = attributes(value), element = document.createElement(data.tag || (value.type.name === "officeBreak" ? "br" : "p"));
      const source = { ...data.source }, id = source["data-node-id"];
      if (id) { if (seen.has(id)) { delete source["data-node-id"]; delete source.id; } else seen.add(id); }
      if (source.id) { if (domIds.has(source.id)) delete source.id; else domIds.add(source.id); }
      for (const [name, content] of Object.entries(source)) element.setAttribute(name, content);
      value.forEach((child) => { appendGap(element, attributes(child).before ?? ""); element.appendChild(render(child)); });
      appendGap(element, data.tail ?? ""); result = element;
    }
    for (const mark of [...value.marks].reverse()) {
      const span = document.createElement("span"); span.setAttribute("style", mark.attrs.style as string); span.appendChild(result); result = span;
    }
    return result;
  };
  const template = document.createElement("template");
  node.forEach((child) => { appendGap(template.content, attributes(child).before ?? ""); template.content.appendChild(render(child)); });
  return template.innerHTML;
}
