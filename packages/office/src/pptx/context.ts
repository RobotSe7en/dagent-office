import { createAssetDescriptor, DocumentError } from "@dagent-office/core";
import type { Annotation, DocumentManifest, PresentationMetadata } from "@dagent-office/core";
import { openOpc } from "../ooxml/package.js";
import type { OpcPackage } from "../ooxml/package.js";
import { attr, fail, number, R, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { declarations, escapeHtml } from "../ooxml/styles.js";
import type { Css } from "../ooxml/styles.js";

export interface Html { tag: string; attrs: Record<string, string>; children: (Html | string)[] }
export interface Surface { id: string; root: Html; source: XmlNode; themeId: string; masterId?: string; layoutId?: string }
export interface NativeParagraph { nodeId: string; start: number; text: string; nativeText: string }
export function serialize(node: Html): string {
  const attrs = Object.entries(node.attrs).map(([name, value]) => " " + name + '="' + escapeHtml(value) + '"').join("");
  return "<" + node.tag + attrs + ">" + (["br", "img"].includes(node.tag) ? "" : node.children.map(item => typeof item === "string" ? escapeHtml(item) : serialize(item)).join("") + "</" + node.tag + ">");
}
export const encode = (value: string): Uint8Array => new (globalThis as unknown as { TextEncoder: new () => { encode(value: string): Uint8Array } }).TextEncoder().encode(value);
export const value = (node: XmlNode | undefined, name = "val"): string | undefined => attr(node, name, "");
export const integer = (node: XmlNode, name: string, fallback?: number): number => number(value(node, name), node, fallback);
export function unsigned(node: XmlNode, name: string, minimum = 0, maximum = 4294967295): number {
  const n = integer(node, name);
  if (n < minimum || n > maximum) fail(node, "A native unsigned identity or index is outside its range.");
  return n;
}
export const points = (node: XmlNode, name: string, fallback?: number): number => integer(node, name, fallback) / 12700;
export function bool(node: XmlNode, name: string, fallback = false): boolean {
  const v = value(node, name);
  if (v === undefined) return fallback;
  if (["true", "1"].includes(v)) return true;
  if (["false", "0"].includes(v)) return false;
  return fail(node, "A DrawingML boolean is invalid.");
}
export function rootName(node: XmlNode, name: string, uri: string): void {
  if (node.uri !== uri || node.name !== name) fail(node, "An OOXML part has the wrong root.");
}
export function fontName(name: string): string {
  if (/[\u0000-\u001f\u007f]/.test(name)) throw new DocumentError("InvalidDocument", "A font name contains control characters.");
  return '"' + name.replace(/["\\]/g, "\\$&") + '"';
}
export class Context {
  serial = 0;
  work = 0;
  outputBytes = 0;
  readonly css: string[] = [];
  readonly assets: DocumentManifest["assets"] = [];
  readonly files: Map<string, Uint8Array> = new Map();
  readonly surfaces: Map<string, Surface> = new Map();
  readonly nativeNodes: Map<string, Map<string, string>> = new Map();
  readonly slidePaths: Map<string, string> = new Map();
  readonly annotations: Annotation[] = [];
  readonly paragraphs: Map<string, NativeParagraph[]> = new Map();
  readonly workbooks: Map<string, OpcPackage> = new Map();
  readonly dataSetsByRange: Map<string, string> = new Map();
  nestedBytes: number;
  nestedFiles: number;
  readonly metadata: PresentationMetadata = { size: { unit: "pt", width: 960, height: 540 }, commentFormat: "modern",
    slides: [], themes: [], masters: [], layouts: [], notesMasters: [], stories: [], objects: [], charts: [], dataSets: [], comments: [], hyperlinks: [] };
  constructor(readonly opc: OpcPackage) { this.nestedBytes = [...opc.files.values()].reduce((n, bytes) => n + bytes.length, 0); this.nestedFiles = opc.files.size; }
  workbook(path: string): OpcPackage {
    const cached = this.workbooks.get(path);
    if (cached) return cached;
    const limits = { ...this.opc.limits, maxPackageBytes: this.opc.limits.maxPackageBytes - this.nestedBytes, maxFiles: this.opc.limits.maxFiles - this.nestedFiles };
    if (limits.maxPackageBytes <= 0 || limits.maxFiles <= 0) throw new DocumentError("ResourceLimitExceeded", "Nested chart packages exceed the shared budget.");
    const book = openOpc(this.opc.files.get(path)!, { limits }, this.opc.xmlBudget);
    this.nestedBytes += [...book.files.values()].reduce((n, data) => n + data.length, 0); this.nestedFiles += book.files.size;
    this.workbooks.set(path, book); return book;
  }
  count(amount = 1): void {
    this.work += amount;
    if (this.work > this.opc.limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "PPTX mapping exceeds its work budget.");
  }
  bytes(amount: number): void {
    this.outputBytes += amount;
    if (this.outputBytes > this.opc.limits.maxPackageBytes) throw new DocumentError("ResourceLimitExceeded", "PPTX mapped output exceeds its byte budget.");
  }
  make(tag: string, attrs: Record<string, string> = {}, children: (Html | string)[] = []): Html {
    this.count();
    this.bytes(tag.length * 2 + 64 + Object.entries(attrs).reduce((n, [k, v]) => n + k.length + v.length * 6, 0));
    return { tag, attrs: { "data-node-id": this.id("node"), ...attrs }, children };
  }
  id(prefix: string): string { return prefix + "-" + ++this.serial; }
  content(node: Html, value: string): void { this.bytes(value.length * 6); node.children.push(value); }
  style(node: Html, css: Css): void {
    const source = declarations(css);
    this.bytes(source.length * 6);
    if (source) node.attrs.style = source;
  }
  rule(name: string, css: Css): void { const text = "." + name + "{" + declarations(css) + "}"; this.bytes(text.length * 4); this.css.push(text); }
  relation(source: string, id: string | undefined, kind: string, external = false): { target: string; external: boolean } {
    const relation = id ? this.opc.relationships(source).get(id) : undefined;
    if (!relation || relation.type !== R + "/" + kind) fail({ part: source, name: "Relationship" }, "A required relationship is missing or has the wrong type.");
    if (relation.external && !external) unsupported({ part: source, name: "Relationship" } as XmlNode, "External PPTX resources are not supported.");
    return relation;
  }
  related(source: string, kind: string, required = true): string | undefined {
    const list = [...this.opc.relationships(source).values()].filter(r => r.type === R + "/" + kind);
    if (list.length > 1 || required && !list.length) fail({ part: source, name: kind }, "A singleton PPTX relationship is missing or ambiguous.");
    return list[0] ? this.relation(source, list[0].id, kind).target : undefined;
  }
  part(path: string, kind: string, name: string, uri: string): XmlNode {
    if (this.opc.contentType(path) !== "application/vnd.openxmlformats-officedocument." + kind + "+xml") fail({ part: path, name }, "A PPTX part content type is invalid.");
    const root = this.opc.read(path); rootName(root, name, uri); return root;
  }
  image(source: string, id: string | undefined): string {
    const path = this.relation(source, id, "image").target, data = this.opc.files.get(path)!;
    const mime = this.opc.contentType(path);
    if (!mime || !["image/png", "image/jpeg", "image/webp"].includes(mime)) unsupported({ part: path, name: "image" } as XmlNode, "This image container is not supported.");
    const known = this.assets.find(a => a.path === "assets/" + path);
    if (known) return known.path;
    const target = "assets/" + path;
    this.assets.push(createAssetDescriptor(target, mime, data)); this.files.set(target, data); this.bytes(data.length);
    return target;
  }
}
