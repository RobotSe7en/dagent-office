import type { Anchor, Annotation, AnnotationReply, PresentationMetadata, PresentationObject, TextRange } from "@dagent-office/core";
import { attr, check, child, elements, fail, P, R, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { Context, integer, points, rootName, unsigned, value } from "./context.js";
import type { Html, Surface } from "./context.js";
import { importText } from "./text.js";

const MODERN = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
const PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
const AC = "http://schemas.microsoft.com/office/drawing/2013/main/command";
const REL = "http://schemas.microsoft.com/office/2018/10/relationships/";
const EXT = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}";
type Author = { id: string; display: NonNullable<Annotation["author"]> };
function guid(node: XmlNode, name: string): string {
  const id = value(node, name);
  if (!id || !/^\{[\da-f]{8}-(?:[\da-f]{4}-){3}[\da-f]{12}\}$/i.test(id)) fail(node, "A native comment GUID is invalid.");
  return id.toUpperCase();
}
function date(node: XmlNode, name: string): string {
  const d = value(node, name);
  if (!d || !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)?$/.test(d) || !Number.isFinite(Date.parse(d))) fail(node, "A comment timestamp is invalid.");
  return d;
}
function status(node: XmlNode): "open" | "resolved" | "closed" {
  const s = value(node, "status") ?? "active";
  if (s === "active") return "open";
  if (s === "resolved" || s === "closed") return s;
  return fail(node, "A native comment status is invalid.");
}
/** MS-ODRAWXML 1.3.9.1, unsigned 32-bit hash over UTF-16 code units. */
export function nativeTextHash(input: string): number {
  let hash = 0;
  for (let index = 0; index < input.length; index++) hash = (Math.imul(hash, 33) + input.charCodeAt(index)) >>> 0;
  return hash;
}
export class Comments {
  readonly authors: Map<string, Author> = new Map();
  readonly identities: Set<string> = new Set();
  readonly parts: Set<string> = new Set();
  readonly objects: Map<string, PresentationObject>;
  readonly parents: Map<string, string> = new Map();
  constructor(readonly context: Context, readonly main: string, readonly documentId: string, readonly roots: Html[]) {
    this.objects = new Map(context.metadata.objects.map(o => [o.nodeId, o]));
    const visit = (node: Html): void => { for (const child of node.children) if (typeof child !== "string") { context.count(); this.parents.set(child.attrs["data-node-id"]!, node.attrs["data-node-id"]!); visit(child); } };
    for (const root of roots) visit(root);
    for (const relation of context.opc.relationships(main).values()) {
      if (relation.type !== R + "/commentAuthors" && relation.type !== REL + "authors") continue;
      if (relation.external) unsupported({ part: main, name: "authors" } as XmlNode, "External comment authors are not supported.");
      const modern = relation.type === REL + "authors", uri = modern ? MODERN : P;
      const mime = modern ? "application/vnd.ms-powerpoint.authors+xml" : "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
      if (context.opc.contentType(relation.target) !== mime) fail({ part: relation.target, name: "authors" }, "A comment author part has the wrong content type.");
      const source = context.opc.read(relation.target); rootName(source, modern ? "authorLst" : "cmAuthorLst", uri);
      check(source, [modern ? "author" : "cmAuthor"], uri);
      for (const native of elements(source, modern ? "author" : "cmAuthor", uri)) {
        const extensions = child(native, "extLst", uri);
        if (extensions?.content.some(n => typeof n !== "string" || n.trim())) unsupported(extensions, "Author extensions are not supported.");
        context.count();
        const nativeId = modern ? guid(native, "id") : String(unsigned(native, "id"));
        const key = (modern ? "modern:" : "classic:") + nativeId;
        if (this.authors.has(key)) fail(native, "Comment author identities are duplicated.");
        const name = value(native, "name"), initials = value(native, "initials");
        if (name === undefined) fail(native, "A comment author name is missing.");
        const authorId = context.id("comment-author"), display = { name, ...(initials !== undefined ? { initials } : {}) };
        if (modern && (value(native, "userId") === undefined || value(native, "providerId") === undefined)) fail(native, "Modern author provider fields are missing.");
        context.bytes(128 + (value(native, "userId")?.length ?? 0) * 6 + (value(native, "providerId")?.length ?? 0) * 6);
        (context.metadata.commentAuthors ??= []).push({ authorId, format: modern ? "modern" : "classic", nativeId,
          ...(modern ? { userId: value(native, "userId")!, providerId: value(native, "providerId")! } : {}) });
        this.authors.set(key, { id: authorId, display });
      }
    }
  }
  body(slideId: string, source?: XmlNode, plain?: string): Html {
    const id = this.context.id("comment-story"), root = this.context.make("section", { "data-office-kind": "comment", "data-office-story": "comment" });
    if (source) importText(this.context, source, root, source.part);
    else { const p = this.context.make("p"); if (plain !== undefined) this.context.content(p, plain); root.children.push(p); }
    this.roots.push(root); this.context.metadata.stories.push({ kind: "comment", storyId: id, rootNodeId: root.attrs["data-node-id"]!, slideId }); return root;
  }
  read(slide: Surface, slidePath: string): void {
    const relations = [...this.context.opc.relationships(slidePath).values()].filter(r => r.type === R + "/comments" || r.type === REL + "comments");
    const formats = new Set<string>();
    for (const relation of relations) {
      const modern = relation.type === REL + "comments", format = modern ? "modern" : "classic";
      if (formats.has(format) || this.parts.has(relation.target)) fail(slide.source, "Comment parts have duplicate or conflicting ownership.");
      formats.add(format); this.parts.add(relation.target);
      if (relation.external) unsupported(slide.source, "External comment parts are not supported.");
      if (modern) {
        const refs = elements(child(slide.source, "extLst", P), "ext", P).filter(ext => value(ext, "uri")?.toUpperCase() === EXT).flatMap(ext => elements(ext, "commentRel", MODERN));
        if (refs.length !== 1 || attr(refs[0], "id", R) !== relation.id) fail(slide.source, "A modern comment relationship extension is missing or ambiguous.");
      }
      const mime = modern ? "application/vnd.ms-powerpoint.comments+xml" : "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
      if (this.context.opc.contentType(relation.target) !== mime) fail(slide.source, "A comment part has the wrong content type.");
      const source = this.context.opc.read(relation.target); rootName(source, "cmLst", modern ? MODERN : P);
      check(source, ["cm"], modern ? MODERN : P);
      for (const comment of elements(source, "cm", modern ? MODERN : P)) modern ? this.modern(comment, slide) : this.classic(comment, slide);
    }
  }
  classic(source: XmlNode, slide: Surface): void {
    check(source, ["pos", "text", "extLst"], P);
    if (elements(child(source, "extLst", P), undefined, P).length) unsupported(source, "Extended classic comment threads are not supported.");
    const authorNative = String(unsigned(source, "authorId")), index = unsigned(source, "idx"), key = "classic:" + authorNative;
    const author = this.authors.get(key);
    if (!author) fail(source, "A classic comment references an undefined author.");
    this.authorBytes(author);
    const identity = key + ":" + index;
    if (this.identities.has(identity)) fail(source, "Classic comment identities are duplicated.");
    this.identities.add(identity);
    const position = child(source, "pos", P), content = child(source, "text", P);
    if (!position || !content) fail(source, "A classic comment is missing its position or text.");
    // Independent UNO probe: 25.4 mm (72 pt) serializes as 576, i.e. 1/8 pt.
    const x = integer(position, "x") / 8, y = integer(position, "y") / 8;
    if (x < 0 || y < 0) unsupported(position, "Negative classic comment points cannot be represented by the current region anchor.");
    const body = this.body(slide.id, undefined, text(content)), annotationId = this.context.id("annotation");
    this.context.annotations.push({ annotationId, documentId: this.documentId, baseRevision: 0, status: "open", sourceId: identity,
      author: author.display, date: date(source, "dt"), bodyNodeId: body.attrs["data-node-id"]!,
      anchor: { kind: "region", surfaceId: slide.root.attrs["data-node-id"]!, pageId: slide.id, rect: { x, y, width: 1, height: 1 }, nodeIds: [], layoutKey: "pptx-classic-point" } });
    this.context.metadata.comments.push({ annotationId, format: "classic", authorId: author.id });
    if (!this.context.metadata.comments.some(c => c.format === "modern")) this.context.metadata.commentFormat = "classic";
  }
  modern(source: XmlNode, slide: Surface): void {
    this.extensions(source);
    const rootId = guid(source, "id"), author = this.authors.get("modern:" + guid(source, "authorId"));
    if (!author) fail(source, "A modern comment references an undefined author.");
    this.authorBytes(author);
    this.identity(rootId, source);
    for (const name of ["startDate", "dueDate", "assignedTo", "title"]) if (value(source, name)) unsupported(source, "Comment task and assignment fields are not supported.");
    if (value(source, "complete") && !["0", "0%"].includes(value(source, "complete")!)) unsupported(source, "Comment tasks are not supported.");
    const anchors = elements(source, undefined, MODERN).filter(n => n.name === "unknownAnchor");
    const monikers = source.content.filter((n): n is XmlNode => typeof n !== "string" && (n.uri === PC && n.name === "sldMkLst" || n.uri === AC && ["deMkLst", "txMkLst"].includes(n.name)));
    for (const node of source.content) if (typeof node !== "string" && !monikers.includes(node) && !(node.uri === MODERN && ["unknownAnchor", "pos", "replyLst", "txBody", "extLst"].includes(node.name))) unsupported(node, "This modern comment property is not supported.");
    if (anchors.length > 1 || anchors.length && monikers.length || !anchors.length && !monikers.length) fail(source, "A modern comment requires one anchor representation.");
    const workflow = status(source), mapped = anchors.length ? { anchor: { kind: "node" as const, nodeId: "unresolved-" + rootId }, recovery: true } : { anchor: this.anchor(monikers, slide), recovery: false };
    const body = this.body(slide.id, child(source, "txBody", MODERN)), annotationId = this.context.id("annotation");
    const position = child(source, "pos", MODERN);
    if (position) { body.attrs["data-office-comment-x"] = String(points(position, "x")); body.attrs["data-office-comment-y"] = String(points(position, "y")); }
    const replies: AnnotationReply[] = [], replyAuthors: NonNullable<PresentationMetadata["comments"][number]["replyAuthors"]> = [];
    const replyList = child(source, "replyLst", MODERN);
    check(replyList, ["reply"], MODERN);
    for (const reply of elements(replyList, "reply", MODERN)) {
      this.extensions(reply);
      check(reply, ["txBody", "extLst"], MODERN);
      const nativeId = guid(reply, "id"), replyAuthor = this.authors.get("modern:" + guid(reply, "authorId"));
      if (!replyAuthor) fail(reply, "A reply references an undefined author.");
      this.authorBytes(replyAuthor);
      this.identity(nativeId, reply);
      const replyBody = this.body(slide.id, child(reply, "txBody", MODERN)), replyId = this.context.id("reply");
      replies.push({ replyId, author: replyAuthor.display, date: date(reply, "created"), status: status(reply), sourceId: nativeId, bodyNodeId: replyBody.attrs["data-node-id"]! });
      replyAuthors.push({ replyId, authorId: replyAuthor.id });
    }
    this.context.annotations.push({ annotationId, documentId: this.documentId, baseRevision: 0, status: mapped.recovery ? "needs_reanchor" : workflow,
      anchor: mapped.anchor, author: author.display, date: date(source, "created"), sourceId: rootId, bodyNodeId: body.attrs["data-node-id"]!, ...(replies.length ? { replies } : {}) });
    this.context.metadata.comments.push({ annotationId, format: "modern", authorId: author.id, ...(mapped.recovery ? { recoveryStatus: workflow } : {}), ...(replyAuthors.length ? { replyAuthors } : {}) });
    this.context.metadata.commentFormat = "modern";
  }
  identity(id: string, source: XmlNode): void {
    this.context.count();
    if (this.identities.has(id)) fail(source, "Modern comment and reply identities must be unique.");
    this.identities.add(id);
  }
  authorBytes(author: Author): void {
    this.context.bytes(512 + (author.display.name.length + (author.display.initials?.length ?? 0)) * 6);
  }
  extensions(source: XmlNode): void {
    const list = child(source, "extLst", MODERN);
    if (list?.content.some(n => typeof n !== "string" || n.trim())) unsupported(list, "Undeclared modern comment extensions are not supported.");
  }
  anchor(lists: XmlNode[], slide: Surface): Anchor {
    const nodes: string[] = [], ranges: TextRange[] = [];
    for (const list of lists) {
      const contents = list.content.filter((n): n is XmlNode => typeof n !== "string");
      const document = contents.shift(), page = contents.shift();
      if (!document || document.uri !== PC || document.name !== "docMk" || document.attrs.size || document.content.length) unsupported(list, "This document moniker is not supported.");
      if (!page || page.uri !== PC || page.name !== "sldMk" || "slide-" + unsigned(page, "sldId", 256, 2147483647) !== slide.id) fail(list, "A modern anchor refers to a different or missing slide.");
      if (list.name === "sldMkLst") { if (contents.length) unsupported(list); nodes.push(slide.root.attrs["data-node-id"]!); continue; }
      let target: string | undefined, lastGroup: string | undefined, textMoniker: XmlNode | undefined;
      for (const moniker of contents) {
        if (moniker.uri === AC && moniker.name === "txMk" && list.name === "txMkLst") { if (textMoniker) fail(moniker, "A text range moniker is repeated."); textMoniker = moniker; continue; }
        const kinds: Record<string, PresentationObject["kind"][]> = { spMk: ["text", "geometry"], grpSpMk: ["group"], graphicFrameMk: ["table", "chart"], picMk: ["image"], cxnSpMk: ["connector"] };
        if (moniker.uri !== AC || !kinds[moniker.name] || textMoniker) unsupported(moniker, "This drawing moniker is not supported.");
        const id = this.context.nativeNodes.get(slide.id)!.get(String(unsigned(moniker, "id"))), object = id && this.objects.get(id);
        if (!id || !object || !kinds[moniker.name]!.includes(object.kind)) fail(moniker, "A modern anchor targets a missing or different object type.");
        if (value(moniker, "creationId") !== undefined && (!object.creationId || guid(moniker, "creationId") !== object.creationId.toUpperCase())) fail(moniker, "A modern anchor creation identity no longer matches its object.");
        if (lastGroup) { let parent = this.parents.get(id); while (parent && parent !== lastGroup) parent = this.parents.get(parent); if (!parent) fail(moniker, "A drawing moniker group path is invalid."); }
        target = id; lastGroup = object.kind === "group" ? id : undefined;
      }
      if (!target) fail(list, "A drawing anchor target is missing.");
      if (list.name === "deMkLst") nodes.push(target);
      else {
        if (!textMoniker) fail(list, "A modern text range is missing.");
        const start = integer(textMoniker, "cp"), length = value(textMoniker, "len") === undefined ? 0 : integer(textMoniker, "len");
        if (start < 0 || length < 0 || !Number.isSafeInteger(start + length)) fail(textMoniker, "A modern text range is invalid.");
        const paragraphs = this.context.paragraphs.get(target);
        if (!paragraphs?.length) unsupported(textMoniker, "This object's native text scope is not supported.");
        check(textMoniker, ["context"], AC);
        const sourceContext = child(textMoniker, "context", AC);
        if (sourceContext) {
          const plain = paragraphs.map(p => p.nativeText).join("\r"), hash = unsigned(sourceContext, "hash");
          const length = value(sourceContext, "len") === undefined ? 0 : unsigned(sourceContext, "len");
          // Producers can include the final implicit paragraph mark. The native
          // length/hash must identify the exact current body in either encoding;
          // no alternative body or guessed range is searched.
          if (![plain, plain + "\r"].some(candidate => (!length || candidate.length === length) && nativeTextHash(candidate) === hash)) fail(sourceContext, "The native text context no longer matches the current source.");
        }
        const end = start + length, last = paragraphs[paragraphs.length - 1]!;
        if (end > last.start + last.text.length) unsupported(textMoniker, "A range ending after the final native paragraph mark is not supported.");
        for (const p of paragraphs) {
          const pEnd = p.start + p.text.length;
          if (end < p.start || start > pEnd) continue;
          const from = Math.max(0, start - p.start), to = Math.min(p.text.length, end - p.start);
          if (from > to) continue;
          ranges.push({ nodeId: p.nodeId, from, to, quote: p.text.slice(from, to) });
        }
        if (!ranges.length) fail(textMoniker, "A native text range is outside its source.");
      }
    }
    if (ranges.length && nodes.length) unsupported(lists[0]!, "Mixed text and object anchors are not supported.");
    if (ranges.length) {
      ranges.sort((a, b) => Number(a.nodeId.slice(5)) - Number(b.nodeId.slice(5)) || a.from - b.from);
      return ranges.length === 1 ? { kind: "text", ...ranges[0]! } : { kind: "text-set", ranges };
    }
    return nodes.length === 1 ? { kind: "node", nodeId: nodes[0]! } : { kind: "node-set", nodeIds: nodes };
  }
}
