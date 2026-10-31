import type { Annotation, AnnotationReply, DeepReadonly, TextRange } from "@dagent-office/core";
import { children, sourceText } from "../export/source.js";
import type { SourceElement } from "../export/source.js";
import { attributes, declaration, empty, escapeXml } from "../export/xml.js";
import { nativeTextHash } from "../pptx/comments.js";
import { mainPart, nativeNamespaces, relativePart } from "./context.js";
import type { NativeSurface, PresentationExportContext } from "./context.js";
import type { NativeTextBody } from "./text.js";
import { PresentationTextWriter } from "./text.js";
import { emu, integer, unsupported } from "./values.js";
const MODERN = "http://schemas.microsoft.com/office/powerpoint/2018/8/main", PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command", AC = "http://schemas.microsoft.com/office/drawing/2013/main/command", REL = "http://schemas.microsoft.com/office/2018/10/relationships/", EXT = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}";
type Comment = DeepReadonly<Annotation>;
type Author = {
    id: string;
    name: string;
    initials?: string | undefined;
    userId: string;
    providerId: string;
};
export class PresentationCommentWriter {
    readonly extensions: Map<string, string> = new Map<string, string>();
    private readonly authors = new Map<string, Map<string, Author>>();
    private readonly indices = new Map<string, number>();
    constructor(private readonly ctx: PresentationExportContext, private readonly text: PresentationTextWriter) { }
    private loss(annotation: Comment, kind: "anchor" | "thread" | "status" | "format" | "metadata", message: string): void {
        if (!this.ctx.annotationLosses.some(loss => loss.annotationId === annotation.annotationId && loss.kind === kind && loss.message === message))
            this.ctx.annotationLosses.push({ annotationId: annotation.annotationId, kind, message });
    }
    private date(annotation: Comment, value: string | undefined): string {
        if (value !== undefined) {
            if (!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)?$/.test(value) || !Number.isFinite(Date.parse(value)))
                unsupported("A comment date cannot be represented as XML dateTime.");
            return value;
        }
        this.loss(annotation, "metadata", "A missing comment timestamp is represented by the neutral epoch required by the native format.");
        return "1970-01-01T00:00:00Z";
    }
    private author(annotation: Comment, format: "classic" | "modern", reply?: DeepReadonly<AnnotationReply>): Author {
        const metadata = this.ctx.p.comments.find(c => c.annotationId === annotation.annotationId), authorId = reply ? metadata?.replyAuthors?.find(author => author.replyId === reply.replyId)?.authorId : metadata?.authorId, source = this.ctx.p.commentAuthors?.find(author => author.authorId === authorId), display = reply?.author ?? (!reply ? annotation.author : undefined), key = JSON.stringify([source?.authorId ?? null, display?.name ?? "", display?.initials ?? ""]);
        let authors = this.authors.get(format);
        if (!authors) {
            authors = new Map();
            this.authors.set(format, authors);
        }
        let author = authors.get(key);
        if (!author) {
            author = { id: format === "modern" ? this.ctx.guid(source?.format === "modern" ? source.nativeId : undefined) : String(authors.size), name: display?.name ?? "", initials: display?.initials, userId: source?.userId ?? "", providerId: source?.providerId ?? "" };
            authors.set(key, author);
        }
        return author;
    }
    private slide(annotation: Comment): NativeSurface | undefined {
        const body = annotation.bodyNodeId ? this.ctx.p.stories.find(story => story.rootNodeId === annotation.bodyNodeId) : undefined, anchor = annotation.anchor;
        const ids = anchor.kind === "region" ? [anchor.surfaceId] : anchor.kind === "node" || anchor.kind === "text" ? [anchor.nodeId] : anchor.kind === "node-set" ? anchor.nodeIds : anchor.ranges.map(range => range.nodeId), surfaces = ids.flatMap(id => { const node = this.ctx.tree.byId.get(id), surface = node && this.ctx.surfaceFor(node); return surface?.kind === "slide" ? [surface] : []; });
        const surface = surfaces[0];
        if (surface && surfaces.every(other => other === surface))
            return surface;
        return body?.kind === "comment" && body.slideId ? this.ctx.surfaces.find(surface => surface.kind === "slide" && surface.definitionId === body.slideId) : undefined;
    }
    private drawingPath(node: SourceElement, slide: NativeSurface): string | undefined {
        const path: SourceElement[] = [];
        let current: SourceElement | undefined = node;
        while (current && current !== slide.node) {
            if (this.ctx.objects.has(current.attrs["data-node-id"]!))
                path.unshift(current);
            current = current.parent;
        }
        if (current !== slide.node || !path.length)
            return;
        return path.map(node => { const id = node.attrs["data-node-id"]!, object = this.ctx.objects.get(id)!, name = ({ text: "spMk", geometry: "spMk", group: "grpSpMk", image: "picMk", connector: "cxnSpMk", table: "graphicFrameMk", chart: "graphicFrameMk" }[object.kind]); return empty("ac:" + name, { id: this.ctx.nativeIds.get(id)!, creationId: this.ctx.creationIds.get(id)! }); }).join("");
    }
    private textRange(range: DeepReadonly<TextRange>, slide: NativeSurface, prefix: string): string | undefined {
        const node = this.ctx.tree.byId.get(range.nodeId);
        if (!node)
            return;
        let body: NativeTextBody | undefined;
        for (let current: SourceElement | undefined = node; current; current = current.parent) {
            body = this.text.bodies.get(current.attrs["data-node-id"]!);
            if (body)
                break;
        }
        if (!body || !this.ctx.objects.has(body.node.attrs["data-node-id"]!))
            return;
        const path = this.drawingPath(body.node, slide);
        if (!path)
            return;
        const absoluteFrom = node.from + range.from, absoluteTo = node.from + range.to, position = (offset: number, end: boolean): number | undefined => { const matching = body!.paragraphs.filter(p => offset >= p.node.from && offset <= p.node.to); const p = end ? matching.at(-1) : matching[0]; return p ? p.start + offset - p.node.from : undefined; }, from = position(absoluteFrom, false), to = position(absoluteTo, true);
        if (from === undefined || to === undefined || from > to || sourceText(node).slice(range.from, range.to) !== range.quote)
            return;
        return "<ac:txMkLst>" + prefix + path + "<ac:txMk" + attributes({ cp: integer(from, 1, 0, 2147483647), len: integer(to - from, 1, 0, 2147483647) }) + ">" + empty("ac:context", { len: body.text.length, hash: nativeTextHash(body.text) }) + "</ac:txMk></ac:txMkLst>";
    }
    private anchor(annotation: Comment, slide: NativeSurface): string {
        const anchor = annotation.anchor, prefix = "<pc:docMk/>" + empty("pc:sldMk", { sldId: slide.nativeSlideId! });
        let result: string | undefined;
        if (annotation.status !== "needs_reanchor") {
            if (anchor.kind === "text" || anchor.kind === "text-set") {
                const ranges = anchor.kind === "text" ? [anchor] : anchor.ranges, mapped = ranges.map(range => this.textRange(range, slide, prefix));
                if (mapped.every(Boolean))
                    result = mapped.join("");
            }
            else if (anchor.kind === "node" || anchor.kind === "node-set") {
                const ids = anchor.kind === "node" ? [anchor.nodeId] : anchor.nodeIds, items = ids.map(id => {
                    const node = this.ctx.tree.byId.get(id);
                    if (node === slide.node)
                        return "<pc:sldMkLst>" + prefix + "</pc:sldMkLst>";
                    const path = node && this.ctx.objects.has(id) && this.drawingPath(node, slide);
                    return path ? "<ac:deMkLst>" + prefix + path + "</ac:deMkLst>" : undefined;
                });
                if (items.every(Boolean) && (ids.length === 1 || !ids.includes(slide.rootNodeId)))
                    result = items.join("");
            }
        }
        if (result)
            return result;
        this.loss(annotation, "anchor", "The native modern format cannot preserve this source anchor; the complete thread is retained with an unknown anchor for recovery.");
        return "<p188:unknownAnchor/>";
    }
    private body(annotation: Comment, slide: NativeSurface, part: string, reply?: DeepReadonly<AnnotationReply>): string {
        const value = reply ?? annotation, nodeId = value.bodyNodeId;
        return nodeId ? this.text.body(this.ctx.tree.byId.get(nodeId)!, slide, "p188:txBody", part) : this.text.plain(value.instruction!);
    }
    private plain(annotation: Comment): string {
        const value = (body: Comment | DeepReadonly<AnnotationReply>): string => body.bodyNodeId ? children(this.ctx.tree.byId.get(body.bodyNodeId)!).map(sourceText).join("\n") : body.instruction!;
        return [value(annotation), ...(annotation.replies ?? []).map(reply => (reply.author?.name ? reply.author.name + ": " : "") + value(reply))].join("\n\n");
    }
    write(): void {
        const groups = new Map<NativeSurface, {
            classic: Comment[];
            modern: Comment[];
        }>();
        for (const annotation of this.ctx.snapshot.annotations) {
            const slide = this.slide(annotation);
            if (!slide) {
                this.ctx.unexportedAnnotations.push({ annotationId: annotation.annotationId, reason: "missing_surface" });
                continue;
            }
            const requested = this.ctx.options.commentFormat ?? "preserve", format = requested === "preserve" ? this.ctx.p.comments.find(comment => comment.annotationId === annotation.annotationId)?.format ?? this.ctx.p.commentFormat : requested;
            let group = groups.get(slide);
            if (!group) {
                group = { classic: [], modern: [] };
                groups.set(slide, group);
            }
            group[format].push(annotation);
        }
        let sequence = 0;
        for (const [slide, group] of groups)
            for (const format of ["classic", "modern"] as const) {
                if (!group[format].length)
                    continue;
                const modern = format === "modern", path = "ppt/comments/" + (modern ? "modern" : "classic") + (++sequence) + ".xml", out = this.ctx.opc.part(path, modern ? "application/vnd.ms-powerpoint.comments+xml" : "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"), relation = this.ctx.opc.relate(slide.part, modern ? REL + "comments" : "comments", relativePart(slide.part, path));
                if (modern)
                    this.extensions.set(slide.part, "<p:extLst><p:ext" + attributes({ uri: EXT }) + ">" + empty("p188:commentRel", { "xmlns:p188": MODERN, "r:id": relation }) + "</p:ext></p:extLst>");
                out.add(declaration + "<" + (modern ? "p188" : "p") + ":cmLst" + nativeNamespaces + (modern ? ' xmlns:p188="' + MODERN + '" xmlns:pc="' + PC + '" xmlns:ac="' + AC + '"' : "") + ">");
                for (const annotation of group[format]) {
                    const author = this.author(annotation, format), metadata = this.ctx.p.comments.find(comment => comment.annotationId === annotation.annotationId);
                    this.ctx.exportedAnnotationIds.push(annotation.annotationId);
                    if (modern) {
                        const status = annotation.status === "needs_reanchor" ? metadata?.recoveryStatus ?? "open" : annotation.status;
                        out.add("<p188:cm" + attributes({ id: this.ctx.guid(annotation.sourceId), authorId: author.id, created: this.date(annotation, annotation.date), status: status === "open" ? "active" : status }) + ">" + this.anchor(annotation, slide));
                        const body = annotation.bodyNodeId ? this.ctx.tree.byId.get(annotation.bodyNodeId) : undefined;
                        if (body?.attrs["data-office-comment-x"] !== undefined && body.attrs["data-office-comment-y"] !== undefined)
                            out.add(empty("p188:pos", { x: emu(Number(body.attrs["data-office-comment-x"])), y: emu(Number(body.attrs["data-office-comment-y"])) }));
                        if (annotation.replies?.length) {
                            out.add("<p188:replyLst>");
                            for (const reply of annotation.replies) {
                                const replyAuthor = this.author(annotation, format, reply);
                                out.add("<p188:reply" + attributes({ id: this.ctx.guid(reply.sourceId), authorId: replyAuthor.id, created: this.date(annotation, reply.date), status: reply.status === "resolved" || reply.status === "closed" ? reply.status : "active" }) + ">" + this.body(annotation, slide, path, reply) + "</p188:reply>");
                            }
                            out.add("</p188:replyLst>");
                        }
                        out.add(this.body(annotation, slide, path) + "</p188:cm>");
                    }
                    else {
                        const index = (this.indices.get(author.id) ?? 0) + 1;
                        this.indices.set(author.id, index);
                        const anchor = annotation.anchor, classicPoint = anchor.kind === "region" && anchor.layoutKey === "pptx-classic-point";
                        let x = 0, y = 0;
                        if (anchor.kind === "region") {
                            x = anchor.rect.x;
                            y = anchor.rect.y;
                        }
                        if (!classicPoint)
                            this.loss(annotation, "anchor", "Classic comments retain a page point; precise text, object or region anchors cannot be represented.");
                        if (annotation.status !== "open")
                            this.loss(annotation, "status", "Classic comments do not preserve resolved, closed or recovery status.");
                        if (annotation.replies?.length)
                            this.loss(annotation, "thread", "Replies are flattened into the root text; reply identities, timestamps and workflow status are not retained.");
                        const bodies = [annotation, ...annotation.replies ?? []].flatMap(message => message.bodyNodeId ? [this.ctx.tree.byId.get(message.bodyNodeId)!] : []);
                        if (bodies.some(body => children(body).some(node => node.tag !== "p" || Boolean(node.attrs.style || node.attrs.class) || children(node).length > 0)))
                            this.loss(annotation, "format", "Classic comment bodies retain text and paragraph breaks; rich text formatting is not representable.");
                        out.add("<p:cm" + attributes({ authorId: author.id, idx: index, dt: this.date(annotation, annotation.date) }) + ">" + empty("p:pos", { x: integer(Math.max(0, x), 8, 0, 2147483647), y: integer(Math.max(0, y), 8, 0, 2147483647) }) + "<p:text>" + escapeXml(this.plain(annotation)) + "</p:text></p:cm>");
                    }
                }
                out.add("</" + (modern ? "p188" : "p") + ":cmLst>");
                out.finish();
            }
        for (const [format, authors] of this.authors) {
            const modern = format === "modern", path = "ppt/" + (modern ? "authors" : "commentAuthors") + ".xml", out = this.ctx.opc.part(path, modern ? "application/vnd.ms-powerpoint.authors+xml" : "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml");
            this.ctx.opc.relate(mainPart, modern ? REL + "authors" : "commentAuthors", relativePart(mainPart, path));
            out.add(declaration + (modern ? '<p188:authorLst xmlns:p188="' + MODERN + '">' : '<p:cmAuthorLst' + nativeNamespaces + '>'));
            for (const author of authors.values())
                out.add(empty(modern ? "p188:author" : "p:cmAuthor", { id: author.id, name: author.name, initials: modern ? author.initials : author.initials ?? "", ...(modern ? { userId: author.userId, providerId: author.providerId } : { lastIdx: this.indices.get(author.id) ?? 0, clrIdx: Number(author.id) }) }));
            out.add(modern ? "</p188:authorLst>" : "</p:cmAuthorLst>");
            out.finish();
        }
    }
}
