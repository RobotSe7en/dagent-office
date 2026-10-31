import { createOpaqueId, DocumentError } from "@dagent-office/core";
import type { DeepReadonly, DocumentModel, DocumentPackage, DocumentReadLimits, DocumentSnapshot, PresentationMetadata, PresentationObject } from "@dagent-office/core";
import { OutputOpc } from "../export/opc.js";
import { children, parseSource } from "../export/source.js";
import type { SourceElement, SourceTree } from "../export/source.js";
import { XmlParts } from "../export/xml.js";
import type { XmlPartBudget } from "../export/xml.js";
import type { XmlBudget } from "../ooxml/xml.js";
import { PresentationExportStyles } from "./styles.js";
import type { PresentationStyleSurface } from "./styles.js";
import type { PptxAnnotationLoss, PptxExportOptions, UnexportedPptxAnnotation } from "./types.js";
export interface NativeSurface extends PresentationStyleSurface {
    readonly kind: "slide" | "master" | "layout" | "notes-master" | "notes";
    readonly definitionId: string;
    readonly part: string;
    readonly node: SourceElement;
    readonly nativeSlideId?: number;
}
export const mainPart = "ppt/presentation.xml";
export const pmlType = (kind: string): string => "application/vnd.openxmlformats-officedocument.presentationml." + kind + "+xml";
export const nativeNamespaces = ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
export function relativePart(from: string, to: string): string {
    const origin = from.split("/").slice(0, -1), target = to.split("/");
    while (origin.length && origin[0] === target[0]) {
        origin.shift();
        target.shift();
    }
    return [...origin.map(() => ".."), ...target].join("/");
}
export function resourcePath(reference: string, entrypoint: string): string {
    const parts = entrypoint.split("/").slice(0, -1);
    for (const token of reference.split("/")) {
        const value = decodeURIComponent(token);
        if (value === "..")
            parts.pop();
        else if (value !== ".")
            parts.push(value);
    }
    return parts.join("/");
}
type FormatKind = "fill" | "line" | "effect" | "background";
export class PresentationExportContext {
    readonly snapshot: DocumentSnapshot;
    readonly p: DeepReadonly<PresentationMetadata>;
    readonly files: DocumentPackage;
    readonly tree: SourceTree;
    readonly styles: PresentationExportStyles;
    readonly budget: XmlPartBudget = { bytes: 0, nodes: 0, files: 0 };
    readonly xmlBudget: XmlBudget = { nodes: 0 };
    readonly opc: OutputOpc;
    readonly surfaces: readonly NativeSurface[];
    readonly objects: ReadonlyMap<string, DeepReadonly<PresentationObject>>;
    readonly nativeIds: Map<string, number> = new Map();
    readonly creationIds: Map<string, string> = new Map();
    readonly exportedAnnotationIds: string[] = [];
    readonly unexportedAnnotations: UnexportedPptxAnnotation[] = [];
    readonly annotationLosses: PptxAnnotationLoss[] = [];
    private readonly guids = new Set<string>();
    private readonly media = new Map<string, string>();
    private readonly formats = new Map<string, Map<FormatKind, string[]>>();
    private readonly encoder = new (globalThis as unknown as {
        TextEncoder: new () => {
            encode(value: string): Uint8Array;
        };
    }).TextEncoder();
    private generatedBytes = 0;
    private generatedNodes = 0;
    private readonly nextIds = new Map<string, number>();
    constructor(model: DocumentModel, readonly limits: DocumentReadLimits, readonly options: PptxExportOptions) {
        this.snapshot = model.snapshot();
        const p = this.snapshot.metadata.presentation;
        if (!p)
            throw new DocumentError("UnsupportedFormat", "PPTX export requires a presentation package.");
        this.p = p;
        this.files = model.serialize();
        this.tree = parseSource(this.snapshot.html, limits, "presentation");
        this.styles = new PresentationExportStyles(p, this.tree, this.snapshot.styles.map(style => style.css), limits);
        this.objects = new Map(p.objects.map(object => [object.nodeId, object]));
        this.opc = new OutputOpc(new XmlParts(limits, this.budget));
        const themes = new Map(p.masters.map(master => [master.masterId, master.themeId])), surfaces: NativeSurface[] = [];
        const append = (kind: NativeSurface["kind"], rootNodeId: string, definitionId: string, themeId: string, part: string, extra: Partial<PresentationStyleSurface> = {}, nativeSlideId?: number): void => {
            surfaces.push({ kind, rootNodeId, definitionId, themeId, part, node: this.tree.byId.get(rootNodeId)!, ...extra, ...(nativeSlideId === undefined ? {} : { nativeSlideId }) });
        };
        for (const [index, master] of p.masters.entries())
            append("master", master.rootNodeId, master.masterId, master.themeId, "ppt/slideMasters/slideMaster" + (index + 1) + ".xml", { master });
        for (const [index, layout] of p.layouts.entries())
            append("layout", layout.rootNodeId, layout.layoutId, themes.get(layout.masterId)!, "ppt/slideLayouts/slideLayout" + (index + 1) + ".xml", { master: p.masters.find(master => master.masterId === layout.masterId)!, layout });
        for (const [index, master] of p.notesMasters.entries())
            append("notes-master", master.rootNodeId, master.notesMasterId, master.themeId, "ppt/notesMasters/notesMaster" + (index + 1) + ".xml", { notesMaster: master });
        const orderedSlides = children(this.tree.root).flatMap(node => { const slide = p.slides.find(slide => slide.rootNodeId === node.attrs["data-node-id"]); return slide ? [slide] : []; });
        for (const [index, slide] of orderedSlides.entries()) {
            const layout = p.layouts.find(layout => layout.layoutId === slide.layoutId)!, master = p.masters.find(master => master.masterId === layout.masterId)!;
            append("slide", slide.rootNodeId, slide.slideId, master.themeId, "ppt/slides/slide" + (index + 1) + ".xml", { master, layout }, 256 + index);
        }
        for (const story of p.stories)
            if (story.kind === "notes") {
                const master = p.notesMasters.find(master => master.notesMasterId === story.notesMasterId);
                append("notes", story.rootNodeId, story.storyId, master?.themeId ?? surfaces.find(surface => surface.kind === "slide" && surface.definitionId === story.slideId)!.themeId, "ppt/notesSlides/notesSlide" + (surfaces.filter(surface => surface.kind === "notes").length + 1) + ".xml", master ? { notesMaster: master } : {});
            }
        this.surfaces = surfaces;
        for (const surface of surfaces) {
            let id = 1;
            const visit = (node: SourceElement): void => {
                if (this.objects.has(node.attrs["data-node-id"]!)) {
                    const source = node.attrs["data-node-id"]!;
                    this.nativeIds.set(source, ++id);
                    this.creationIds.set(source, this.guid(this.objects.get(source)!.creationId));
                }
                for (const child of children(node))
                    visit(child);
            };
            visit(surface.node);
            this.nextIds.set(surface.rootNodeId, id + 1);
        }
    }
    guid(preferred?: string): string {
        let id = preferred?.toUpperCase();
        if (!id || !/^\{[\dA-F]{8}-(?:[\dA-F]{4}-){3}[\dA-F]{12}\}$/.test(id) || this.guids.has(id))
            do {
                id = "{" + createOpaqueId().toUpperCase() + "}";
            } while (this.guids.has(id));
        this.guids.add(id);
        return id;
    }
    surfaceFor(node: SourceElement): NativeSurface | undefined {
        for (let current: SourceElement | undefined = node; current; current = current.parent) {
            const surface = this.surfaces.find(surface => surface.rootNodeId === current!.attrs["data-node-id"]);
            if (surface)
                return surface;
        }
        return;
    }
    extraId(surface: NativeSurface): number { const value = this.nextIds.get(surface.rootNodeId)!; this.nextIds.set(surface.rootNodeId, value + 1); return value; }
    themePart(themeId: string): string {
        const index = this.p.themes.findIndex(theme => theme.themeId === themeId);
        if (index < 0)
            throw new DocumentError("InvalidDocument", "An exported theme is missing.");
        return "ppt/theme/theme" + (index + 1) + ".xml";
    }
    /** Reserve retained drawing and unique theme fragments before composing
     * complete parts. Final XmlParts checks still count the exact full package. */
    checkGenerated(source: string): void {
        this.generatedBytes += this.encoder.encode(source).length;
        this.generatedNodes += (source.match(/<(?![!?/])/g) ?? []).length;
        if (this.generatedBytes + this.budget.bytes > this.limits.maxPackageBytes || this.generatedNodes + this.budget.nodes > this.limits.maxNodes)
            throw new DocumentError("ResourceLimitExceeded", "Generated PPTX fragments exceed the shared output budget.");
    }
    format(themeId: string, kind: FormatKind, source: string): number {
        let kinds = this.formats.get(themeId);
        if (!kinds) {
            kinds = new Map();
            this.formats.set(themeId, kinds);
        }
        let values = kinds.get(kind);
        if (!values) {
            values = [];
            kinds.set(kind, values);
        }
        let index = values.indexOf(source);
        if (index < 0) {
            if (kind === "fill" && values.length >= 999)
                throw new DocumentError("ResourceLimitExceeded", "The native fill style index range is exhausted.");
            this.checkGenerated(source);
            values.push(source);
            index = values.length - 1;
        }
        return index + 1 + (kind === "background" ? 1000 : 0);
    }
    formatValues(themeId: string, kind: FormatKind): readonly string[] { return this.formats.get(themeId)?.get(kind) ?? []; }
    image(node: SourceElement, part: string): string {
        const source = resourcePath(node.attrs.src!, this.snapshot.manifest.entrypoint), descriptor = this.snapshot.manifest.assets.find(asset => asset.path === source);
        if (!descriptor)
            throw new DocumentError("AssetUnavailable", "An exported presentation image is missing.");
        let target = this.media.get(source);
        if (!target) {
            const extension = ({ "image/png": "png", "image/jpeg": "jpeg", "image/webp": "webp" } as Readonly<Record<string, string>>)[descriptor.mimeType];
            if (!extension)
                throw new DocumentError("UnsupportedFeature", "This image container is outside the PPTX export profile.");
            target = "ppt/media/image" + (this.media.size + 1) + "." + extension;
            this.media.set(source, target);
            this.opc.binary(target, descriptor.mimeType, this.files.get(source)!);
        }
        return this.opc.relate(part, "image", relativePart(part, target));
    }
}
