import { z } from "zod";
import { IdentifierSchema } from "./contracts/primitives.js";
import { PresentationSizeSchema } from "./contracts/presentation.js";
import type { PresentationDocumentMetadata, PresentationMetadata, PresentationGeometry } from "./contracts/presentation.js";
import { DocumentError } from "./errors.js";
import { openDocument } from "./document-model.js";
import { encode } from "./document/package.js";
import type { DocumentModel, DocumentPackage, OpenDocumentOptions } from "./runtime.js";
import { checkLimit, freeze, readLimits } from "./document/validation.js";
import { presentationTransforms } from "./session/presentation-geometry.js";

/** Resolve source-local geometry through the declared placeholder inheritance chain. */
export function resolvePresentationGeometry(model: DocumentModel, nodeId: string): Readonly<PresentationGeometry> {
  const p=model.snapshot().metadata.presentation;
  if(!p) throw new DocumentError("UnsupportedFormat","Presentation geometry requires a presentation model.");
  if(!p.objects.some(object=>object.nodeId===nodeId)) throw new DocumentError("InvalidOperation","Geometry resolution requires a registered object.");
  return freeze(presentationTransforms(p,{peekNode:id=>model.getNode(id)}).geometry(nodeId));
}

const optionsSchema = z.strictObject({ documentId: IdentifierSchema, title: z.string().optional(), locale: IdentifierSchema.optional(), size: PresentationSizeSchema.optional() });
export interface CreatePresentationOptions {
  readonly documentId: string; readonly title?: string; readonly locale?: string;
  readonly size?: { readonly unit: "pt"; readonly width: number; readonly height: number };
}
const escape = (value: string): string => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
/** Create a validated, isolated initial package without platform or model I/O. */
export function createPresentation(input: CreatePresentationOptions, readOptions: OpenDocumentOptions = {}): DocumentPackage {
  const parsed = optionsSchema.safeParse(input);
  if (!parsed.success) throw new DocumentError("InvalidDocument", "Invalid new presentation options.");
  const options = parsed.data, size = options.size ?? { unit: "pt" as const, width: 960, height: 540 };
  const limits = readLimits(readOptions.limits);
  checkLimit((options.title?.length ?? 0) * 6 + (options.locale?.length ?? 0) * 6 + options.documentId.length * 6, limits.maxFileBytes);
  for (const value of [options.title, options.locale]) if (value !== undefined) {
    for (const character of value) {
      const point = character.codePointAt(0)!;
      if (point >= 0xd800 && point <= 0xdfff || point === 0) throw new DocumentError("InvalidDocument", "Presentation text must be well-formed Unicode without null characters.");
    }
  }
  const objects: PresentationMetadata["objects"] = [];
  const placeholder = (id: string, type: "title" | "body" | "ctrTitle" | "subTitle", index: number, x: number, y: number, width: number, height: number): string => {
    objects.push({ nodeId: id, kind: "text", placeholder: { type, index }, text: { autoFit: "none", wrap: true } });
    return '<div data-node-id="' + id + '" data-office-kind="text" data-office-x="' + x + '" data-office-y="' + y + '" data-office-width="' + width + '" data-office-height="' + height + '"><p data-node-id="' + id + '-p"></p></div>';
  };
  const width = size.width, height = size.height;
  const titleLayout = placeholder("title-centered", "ctrTitle", 0, width * .075, height / 3, width * .85, height * .22)
    + placeholder("subtitle", "subTitle", 1, width * .075, height * .6, width * .85, height * .18);
  const bodyLayout = placeholder("layout-title", "title", 0, width * .05, height * .05, width * .9, height * .15)
    + placeholder("layout-body", "body", 1, width * .05, height * .25, width * .9, height * .65);
  const metadata: PresentationDocumentMetadata = { title: options.title ?? "未命名演示文稿", locale: options.locale ?? "zh-CN", presentation: {
    size, commentFormat: "modern", slides: [{ slideId: "slide-1", rootNodeId: "slide-root-1", layoutId: "layout-blank" }],
    themes: [{ themeId: "theme-default", cssClass: "theme-default" }],
    masters: [{ masterId: "master-default", rootNodeId: "master-root-default", themeId: "theme-default" }],
    layouts: [{ layoutId: "layout-blank", rootNodeId: "layout-root-blank", masterId: "master-default", name: "空白" },
      { layoutId: "layout-title", rootNodeId: "layout-root-title", masterId: "master-default", name: "标题" },
      { layoutId: "layout-body", rootNodeId: "layout-root-body", masterId: "master-default", name: "标题与正文" }],
    notesMasters: [], stories: [], objects, charts: [], dataSets: [], comments: [], hyperlinks: [],
  } };
  const surface = (id: string, kind: string, content = ""): string => '<section data-node-id="' + id + '" data-office-kind="' + kind + '" data-office-story="' + kind + '">' + content + '</section>';
  const source = '<!doctype html><html lang="' + escape(metadata.locale!) + '"><head><meta charset="utf-8"><title>' + escape(metadata.title!)
    + '</title><link rel="stylesheet" href="styles.css"></head><body><article data-node-id="presentation-root" data-office-kind="presentation" class="theme-default">'
    + surface("slide-root-1", "slide") + surface("master-root-default", "master") + surface("layout-root-blank", "layout")
    + surface("layout-root-title", "layout", titleLayout) + surface("layout-root-body", "layout", bodyLayout) + '</article></body></html>';
  const files: DocumentPackage = new Map([
    ["manifest.json", encode(JSON.stringify({ format: "dagent-office", formatVersion: 1, documentId: options.documentId, kind: "presentation", revision: 0,
      entrypoint: "content.html", styles: ["styles.css"], metadata: "metadata.json", annotations: "annotations.json", assets: [] }))],
    ["content.html", encode(source)], ["metadata.json", encode(JSON.stringify(metadata))], ["annotations.json", encode("[]")],
    ["styles.css", encode('.theme-default{--office-theme-dk1:#000000;--office-theme-lt1:#ffffff;--office-theme-dk2:#44546a;--office-theme-lt2:#e7e6e6;--office-theme-accent1:#4472c4;--office-theme-accent2:#ed7d31;--office-theme-accent3:#a5a5a5;--office-theme-accent4:#ffc000;--office-theme-accent5:#5b9bd5;--office-theme-accent6:#70ad47;--office-theme-hlink:#0563c1;--office-theme-folHlink:#954f72;--office-major-ascii:"Liberation Sans";--office-major-east-asia:"Noto Sans CJK SC";--office-minor-ascii:"Liberation Sans";--office-minor-east-asia:"Noto Sans CJK SC";color:var(--office-theme-dk1);font-family:var(--office-minor-ascii),var(--office-minor-east-asia);font-size:18pt}p{margin:0}')],
  ]);
  return openDocument(files, readOptions).serialize();
}
