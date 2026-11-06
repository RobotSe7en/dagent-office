import { createAssetDescriptor, DocumentError, WorkbookPresetSchema } from "@dagent-office/core";
import type { WorkbookManifest, WorkbookObject, WorkbookObjectHyperlink } from "@dagent-office/core";
import type { OpcPackage } from "../ooxml/package.js";
import { A, attr, C, check, elements, fail, R, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Html } from "../drawingml/html.js";
import { attributes, boolean, integer, one, value } from "./values.js";
import { drawingCoordinate, readDrawingAnchor, XDR } from "./drawing-anchor.js";
import { drawingAppearance } from "./drawing-style.js";
import { readDrawingText } from "./drawing-text.js";
import type { XlsxStyles } from "./styles.js";
import type { WorkbookContent } from "./content.js";
import { WorkbookCharts } from "./chart.js";

function transform(node: XmlNode | undefined, target: Html): void {
  if (!node) return;
  attributes(node, ["rot", "flipH", "flipV"]); check(node, ["off", "ext"], A);
  // The worksheet anchor is authoritative. Validate native xfrm display caches
  // without storing a second competing position/size in the document source.
  for (const [tag, fields] of [["off", ["x", "y"]], ["ext", ["cx", "cy"]]] as const) {
    const item = one(node, tag, A); if (!item) continue; attributes(item, fields); check(item, [], A);
    for (const name of fields) drawingCoordinate(value(item, name), item, tag === "ext");
  }
  if (value(node, "rot") !== undefined) target.attrs["data-office-rotation"] = String(integer(node, "rot", -2147483648, 2147483647) / 60000);
  for (const [native, name] of [["flipH", "flip-h"], ["flipV", "flip-v"]]) if (value(node, native!) !== undefined) target.attrs[`data-office-${name}`] = String(boolean(node, native!));
}
function preset(properties: XmlNode, target: Html, picture = false): string {
  const geometry = one(properties, "prstGeom", A);
  if (!geometry) { if (picture) return "rect"; unsupported(properties, "A worksheet shape requires supported preset geometry."); }
  attributes(geometry, ["prst"]); check(geometry, ["avLst"], A);
  const kind = value(geometry, "prst");
  if (!kind || !WorkbookPresetSchema.safeParse(kind).success || picture && kind !== "rect") unsupported(geometry, "This worksheet preset is not supported.");
  const list = one(geometry, "avLst", A), names = new Set<string>();
  if (list) { attributes(list, []); check(list, ["gd"], A); }
  for (const adjustment of elements(list, "gd", A)) {
    attributes(adjustment, ["name", "fmla"]); check(adjustment, [], A);
    const name = value(adjustment, "name"), formula = value(adjustment, "fmla");
    if (!name || !/^[a-z][a-z0-9]*$/i.test(name) || names.has(name.toLowerCase()) || !formula || !/^val -?\d+$/.test(formula) || !Number.isSafeInteger(Number(formula.slice(4)))) unsupported(adjustment, "Preset adjustments require unique integer constants.");
    if (picture) unsupported(adjustment, "Picture preset adjustments are not supported.");
    names.add(name.toLowerCase()); target.attrs[`data-office-adjust-${name.toLowerCase()}`] = formula.slice(4);
  }
  return kind;
}
const shapeLocks = ["noGrp", "noSelect", "noRot", "noChangeAspect", "noMove", "noResize", "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType", "noTextEdit"] as const;
function locks(source: XmlNode | undefined, kind: "shape" | "image" | "chart"): WorkbookObject["locks"] {
  if (!source) return undefined;
  const allowed = kind === "chart" ? ["noGrp", "noDrilldown", "noSelect", "noChangeAspect", "noMove", "noResize"] : kind === "image" ? [...shapeLocks.filter(name => name !== "noTextEdit"), "noCrop"] : shapeLocks;
  attributes(source, allowed); check(source, [], A);
  return Object.fromEntries(allowed.filter(name => value(source, name) !== undefined).map(name => [name, boolean(source, name)]));
}
export class WorkbookDrawings {
  readonly objects: WorkbookObject[] = [];
  readonly hyperlinks: WorkbookObjectHyperlink[] = [];
  readonly assets: WorkbookManifest["assets"] = [];
  readonly files: Map<string, Uint8Array> = new Map();
  readonly parts: Set<string> = new Set();
  readonly charts: WorkbookCharts;
  private readonly sheets = new Set<string>();
  private readonly nativeIds = new Map<string, Set<string>>();
  private assetBytes = 0;
  constructor(readonly opc: OpcPackage, readonly styles: XlsxStyles, readonly content: WorkbookContent) { this.charts = new WorkbookCharts(opc, styles, content); }
  private hyperlink = (node: XmlNode | undefined, target: Html, part: string): void => {
    if (!node) return;
    for (const key of node.attrs.keys()) if (!key.startsWith("http://www.w3.org/2000/xmlns/ ") && ![R + " id", " tooltip", " history", " highlightClick", " endSnd", " action", " invalidUrl", " tgtFrame"].includes(key)) unsupported(node);
    check(node, [], A);
    if (["action", "invalidUrl", "tgtFrame"].some(name => value(node, name))) unsupported(node, "Active actions, invalid targets and named hyperlink frames are not supported.");
    for (const name of ["history", "highlightClick", "endSnd"]) boolean(node, name);
    const id = attr(node, "id", R), relation = id ? this.opc.relationships(part).get(id) : undefined;
    if (!relation || relation.type !== R + "/hyperlink" || !relation.external) fail(node, "A drawing hyperlink relationship is invalid.");
    const destination = relation.target;
    if (!destination.startsWith("#") && (!/^(https?:\/\/|mailto:)/i.test(destination) || /[\u0000-\u0020\u007f]/.test(destination))) unsupported(node, "This drawing hyperlink target is not supported.");
    if (destination.length > 8192 || (value(node, "tooltip")?.length ?? 0) > 32767) throw new DocumentError("ResourceLimitExceeded", "A drawing hyperlink exceeds its text budget.");
    this.hyperlinks.push({ nodeId: target.attrs["data-node-id"]!, ...(destination.startsWith("#") ? { location: destination.slice(1) } : { url: destination }), ...(value(node, "tooltip") === undefined ? {} : { tooltip: value(node, "tooltip")! }) });
  };
  private image(source: XmlNode, target: Html, description: string): void {
    const fill = one(source, "blipFill", XDR); if (!fill) fail(source, "A worksheet picture has no source.");
    attributes(fill, ["dpi", "rotWithShape"]); check(fill, ["blip", "srcRect", "stretch"], A);
    if (value(fill, "dpi") !== undefined) integer(fill, "dpi");
    const rotates = boolean(fill, "rotWithShape", true);
    if (value(fill, "rotWithShape") !== undefined) target.attrs["data-office-image-rotate-with-shape"] = String(rotates);
    const blip = one(fill, "blip", A), stretch = one(fill, "stretch", A); if (!blip || !stretch) unsupported(fill, "A picture requires an embedded image with rectangular stretch.");
    for (const key of blip.attrs.keys()) if (!key.startsWith("http://www.w3.org/2000/xmlns/ ") && ![R + " embed", " cstate"].includes(key)) unsupported(blip, "External or unknown image references are not supported.");
    check(blip, [], A);
    const compression = value(blip, "cstate"); if (compression && !["email", "screen", "print", "hqprint", "none"].includes(compression)) fail(blip, "An image compression hint is invalid.");
    const id = attr(blip, "embed", R), relation = id ? this.opc.relationships(source.part).get(id) : undefined;
    if (!relation || relation.type !== R + "/image" || relation.external) fail(blip, "A worksheet image relationship is missing or invalid.");
    const data = this.opc.files.get(relation.target), mime = this.opc.contentType(relation.target), path = "assets/" + relation.target;
    if (!data || !mime || !["image/png", "image/jpeg", "image/webp"].includes(mime)) unsupported(blip, "This image container is not supported.");
    if (!this.files.has(path)) {
      if (data.byteLength > this.content.limits.maxFileBytes || this.assetBytes + data.byteLength > this.content.limits.maxPackageBytes || this.files.size + 1 > this.content.limits.maxFiles) throw new DocumentError("ResourceLimitExceeded", "Drawing assets exceed their source budget.");
      this.assetBytes += data.byteLength; this.assets.push(createAssetDescriptor(path, mime, data)); this.files.set(path, data); this.parts.add(relation.target);
    }
    const crop = one(fill, "srcRect", A);
    if (crop) {
      attributes(crop, ["l", "t", "r", "b"]); check(crop, [], A);
      for (const [tag, name] of [["l", "left"], ["t", "top"], ["r", "right"], ["b", "bottom"]]) target.attrs[`data-office-crop-${name}`] = String(integer(crop, tag!, 0, 99999, 0) / 100000);
    }
    attributes(stretch, []); check(stretch, ["fillRect"], A);
    const rectangle = one(stretch, "fillRect", A);
    if (rectangle) { attributes(rectangle, ["l", "t", "r", "b"]); check(rectangle, [], A); for (const name of ["l", "t", "r", "b"]) if (integer(rectangle, name, -2147483648, 2147483647, 0)) unsupported(rectangle, "Offset image stretch rectangles are not supported."); }
    target.children.push(this.content.make("img", { src: path, alt: description }));
  }
  sheet(sheetId: string, root: XmlNode): void {
    if (this.sheets.has(sheetId)) fail(root, "Worksheet drawings may only be consumed once."); this.sheets.add(sheetId);
    const reference = one(root, "drawing"); if (!reference) return;
    attributes(reference, ["id"], R); check(reference, [], XDR);
    const id = attr(reference, "id", R), relation = id ? this.opc.relationships(root.part).get(id) : undefined;
    if (!relation || relation.type !== R + "/drawing" || relation.external || this.parts.has(relation.target)) fail(reference, "A worksheet drawing relationship is invalid or shared between worksheets.");
    if (this.opc.contentType(relation.target) !== "application/vnd.openxmlformats-officedocument.drawing+xml") fail(reference, "A worksheet drawing part has the wrong content type.");
    const drawing = this.opc.read(relation.target); this.parts.add(relation.target);
    if (drawing.name !== "wsDr" || drawing.uri !== XDR) fail(drawing, "A worksheet drawing part has an invalid root.");
    attributes(drawing, []); check(drawing, ["absoluteAnchor", "oneCellAnchor", "twoCellAnchor"], XDR);
    for (const anchor of elements(drawing, undefined, XDR)) this.object(sheetId, anchor);
  }
  private object(sheetId: string, anchorNode: XmlNode): void {
    const native = readDrawingAnchor(anchorNode), source = native.object, picture = source.name === "pic", frame = source.name === "graphicFrame";
    if (!["pic", "sp", "graphicFrame"].includes(source.name)) unsupported(source, "This worksheet drawing object is not yet mapped.");
    attributes(source, picture || frame ? ["macro", "fPublished"] : ["macro", "textlink", "fLocksText", "fPublished"]);
    if (value(source, "macro") || value(source, "textlink")) unsupported(source, "Macros and linked shape text are not supported.");
    boolean(source, "fPublished"); if (!picture && !frame && !boolean(source, "fLocksText", true)) unsupported(source, "An unlocked native text-attachment flag is not supported.");
    const nvName = frame ? "nvGraphicFramePr" : picture ? "nvPicPr" : "nvSpPr", propName = frame ? "cNvGraphicFramePr" : picture ? "cNvPicPr" : "cNvSpPr";
    check(frame ? { ...source, content: source.content.filter(node => typeof node === "string" || node.uri !== A || node.name !== "graphic") } : source, frame ? [nvName, "xfrm"] : [nvName, "spPr", "style", picture ? "blipFill" : "txBody"], XDR);
    const nonvisual = one(source, nvName, XDR), identity = one(nonvisual, "cNvPr", XDR), app = one(nonvisual, propName, XDR), properties = one(source, frame ? "xfrm" : "spPr", XDR);
    if (!nonvisual || !identity || !app || !properties) fail(source, "A worksheet drawing identity or properties are missing.");
    attributes(nonvisual, []); check(nonvisual, ["cNvPr", propName], XDR);
    attributes(identity, ["id", "name", "descr", "hidden", "title"]); check(identity, ["hlinkClick"], A);
    const id = String(integer(identity, "id")), ids = this.nativeIds.get(sheetId) ?? new Set<string>();
    if (ids.has(id)) fail(identity, "Native drawing identities repeat within a worksheet."); ids.add(id); this.nativeIds.set(sheetId, ids);
    const name = value(identity, "name"); if (name === undefined) fail(identity, "A drawing name is missing.");
    for (const property of ["name", "descr", "title"]) if ((value(identity, property)?.length ?? 0) > 32767) throw new DocumentError("ResourceLimitExceeded", "A drawing description exceeds its text budget.");
    const lockName = frame ? "graphicFrameLocks" : picture ? "picLocks" : "spLocks";
    attributes(app, frame ? [] : picture ? ["preferRelativeResize"] : ["txBox"]); check(app, [lockName], A);
    const lock = locks(one(app, lockName, A), frame ? "chart" : picture ? "image" : "shape"), description = value(identity, "descr") ?? "";
    const target = this.content.make("div", { ...(description ? { "data-office-alt": description } : {}), ...(value(identity, "title") === undefined ? {} : { title: value(identity, "title")! }) });
    if (value(identity, "hidden") !== undefined) target.attrs["data-office-hidden"] = String(boolean(identity, "hidden"));
    if (frame) transform(properties, target);
    else {
      attributes(properties, ["bwMode"]); check(properties, ["xfrm", "prstGeom", "noFill", "solidFill", "gradFill", "ln", "effectLst"], A);
      if (value(properties, "bwMode") && value(properties, "bwMode") !== "auto") unsupported(properties, "Explicit drawing black-and-white modes are not supported.");
      transform(one(properties, "xfrm", A), target);
      const classes = drawingAppearance(one(source, "style", XDR), properties, this.styles); if (classes) target.attrs.class = classes;
    }
    this.hyperlink(one(identity, "hlinkClick", A), target, source.part);
    const base = { nodeId: target.attrs["data-node-id"]!, sheetId, anchor: native.anchor, nativeId: id, name, locked: native.locked, printable: native.printable, ...(lock ? { locks: lock } : {}) };
    let object: WorkbookObject;
    if (frame) {
      const graphic = one(source, "graphic", A), data = one(graphic, "graphicData", A);
      if (!graphic || !data) fail(source, "A chart frame is missing graphic data."); attributes(graphic, []); check(graphic, ["graphicData"], A); attributes(data, ["uri"]);
      if (value(data, "uri") !== C) unsupported(data, "Only native basic chart frames are supported."); check(data, ["chart"], C);
      const reference = one(data, "chart", C); if (!reference) fail(data, "A chart frame has no chart reference.");
      object = { ...base, kind: "chart", chartId: this.charts.read(reference, target, this.hyperlink) };
    } else if (picture) {
      preset(properties, target, true); this.image(source, target, description);
      object = { ...base, kind: "image", ...(value(app, "preferRelativeResize") === undefined ? {} : { preferRelativeResize: boolean(app, "preferRelativeResize") }) };
    } else {
      const shape = preset(properties, target), textBox = boolean(app, "txBox");
      if (textBox && shape !== "rect") unsupported(app, "A nonrectangular text box is not supported.");
      const body = one(source, "txBody", XDR), text = body ? readDrawingText(body, target, this.content, this.styles, this.hyperlink) : undefined;
      object = textBox ? { ...base, kind: "text", ...(text ? { text } : {}) } : { ...base, kind: "geometry", preset: WorkbookPresetSchema.parse(shape), ...(text ? { text } : {}) };
    }
    target.attrs["data-office-kind"] = object.kind; this.objects.push(object); this.content.appendObject(target);
  }
}
