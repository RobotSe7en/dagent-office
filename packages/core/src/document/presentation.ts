import { parse, walk, ident } from "css-tree";
import type { Annotation } from "../contracts/annotations.js";
import type { PresentationGeometry, PresentationMetadata, PresentationObject } from "../contracts/presentation.js";
import { PresentationGeometrySchema } from "../contracts/presentation.js";
import type { ParsedHtml } from "./html.js";
import type { DocumentReadLimits } from "../runtime.js";
import { checkLimit, invalid } from "./validation.js";

const geometryNames = ["x", "y", "width", "height"] as const;
const numeric = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i;
function number(value: string | undefined): number {
  if (value === undefined || !numeric.test(value) || !Number.isFinite(Number(value))) invalid("Presentation geometry must contain finite numbers.");
  return Number(value);
}
function boolean(value: string | undefined): boolean {
  if (value === undefined || value === "false") return false;
  if (value === "true") return true;
  return invalid("Presentation flips must be true or false.");
}
/** Source-local transform. CSS positioning and screen scale are never inputs. */
export function readPresentationGeometry(attributes: Readonly<Record<string, string>>): PresentationGeometry | undefined {
  if (!geometryNames.some((name) => attributes["data-office-" + name] !== undefined)) {
    if (["rotation", "flip-h", "flip-v"].some((name) => attributes["data-office-" + name] !== undefined)) invalid("A partial presentation transform is invalid.");
    return undefined;
  }
  const value = {
    x: number(attributes["data-office-x"]), y: number(attributes["data-office-y"]),
    width: number(attributes["data-office-width"]), height: number(attributes["data-office-height"]),
    rotation: attributes["data-office-rotation"] === undefined ? 0 : number(attributes["data-office-rotation"]),
    flipH: boolean(attributes["data-office-flip-h"]), flipV: boolean(attributes["data-office-flip-v"]),
  };
  if (!PresentationGeometrySchema.safeParse(value).success) invalid("Presentation dimensions must be nonnegative.");
  return value;
}
function unique<Value>(values: readonly Value[], key: (value: Value) => string): Map<string, Value> {
  const result = new Map<string, Value>();
  for (const value of values) {
    const id = key(value);
    if (result.has(id)) invalid("Presentation identities must be unique within their definition type.");
    result.set(id, value);
  }
  return result;
}

/** Validate references against the one HTML/CSS source, without DOM or Office I/O. */
export function validatePresentation(presentation: PresentationMetadata, annotations: readonly Annotation[], html: ParsedHtml, styles: readonly { css: string }[], limits: DocumentReadLimits): void {
  let work = Object.values(presentation).reduce<number>((count, value) => count + (Array.isArray(value) ? value.length : 0), 0);
  checkLimit(work, limits.maxNodes);
  for (const object of presentation.objects) if (object.kind === "table") { work += object.grid.length; checkLimit(work, limits.maxNodes); }
  for (const dataSet of presentation.dataSets) {
    work += dataSet.categories.length + dataSet.series.length;
    for (const series of dataSet.series) work += series.values.length;
    checkLimit(work, limits.maxNodes);
  }
  for (const annotation of annotations) { work += 1 + (annotation.replies?.length ?? 0); checkLimit(work, limits.maxNodes); }
  const classes = new Set<string>();
  for (const style of styles) walk(parse(style.css), (node) => { if (node.type === "ClassSelector") classes.add(ident.decode(node.name)); });
  const themes = unique(presentation.themes, (item) => item.themeId);
  const masters = unique(presentation.masters, (item) => item.masterId);
  const layouts = unique(presentation.layouts, (item) => item.layoutId);
  const slides = unique(presentation.slides, (item) => item.slideId);
  const objects = unique(presentation.objects, (item) => item.nodeId);
  const charts = unique(presentation.charts, (item) => item.chartId);
  const dataSets = unique(presentation.dataSets, (item) => item.dataSetId);
  const notesMasters = unique(presentation.notesMasters, (item) => item.notesMasterId);
  unique(presentation.stories, (item) => item.storyId);
  const roots = new Map<string, string>();
  const node = (id: string) => {
    const value = html.peekNode(id);
    if (!value) invalid("A presentation reference targets a missing source node.");
    return value;
  };
  const presentationRoots = html.nodeIds().filter((id) => node(id).attributes["data-office-kind"] === "presentation");
  if (presentationRoots.length !== 1 || node(presentationRoots[0]!).parentNodeId !== null) invalid("Presentation requires exactly one root.");
  const root = presentationRoots[0]!;
  if (html.directText(root)?.trim()) invalid("Unregistered text cannot appear directly in a presentation root.");
  const addRoot = (id: string, kind: string): void => {
    const value = node(id);
    if (roots.has(id) || !html.isContainer(id) || value.parentNodeId !== root
      || value.attributes["data-office-kind"] !== kind || value.attributes["data-office-story"] !== kind) invalid("Presentation content roots require distinct, matching source containers.");
    roots.set(id, kind);
  };
  for (const theme of themes.values()) if (!classes.has(theme.cssClass)) invalid("A presentation theme refers to an undefined CSS class.");
  for (const master of masters.values()) {
    if (!themes.has(master.themeId)) invalid("A slide master refers to a missing theme.");
    for (const cssClass of [master.titleStyle, master.bodyStyle, master.otherStyle])
      if (cssClass !== undefined && !classes.has(cssClass)) invalid("A master text style refers to an undefined CSS class.");
    addRoot(master.rootNodeId, "master");
  }
  for (const layout of layouts.values()) {
    if (!masters.has(layout.masterId)) invalid("A slide layout refers to a missing master.");
    addRoot(layout.rootNodeId, "layout");
  }
  for (const slide of slides.values()) {
    if (!layouts.has(slide.layoutId)) invalid("A slide refers to a missing layout.");
    addRoot(slide.rootNodeId, "slide");
  }
  for (const master of presentation.notesMasters) {
    if (!themes.has(master.themeId)) invalid("A notes master refers to a missing theme.");
    addRoot(master.rootNodeId, "notes-master");
  }
  const notes = new Set<string>();
  for (const story of presentation.stories) {
    if (story.slideId !== undefined && !slides.has(story.slideId)) invalid("A presentation story refers to a missing slide.");
    if (story.kind === "notes") {
      if (!story.slideId || notes.has(story.slideId)) invalid("Each notes story requires its own existing slide.");
      if (story.notesMasterId !== undefined && !notesMasters.has(story.notesMasterId)) invalid("A notes story refers to a missing notes master.");
      notes.add(story.slideId);
    }
    addRoot(story.rootNodeId, story.kind);
  }
  for (const id of node(root).childNodeIds) if (!roots.has(id)) invalid("Presentation root children must be registered source surfaces.");
  const owner = (id: string): string => {
    for (let value: string | null = id; value !== null; value = node(value).parentNodeId) if (roots.has(value)) return value;
    return invalid("Presentation objects must belong to a registered surface.");
  };
  const bySurface = new Map<string, PresentationObject[]>();
  const placeholders = new Map<string, Map<string, PresentationObject>>();
  const transforms = new Map<string, PresentationGeometry | undefined>();
  const slidesByRoot = new Map(presentation.slides.map((item) => [item.rootNodeId, item]));
  const layoutsByRoot = new Map(presentation.layouts.map((item) => [item.rootNodeId, item]));
  const nativeIds = new Map<string, Set<string>>();
  const placeholderKeys = new Map<string, Set<string>>();
  for (const object of objects.values()) {
    const value = node(object.nodeId), surface = owner(object.nodeId), parent = value.parentNodeId;
    if (!html.isContainer(object.nodeId) || value.attributes["data-office-kind"] !== object.kind
      || parent === null || parent !== surface && objects.get(parent)?.kind !== "group") invalid("A presentation object has an invalid source type or parent.");
    if (["notes", "comment"].includes(roots.get(surface)!)) invalid("Rich-text stories cannot contain positioned slide objects.");
    const transform = readPresentationGeometry(value.attributes);
    transforms.set(object.nodeId, transform);
    if (!transform && !object.placeholder) invalid("Positioned objects require a source transform.");
    if (transform && object.kind !== "connector" && !(object.kind === "geometry" && object.preset === "line") && (!transform.width || !transform.height)) invalid("Two-dimensional objects require positive dimensions.");
    if (object.nativeId !== undefined) {
      const ids = nativeIds.get(surface) ?? new Set<string>();
      if (ids.has(object.nativeId)) invalid("Native object identities must be unique within their surface.");
      ids.add(object.nativeId); nativeIds.set(surface, ids);
    }
    if (object.placeholder) {
      const key = object.placeholder.type + ":" + object.placeholder.index;
      const keys = placeholderKeys.get(surface) ?? new Set<string>();
      if (keys.has(key)) invalid("Placeholder identities must be unique within a surface.");
      keys.add(key); placeholderKeys.set(surface, keys);
      const entries = placeholders.get(surface) ?? new Map<string, PresentationObject>();
      entries.set(key, object); placeholders.set(surface, entries);
    }
    if (object.kind === "group") {
      for (const name of geometryNames) {
        const part = number(value.attributes["data-office-child-" + name]);
        if (["width", "height"].includes(name) && part <= 0) invalid("Group child extents must be positive.");
      }
      if (value.childNodeIds.some((id) => !objects.has(id))) invalid("A group may only contain registered objects.");
      if (html.directText(object.nodeId)?.trim()) invalid("Unregistered text cannot appear directly in a group.");
    }
    if (object.kind === "image") {
      if (value.childNodeIds.length !== 1 || node(value.childNodeIds[0]!).tagName !== "img" || html.directText(object.nodeId)?.trim()) invalid("An image object requires exactly one image source.");
      for (const name of ["left", "top", "right", "bottom"]) {
        const raw = value.attributes["data-office-crop-" + name];
        if (raw !== undefined && (number(raw) < 0 || number(raw) >= 1)) invalid("Image crop fractions must lie in [0, 1).");
      }
      const crop = (name: string): number => Number(value.attributes["data-office-crop-" + name] ?? 0);
      if (crop("left") + crop("right") >= 1 || crop("top") + crop("bottom") >= 1) invalid("An image crop must retain a visible area.");
    }
    if (object.kind === "table") {
      const tables = value.childNodeIds.filter((id) => node(id).tagName === "table");
      if (tables.length !== 1 || value.childNodeIds.length !== 1 || html.directText(object.nodeId)?.trim()) invalid("A table object requires exactly one table source.");
      const rows = node(tables[0]!).childNodeIds;
      const occupied = new Array<number>(object.grid.length).fill(0);
      for (const [rowIndex, rowId] of rows.entries()) {
        if (node(rowId).tagName !== "tr") invalid("Presentation table children must be rows.");
        let column = 0;
        for (const cellId of node(rowId).childNodeIds) {
          const cell = node(cellId);
          if (!["td", "th"].includes(cell.tagName)) invalid("Presentation table rows require cells.");
          while (column < occupied.length && occupied[column]! > rowIndex) { checkLimit(++work, limits.maxNodes); column++; }
          const span = Number(cell.attributes.colspan ?? 1);
          const rowSpan = Number(cell.attributes.rowspan ?? 1) || rows.length - rowIndex;
          if (column + span > occupied.length || rowIndex + rowSpan > rows.length) invalid("A presentation table span exceeds its grid.");
          for (let end = column + span; column < end; column++) {
            checkLimit(++work, limits.maxNodes);
            if (occupied[column]! > rowIndex) invalid("Presentation table spans overlap.");
            occupied[column] = rowIndex + rowSpan;
          }
        }
        for (const end of occupied) { checkLimit(++work, limits.maxNodes); if (end <= rowIndex) invalid("Presentation table cells must cover their row grid."); }
      }
      if (!rows.length) invalid("A presentation table requires rows.");
    }
    if (object.kind === "chart" && charts.get(object.chartId)?.nodeId !== object.nodeId) invalid("A chart object requires a matching chart definition.");
    if (object.kind === "connector") for (const target of [object.startNodeId, object.endNodeId])
      if (target !== undefined && (!objects.has(target) || target === object.nodeId || owner(target) !== surface)) invalid("Connector endpoints must be other objects on the same surface.");
    const list = bySurface.get(surface) ?? []; list.push(object); bySurface.set(surface, list);
  }
  // The inherited placeholder transform must resolve through layout/master, not
  // through a nearby object or a duplicate copy of a rendered instance.
  for (const [surface, list] of bySurface) for (const object of list) {
    if (transforms.get(object.nodeId)) continue;
    const slide = slidesByRoot.get(surface);
    const layout = slide ? layouts.get(slide.layoutId) : layoutsByRoot.get(surface);
    const parents = layout ? [layout.rootNodeId, masters.get(layout.masterId)!.rootNodeId].filter((id) => id !== surface) : [];
    const key = object.placeholder!.type + ":" + object.placeholder!.index;
    if (!parents.some((id) => {
      const inherited = placeholders.get(id)?.get(key);
      return inherited !== undefined && transforms.get(inherited.nodeId) !== undefined;
    })) invalid("A placeholder transform cannot be resolved from its layout or master.");
  }
  for (const id of html.nodeIds()) {
    const value = node(id), kind = value.attributes["data-office-kind"];
    if (kind !== undefined && id !== root && !roots.has(id) && !objects.has(id)) invalid("An unregistered presentation source role is invalid.");
    if (Object.keys(value.attributes).some((key) => key.startsWith("data-word-"))) invalid("Word semantics cannot be embedded in a presentation source.");
    if (value.attributes["data-office-story"] !== undefined && !roots.has(id)) invalid("Presentation stories must be registered content roots.");
    if (roots.has(id) && !["notes", "comment"].includes(roots.get(id)!) && (value.childNodeIds.some((child) => !objects.has(child)) || html.directText(id)?.trim())) invalid("A slide surface may only contain registered objects.");
  }
  for (const dataSet of dataSets.values()) {
    unique(dataSet.series, (item) => item.seriesId);
    if (dataSet.series.some((series) => series.values.length !== dataSet.categories.length)) invalid("Chart categories and series lengths must match.");
  }
  const negativeDataSets = new Set(presentation.dataSets.filter((dataSet) => dataSet.series.some((series) => series.values.some((value) => value !== null && value < 0))).map((item) => item.dataSetId));
  for (const chart of charts.values()) {
    const object = objects.get(chart.nodeId), dataSet = dataSets.get(chart.dataSetId);
    if (!object || object.kind !== "chart" || object.chartId !== chart.chartId || !dataSet) invalid("Chart references must resolve to their object and data set.");
    if (chart.titleNodeId !== undefined) {
      let parent: string | null = chart.titleNodeId;
      while (parent !== null && parent !== chart.nodeId) parent = node(parent).parentNodeId;
      if (parent !== chart.nodeId || chart.titleNodeId === chart.nodeId) invalid("A chart title must be inside its chart source.");
    }
    if (chart.minimum !== undefined && chart.maximum !== undefined && chart.minimum >= chart.maximum) invalid("Chart axis bounds are invalid.");
    if (chart.kind === "pie" && (dataSet.series.length !== 1 || negativeDataSets.has(dataSet.dataSetId))) invalid("A pie chart requires one nonnegative series.");
  }
  for (const link of unique(presentation.hyperlinks, (item) => item.nodeId).values()) {
    node(link.nodeId);
    if ((link.target === undefined) === (link.slideId === undefined) || link.slideId !== undefined && !slides.has(link.slideId)
      || link.target !== undefined && (!/^(https?:\/\/|mailto:)/i.test(link.target) || /[\u0000-\u0020\u007f]/.test(link.target))) invalid("Presentation hyperlinks require one supported target.");
  }
  const annotationMap = unique(annotations, (item) => item.annotationId);
  for (const comment of unique(presentation.comments, (item) => item.annotationId).values())
    if (!annotationMap.has(comment.annotationId)) invalid("Native comment metadata refers to a missing annotation.");
  const replyIds = new Set(annotationMap.keys());
  const bodies = new Set<string>();
  for (const annotation of annotations) for (const body of [annotation, ...(annotation.replies ?? [])]) {
    if ("replyId" in body) {
      if (replyIds.has(body.replyId)) invalid("Replies require distinct document-wide identities.");
      replyIds.add(body.replyId);
    }
    if (body.bodyNodeId !== undefined) {
      if (roots.get(body.bodyNodeId) !== "comment" || bodies.has(body.bodyNodeId)) invalid("Rich comments require distinct comment source roots.");
      bodies.add(body.bodyNodeId);
    }
  }
}
