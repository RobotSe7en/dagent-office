import type { Annotation } from "../contracts/annotations.js";
import type { DocumentContentOperation } from "../contracts/operations.js";
import type { DocumentMetadata } from "../contracts/document.js";
import type { DocumentState } from "../document/state.js";
import type { ParsedHtml } from "../document/html.js";
import type { DocumentEdit } from "../document/edit.js";
import { assetDigest } from "../document/assets.js";
import { encode } from "../document/package.js";
import { invalid } from "../document/validation.js";

export function guardPresentationContent(base: DocumentState, op: DocumentContentOperation): void {
  const p = base.snapshot.metadata.presentation;
  if (!p) return;
  const object = p.objects.find(o => o.nodeId === op.nodeId), node = base.html.peekNode(op.nodeId)!;
  if (op.op === "setStyle" && node.attributes["data-office-kind"] !== undefined && /^(?:(?:min|max)-)?(?:width|height)$/i.test(op.property)) invalid("Presentation object dimensions require a typed geometry operation.");
  if (op.op === "removeNode" && ["presentation", "slide", "master", "layout", "notes-master", "notes", "comment"].includes(node.attributes["data-office-kind"] ?? "")) invalid("Presentation surfaces require a format management operation.");
  if (["replaceText", "replaceFragment", "insertFragment", "reconcileFragment"].includes(op.op) && node.attributes["data-office-kind"] !== undefined
    && (!object || !["text", "geometry", "table", "chart"].includes(object.kind)) && !["notes", "comment"].includes(node.attributes["data-office-kind"]!)) invalid("Content operations require a presentation text content domain.");
}

/** Only complete deleted objects are reconciled; damaged required references still fail. */
export function updatePresentationMetadata(metadata: DocumentMetadata, before: DocumentState, after: ParsedHtml, annotations: Annotation[]): void {
  const p = metadata.presentation;
  if (!p) return;
  p.objects = p.objects.filter(item => after.hasNode(item.nodeId));
  const objects = new Set(p.objects.map(o => o.nodeId));
  for (const object of p.objects) if (object.kind === "connector") {
    if (object.startNodeId && !objects.has(object.startNodeId)) { delete object.startNodeId; delete object.startSite; }
    if (object.endNodeId && !objects.has(object.endNodeId)) { delete object.endNodeId; delete object.endSite; }
  }
  p.charts = p.charts.filter(item => after.hasNode(item.nodeId));
  const usedData = new Set(p.charts.map(c => c.dataSetId)), previousData = new Set(before.snapshot.metadata.presentation?.charts.map(c => c.dataSetId));
  p.dataSets = p.dataSets.filter(item => usedData.has(item.dataSetId) || !previousData.has(item.dataSetId));
  p.hyperlinks = p.hyperlinks.filter(item => after.hasNode(item.nodeId));
  p.stories = p.stories.filter(item => item.kind !== "comment" || after.hasNode(item.rootNodeId));
  const byId = new Map(annotations.map(a => [a.annotationId, a]));
  p.comments = p.comments.filter(item => byId.has(item.annotationId));
  const existing = new Set(p.comments.map(c => c.annotationId));
  for (const annotation of annotations) if (!existing.has(annotation.annotationId)) p.comments.push({ annotationId: annotation.annotationId, format: "modern" });
}

export function removeDeletedCommentBodies(before: DocumentState, annotations: readonly Annotation[], edit: DocumentEdit): void {
  if (!before.snapshot.metadata.presentation) return;
  const kept = new Set(annotations.flatMap(a => [a.bodyNodeId, ...(a.replies ?? []).map(r => r.bodyNodeId)]).filter(id => id !== undefined));
  for (const body of before.snapshot.annotations.flatMap(a => [a, ...(a.replies ?? [])])) {
    if (body.bodyNodeId && !kept.has(body.bodyNodeId)) edit.apply({ op: "removeNode", nodeId: body.bodyNodeId });
  }
}

export function updatePresentationRecovery(metadata: DocumentMetadata, before: DocumentState, annotations: readonly Annotation[]): void {
  const p = metadata.presentation;
  if (!p) return;
  const originals = new Map(before.snapshot.annotations.map(a => [a.annotationId, a]));
  for (const comment of p.comments) {
    const annotation = annotations.find(a => a.annotationId === comment.annotationId)!;
    const original = originals.get(comment.annotationId);
    if (annotation.status === "needs_reanchor") {
      if (original && original.status !== "needs_reanchor") comment.recoveryStatus = original.status;
    } else delete comment.recoveryStatus;
  }
}

/** A derived fingerprint of each slide and its transitive source/data dependencies. */
export function presentationSlideFingerprints(state: Pick<DocumentState, "snapshot" | "html">): Map<string, string> {
  const { snapshot, html } = state, p = snapshot.metadata.presentation, result = new Map<string, string>();
  if (!p) return result;
  const digest = (value: unknown): string => assetDigest(encode(JSON.stringify(value)));
  const hashes = new Map<string, string>();
  const source = (id: string): string => {
    let hash = hashes.get(id); if (hash) return hash;
    const node = html.peekNode(id);
    if (!node) return "missing";
    hash = digest([html.signature(id), node.childNodeIds.map(source)]); hashes.set(id, hash); return hash;
  };
  const roots = new Set([...p.slides, ...p.masters, ...p.layouts, ...p.notesMasters, ...p.stories].map(s => s.rootNodeId));
  const bySurface = new Map<string, unknown[]>(), owners = new Map<string, string | undefined>();
  const owner = (id: string): string | undefined => {
    if (owners.has(id)) return owners.get(id);
    const node = html.peekNode(id), parent = node?.parentNodeId;
    const value = roots.has(id) ? id : parent ? owner(parent) : undefined; owners.set(id, value); return value;
  };
  const charts = new Map(p.charts.map(c => [c.chartId, c])), data = new Map(p.dataSets.map(d => [d.dataSetId, d]));
  for (const object of [...p.objects].sort((a,b) => a.nodeId.localeCompare(b.nodeId))) {
    const surface = owner(object.nodeId); if (!surface) continue;
    const items = bySurface.get(surface) ?? [], chart = object.kind === "chart" ? charts.get(object.chartId) : undefined;
    items.push([object, chart, chart ? data.get(chart.dataSetId) : undefined]); bySurface.set(surface, items);
  }
  const surface = (id: string): string => digest([source(id), bySurface.get(id)]);
  const surfaces = new Map([...roots].map(id => [id, surface(id)]));
  const shared = digest([p.size, snapshot.styles, html.peekNode(html.nodeIds().find(id => html.peekNode(id)?.attributes["data-office-kind"] === "presentation")!)?.attributes]);
  const layouts = new Map(p.layouts.map(l => [l.layoutId, l])), masters = new Map(p.masters.map(m => [m.masterId, m])), themes = new Map(p.themes.map(t => [t.themeId, t]));
  for (const slide of p.slides) {
    const layout = layouts.get(slide.layoutId), master = layout ? masters.get(layout.masterId) : undefined;
    result.set(slide.slideId, digest([shared, slide, surfaces.get(slide.rootNodeId), layout, layout ? surfaces.get(layout.rootNodeId) : null, master, master ? surfaces.get(master.rootNodeId) : null, master ? themes.get(master.themeId) : null]));
  }
  return result;
}

export function affectedPresentationSlides(before: Pick<DocumentState, "snapshot" | "html">, after: Pick<DocumentState, "snapshot" | "html">): string[] {
  const a = presentationSlideFingerprints(before), b = presentationSlideFingerprints(after);
  return [...new Set([...a.keys(), ...b.keys()])].filter(id => a.get(id) !== b.get(id));
}
