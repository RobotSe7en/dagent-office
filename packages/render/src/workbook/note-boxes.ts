import { DocumentError } from "@doffi/core";
import type { WorkbookSnapshot } from "@doffi/core";
import type { FontResolution } from "../contracts.js";
import { resolvePreviewFontsInScope } from "../fonts.js";
import { addArrows, property } from "../presentation/appearance.js";
import { svgElement } from "../presentation/svg.js";
import { markViewNode } from "../view-nodes.js";
import { createWorkbookCommentText } from "./comment-text.js";
import type { WorkbookCommentText } from "./comment-text.js";
import type { WorkbookContentSources } from "./content-source.js";
import type { AxisDimensions } from "./dimensions.js";
import { projectWorkbookDrawing, workbookDrawingRect } from "./drawing.js";
import type { WorkbookDrawingProjection } from "./drawing.js";
import { createWorkbookDrawingFragment } from "./drawing-view.js";
import type { WorkbookSurface } from "./surface.js";
import type { WorkbookWindowFrame } from "./window-loader.js";
import { checkWindowCancellation } from "./window-source.js";
import { distributeWorkbookTextLines } from "./text-lines.js";

export interface WorkbookNoteProjection {
  readonly annotationId: string; readonly projection: WorkbookDrawingProjection; readonly text: WorkbookCommentText;
}
export interface WorkbookNoteBoxLayout {
  readonly element: HTMLElement; readonly nodes: number; readonly codeUnits: number;
  readonly fonts: readonly FontResolution[]; readonly notes: readonly WorkbookNoteProjection[];
}

/** Native, visible classic notes use their selected source box. The regular
 * comment panel remains responsible for hidden notes and modern threads.
 * Every pane owns fresh text and exact UTF-16 projections; no VML reaches DOM. */
export async function layoutWorkbookNoteBoxes(surface: Pick<WorkbookSurface, "source" | "view">, snapshot: WorkbookSnapshot,
  frame: WorkbookWindowFrame, dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions },
  sources: WorkbookContentSources, signal: AbortSignal,
  limits: { readonly maxNodes: number; readonly maxSourceNodes: number; readonly maxCodeUnits: number; readonly maxExtentPt: number; readonly maxObjects: number }): Promise<WorkbookNoteBoxLayout | undefined> {
  if (!snapshot.metadata.workbook.comments?.some(comment => comment.format === "classic" && comment.box?.visible)) return;
  const annotations = new Map(snapshot.annotations.map(annotation => [annotation.annotationId, annotation]));
  const selected = (snapshot.metadata.workbook.comments ?? []).flatMap(comment => {
    const annotation = annotations.get(comment.annotationId);
    return comment.format === "classic" && comment.box?.visible && annotation?.status !== "needs_reanchor"
      && annotation?.anchor.kind === "cells" && annotation.anchor.sheetId === frame.sheetId && annotation.bodyNodeId
      ? [{ comment, annotation, box: comment.box, bodyNodeId: annotation.bodyNodeId }] : [];
  });
  if (!selected.length) return;
  const document = surface.source.ownerDocument, view = document.defaultView!;
  const projection = sources.project(surface.source, selected.map(item => item.bodyNodeId), limits.maxSourceNodes);
  const element = markViewNode(document.createElement("div")), fonts: FontResolution[] = [], notes: WorkbookNoteProjection[] = [];
  element.dataset.officeView = "workbook-note-boxes"; element.style.cssText = "position:absolute;inset:0;pointer-events:none";
  let nodes = 1, codeUnits = 0, visible = 0;
  const account = (amount: number): void => { nodes += amount; if (nodes > limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Native note boxes exceed their cumulative node budget."); };
  const make: typeof svgElement = (document, tag, attributes) => { account(1); return svgElement(document, tag, attributes); };
  const count = (root: Node): number => { const walker = document.createTreeWalker(root, 0xffff_ffff); let count = 1; while (walker.nextNode()) count++; return count; };
  try {
    for (const [index, item] of selected.entries()) {
      if (index % 16 === 0) await new Promise<void>(resolve => setTimeout(resolve, 0));
      checkWindowCancellation(signal);
      const rect = workbookDrawingRect(item.box.anchor, dimensions);
      if (Math.max(rect.width, rect.height) > limits.maxExtentPt) throw new DocumentError("ResourceLimitExceeded", "A native note exceeds its measurable extent budget.");
      const source = projection.get(item.bodyNodeId)!, probe = document.createElement("span"); probe.className = item.box.cssClass ?? "";
      probe.style.cssText = "position:absolute;visibility:hidden"; account(1); surface.source.append(probe);
      try {
        const style = view.getComputedStyle(probe);
        const number = (name: string, fallback: number, length = false): number => {
          const raw = property(style, name), value = raw ? Number(length ? raw.replace(/pt$/, "") : raw) : fallback;
          if (raw && length && !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?pt$/i.test(raw) || !Number.isFinite(value)) throw new DocumentError("InvalidDocument", "A native note has an invalid numeric appearance.");
          return value;
        };
        const opacity = (name: string): number => { const value = number(name, 1); if (value < 0 || value > 1) throw new DocumentError("InvalidDocument", "A native note opacity is outside its range."); return value; };
        const stroke = property(style, "--office-note-stroke", "black"), weight = number("--office-note-stroke-width", .75, true);
        const shadow = property(style, "--office-note-shadow", "none"), sx = number("--office-note-shadow-x", 2, true), sy = number("--office-note-shadow-y", 2, true);
        const inset = [style.paddingLeft, style.paddingTop, style.paddingRight, style.paddingBottom].map(raw => parseFloat(raw) * .75);
        const extent = Math.max(weight * 8, shadow === "none" ? 0 : Math.abs(sx) + Math.abs(sy), ...inset);
        if (weight < 0 || inset.some(value => !Number.isFinite(value) || value < 0) || !Number.isFinite(extent) || extent > limits.maxExtentPt) throw new DocumentError("ResourceLimitExceeded", "A native note appearance exceeds its extent budget.");
        const parts = projectWorkbookDrawing(rect, 0, frame.layout, extent);
        if (parts.length && ++visible > limits.maxObjects) throw new DocumentError("ResourceLimitExceeded", "Visible native notes exceed their object budget.");
        const dash = property(style, "--office-note-stroke-dashstyle", "solid"), patterns: Record<string, readonly number[]> = { shortdash: [3, 1], shortdot: [1, 1], shortdashdot: [3, 1, 1, 1], shortdashdotdot: [3, 1, 1, 1, 1, 1], dot: [1, 3], dash: [4, 3], longdash: [8, 3], dashdot: [4, 3, 1, 3], longdashdot: [8, 3, 1, 3], longdashdotdot: [8, 3, 1, 3, 1, 3] };
        for (const [partIndex, part] of parts.entries()) {
          checkWindowCancellation(signal); account(1);
          const content = markViewNode(document.createElement("div")); content.style.cssText = `position:absolute;width:${rect.width}pt;height:${rect.height}pt`;
          content.dataset.officeView = "workbook-note-box"; content.dataset.annotationId = item.annotation.annotationId;
          content.dataset.officePrintable = String(item.box.printable); content.dataset.officeBoxSource = item.comment.boxSource ?? "vml";
          content.style.zIndex = property(style, "--office-note-z-index", String(index));
          const graphic = make(document, "svg", { width: "100%", height: "100%", viewBox: `0 0 ${Math.max(.001, rect.width)} ${Math.max(.001, rect.height)}`, "aria-hidden": "true" });
          graphic.style.cssText = "position:absolute;inset:0;overflow:visible;pointer-events:none";
          const fill = property(style, "--office-note-fill", "#ffffe1");
          if (shadow !== "none") graphic.append(make(document, "rect", { x: sx, y: sy, width: rect.width, height: rect.height, fill: fill === "none" && property(style, "--office-note-shadow-obscured", "false") !== "true" ? "none" : shadow, stroke: stroke === "none" ? "none" : shadow, "stroke-width": weight, opacity: opacity("--office-note-shadow-opacity") }));
          const arrows = ["start", "end"].some(side => property(style, `--office-note-${side}-arrow`, "none") !== "none");
          const framePaint = { fill, "fill-opacity": opacity("--office-note-fill-opacity"), stroke, "stroke-width": weight,
            "stroke-opacity": opacity("--office-note-stroke-opacity"), "stroke-linejoin": property(style, "--office-note-stroke-joinstyle", "miter"),
            "stroke-linecap": property(style, "--office-note-stroke-endcap", "flat") === "flat" ? "butt" : property(style, "--office-note-stroke-endcap"),
            "stroke-dasharray": patterns[dash]?.map(value => value * weight).join(" ") ?? "none" };
          if (arrows && stroke !== "none") {
            const path = make(document, "path", { d: `M0 0H${rect.width}V${rect.height}H0Z`, ...framePaint }); graphic.append(path);
            probe.style.setProperty("--office-line-color", stroke); probe.style.setProperty("--office-line-alpha", String(opacity("--office-note-stroke-opacity")));
            for (const side of ["start", "end"]) {
              const names: Record<string, string> = { none: "none", block: "triangle", classic: "stealth", oval: "oval", diamond: "diamond", open: "arrow" };
              const arrow = names[property(style, `--office-note-${side}-arrow`, "none")];
              if (!arrow) throw new DocumentError("InvalidDocument", "A native note has an unknown arrow shape.");
              probe.style.setProperty(`--office-line-${side}-arrow`, arrow);
              probe.style.setProperty(`--office-line-${side}-w`, ({ narrow: "sm", medium: "med", wide: "lg" } as Record<string, string>)[property(style, `--office-note-${side}-arrowwidth`, "medium")] ?? "med");
              probe.style.setProperty(`--office-line-${side}-len`, ({ short: "sm", medium: "med", long: "lg" } as Record<string, string>)[property(style, `--office-note-${side}-arrowlength`, "medium")] ?? "med");
            }
            addArrows(document, graphic, path, style, `workbook-note-${index}-${partIndex}`, account);
          } else graphic.append(make(document, "rect", { x: 0, y: 0, width: rect.width, height: rect.height, ...framePaint }));
          content.append(graphic);
          let text = createWorkbookCommentText(source, Math.max(1, limits.maxNodes - nodes), Math.max(1, limits.maxCodeUnits - codeUnits)); account(text.nodes);
          codeUnits += text.sourceCodeUnits + (text.element.textContent?.length ?? 0);
          if (codeUnits > limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Native notes exceed their cumulative text budget.");
          account(1); const area = document.createElement("div"), vertical = property(style, "--office-note-vertical", "top"), horizontal = property(style, "--office-note-horizontal", "left");
          area.style.cssText = `position:absolute;box-sizing:border-box;left:${inset[0]}pt;top:${inset[1]}pt;width:${Math.max(0, rect.width - inset[0]! - inset[2]!)}pt;height:${Math.max(0, rect.height - inset[1]! - inset[3]!)}pt;overflow:hidden;display:flex;flex-direction:column;pointer-events:auto`;
          text.element.style.flex = "0 0 auto"; text.element.style.width = "100%"; text.element.style.whiteSpace = "pre-wrap";
          if (vertical === "center" || vertical === "bottom") text.element.style.marginTop = "auto";
          if (vertical === "center") text.element.style.marginBottom = "auto";
          for (const p of text.element.children) {
            (p as HTMLElement).style.textAlign = horizontal === "distributed" ? "justify" : horizontal;
            (p as HTMLElement).style.textAlignLast = horizontal === "distributed" || item.box.justifyLastLine ? "justify" : "auto";
          }
          area.append(text.element); content.append(area);
          // Resolve on this candidate, then detach before publication.
          const measurement = document.createElement("div"); account(1); measurement.style.cssText = "position:absolute;visibility:hidden"; measurement.append(content); surface.view.append(measurement);
          try {
            const resolved = resolvePreviewFontsInScope(document, text.element, "[data-office-text-style],[data-office-script]"); if (!partIndex) fonts.push(...resolved);
            if (vertical === "justify" || vertical === "distributed") {
              const distributed = distributeWorkbookTextLines(text, Math.max(0, rect.height - inset[1]! - inset[3]!) * 4 / 3, Math.max(1, limits.maxNodes - nodes + text.nodes));
              account(distributed.text.nodes - text.nodes); text = Object.freeze({ ...text, ...distributed.text });
              if (vertical === "distributed") text.element.style.justifyContent = "space-around";
            }
          }
          finally { content.remove(); measurement.remove(); nodes--; }
          const fragment = createWorkbookDrawingFragment(document, item.bodyNodeId, part, frame.layout.widthPt, { rotationDegrees: 0, flipH: false, flipV: false, hidden: false }, snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === frame.sheetId)!.rightToLeft, content);
          account(count(fragment) - count(content)); fragment.dataset.officeView = "workbook-note"; fragment.style.zIndex = content.style.zIndex;
          if (partIndex) fragment.setAttribute("aria-hidden", "true"); else { fragment.setAttribute("role", "note"); if (item.box.altText) fragment.setAttribute("aria-label", item.box.altText); }
          element.append(fragment); notes.push(Object.freeze({ annotationId: item.annotation.annotationId, projection: part, text }));
        }
      } finally { probe.remove(); nodes--; }
    }
    return Object.freeze({ element, nodes, codeUnits, fonts: Object.freeze(fonts), notes: Object.freeze(notes) });
  } catch (error) { element.replaceChildren(); throw error; }
  finally { projection.dispose(); }
}
