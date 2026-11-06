import { DocumentError } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookAnnotationWindow } from "./annotation-projection.js";
import type { WorkbookViewportLayout } from "./viewport.js";

export interface WorkbookAnnotationView {
  readonly element: HTMLElement;
  /** Current generated nodes, including selection highlights. */
  readonly nodes: number;
  readonly groups: readonly { readonly annotationIds: readonly string[]; readonly button: HTMLButtonElement }[];
  select(annotationId: string | undefined): void;
  dispose(): void;
}

/** Read-only SDK controls over clipped source ranges. One accessible entry per
 * annotation; coincident entries share a button. Markers sit at the corner of
 * the first visible fragment, not at a newly invented source anchor. The owner
 * opens the chosen thread/list and manages focus and panel placement. */
export function createWorkbookAnnotationView(document: Document, annotations: WorkbookAnnotationWindow, layout: WorkbookViewportLayout,
  rightToLeft: boolean, open: (annotationIds: readonly string[]) => void, maxNodes = 50_000): WorkbookAnnotationView {
  if (!document.defaultView || !Number.isSafeInteger(maxNodes) || maxNodes < 2 || ![layout.widthPt, layout.heightPt].every(value => Number.isFinite(value) && value > 0))
    throw new DocumentError("InvalidOperation", "Annotation controls require a live document, finite viewport and node budget.");
  const element = markViewNode(document.createElement("div")), highlights = markViewNode(document.createElement("div"));
  element.dataset.officeView = "workbook-annotations"; element.style.cssText = `position:absolute;left:0;top:0;width:${layout.widthPt}pt;height:${layout.heightPt}pt;overflow:hidden;pointer-events:none;direction:ltr`;
  highlights.dataset.officeView = "workbook-annotation-highlights"; highlights.setAttribute("aria-hidden", "true"); element.append(highlights);
  const grouped = new Map<string, { x: number; y: number; width: number; height: number; annotationIds: string[] }>();
  const controls = new document.defaultView.AbortController(); let disposed = false, nodes = 2;
  const account = (count: number): void => { nodes += count; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Annotation controls exceed their cumulative node budget."); };
  const byId = new Map(annotations.visible.map(item => [item.annotation.annotationId, item]));
  try {
    for (const item of annotations.visible) {
      const first = item.regions[0]; if (!first || item.annotation.status === "needs_reanchor") continue;
      const { clip } = first, width = Math.min(12, clip.width), height = Math.min(12, clip.height);
      const x = rightToLeft ? layout.widthPt - clip.x - clip.width : clip.x + clip.width - width, y = clip.y;
      if (![x, y, width, height].every(Number.isFinite) || Math.min(width, height) <= 0) throw new DocumentError("InvalidOperation", "Annotation controls require finite visible fragment geometry.");
      const key = `${x}:${y}:${width}:${height}`;
      const group = grouped.get(key);
      if (group) group.annotationIds.push(item.annotation.annotationId);
      else { account(2); grouped.set(key, { x, y, width, height, annotationIds: [item.annotation.annotationId] }); }
    }
    const groups = [...grouped.values()].map(group => {
      const button = document.createElement("button"), ids = Object.freeze(group.annotationIds);
      button.type = "button"; button.textContent = ids.length > 9 ? "9+" : String(ids.length); button.setAttribute("aria-label", "查看 " + ids.length + " 条批注");
      button.style.cssText = `position:absolute;box-sizing:border-box;left:${group.x}pt;top:${group.y}pt;width:${group.width}pt;height:${group.height}pt;margin:0;padding:0;border:1px solid #3464bd;border-radius:3px;background:#edf3ff;color:#173d7a;font:8pt/1 sans-serif;cursor:pointer;pointer-events:auto;overflow:hidden;white-space:nowrap`;
      button.addEventListener("click", () => open(ids), { signal: controls.signal }); element.append(button);
      return Object.freeze({ annotationIds: ids, button });
    });
    const baseNodes = nodes;
    const select = (annotationId: string | undefined): void => {
      if (disposed) throw new DocumentError("InvalidOperation", "These workbook annotation controls have been disposed.");
      const item = annotationId === undefined ? undefined : byId.get(annotationId), fragments = item?.regions ?? [];
      if (baseNodes + fragments.length > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Annotation highlights exceed their cumulative node budget.");
      const next = document.createDocumentFragment();
      for (const { rect, clip } of fragments) {
        const node = markViewNode(document.createElement("div"));
        const x = rightToLeft ? layout.widthPt - clip.x - clip.width : clip.x;
        node.style.cssText = `position:absolute;box-sizing:border-box;left:${x}pt;top:${clip.y}pt;width:${clip.width}pt;height:${clip.height}pt;background:rgba(52,100,189,.12);pointer-events:none`;
        // A viewport/frozen cut is not a source range edge. Paint only edges
        // present in this fragment, keeping giant ranges in small view boxes.
        const edge = (name: string, visible: boolean): void => { if (visible) node.style.setProperty("border-" + name, "1px solid #3464bd"); };
        edge("top", Math.abs(rect.y - clip.y) < 1e-6); edge("bottom", Math.abs(rect.y + rect.height - clip.y - clip.height) < 1e-6);
        edge(rightToLeft ? "right" : "left", Math.abs(rect.x - clip.x) < 1e-6); edge(rightToLeft ? "left" : "right", Math.abs(rect.x + rect.width - clip.x - clip.width) < 1e-6);
        next.append(node);
      }
      highlights.replaceChildren(next); nodes = baseNodes + fragments.length;
      for (const group of groups) group.button.setAttribute("aria-pressed", String(annotationId !== undefined && group.annotationIds.includes(annotationId)));
    };
    select(undefined);
    return Object.freeze({ element, get nodes() { return nodes; }, groups: Object.freeze(groups), select,
      dispose() { if (disposed) return; disposed = true; controls.abort(); element.remove(); element.replaceChildren(); nodes = 0; } });
  } catch (error) { controls.abort(); element.remove(); element.replaceChildren(); throw error; }
}
