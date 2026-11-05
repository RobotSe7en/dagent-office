import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookChart, WorkbookDisplayFormat, WorkbookNumberFormatter } from "@dagent-office/core";
import type { BasicChartLayout, BasicChartRect } from "../chart.js";
import type { FontResolution } from "../contracts.js";
import { cssStringValue, property } from "../presentation/appearance.js";
import { createWorkbookChartView } from "./chart-view.js";
import type { WorkbookChartBinding, WorkbookChartDataDiagnostic } from "./chart-values.js";
import type { WorkbookChartViewDiagnostic } from "./chart-view.js";
import { layoutWorkbookChartTitleContents } from "./drawing-text-layout.js";
import type { WorkbookChartTitleContentInput, WorkbookDrawingTextLayout } from "./drawing-text-layout.js";
import type { WorkbookChartTextRole } from "./drawing-text.js";

export interface WorkbookChartLayoutLimits { readonly maxNodes?: number; readonly maxPoints?: number; readonly maxCodeUnits?: number }
export interface WorkbookChartLayout {
  readonly element: HTMLElement; readonly nodes: number; readonly layout: BasicChartLayout;
  readonly texts: readonly WorkbookDrawingTextLayout[]; readonly fonts: readonly FontResolution[];
  readonly usedFallbackLineMetrics: boolean;
  readonly diagnostics: readonly (WorkbookChartDataDiagnostic | WorkbookChartViewDiagnostic)[];
  readonly overflowingTitles: readonly WorkbookChartTextRole[];
}

/** Compose current graphics and both kinds of titles before publishing a
 * candidate. Titles retain their measured local boxes; axis-title strips and
 * legend bands occupy separate space. Object transforms, frame effects and
 * worksheet pane clipping remain the drawing compositor's responsibility. */
export function layoutWorkbookChart(scope: HTMLElement, source: HTMLElement, chart: DeepReadonly<WorkbookChart>,
  binding: Extract<WorkbookChartBinding, { kind: "available" }>, widthPt: number, heightPt: number, id: string,
  formatter: WorkbookNumberFormatter, limits: WorkbookChartLayoutLimits = {}): WorkbookChartLayout {
  const document = scope.ownerDocument, view = document.defaultView, maxNodes = limits.maxNodes ?? 50_000, maxUnits = limits.maxCodeUnits ?? 8 * 1024 * 1024;
  if (!scope.isConnected || !source.isConnected || source.ownerDocument !== document || !view || binding.chartId !== chart.chartId
    || source.dataset.nodeId !== chart.nodeId || source.dataset.officeKind !== "chart"
    || ![maxNodes, maxUnits, limits.maxPoints ?? 50_000].every(value => Number.isSafeInteger(value) && value > 0)
    || ![widthPt, heightPt].every(value => Number.isFinite(value) && value > 0 && value <= 1_000_000))
    throw new DocumentError("InvalidOperation", "Chart composition requires its own current binding, connected surface and finite budgets.");
  const position = (classes: string | undefined, fallback: string): string => {
    if (maxNodes < 2) throw new DocumentError("ResourceLimitExceeded", "Chart composition exceeds its temporary node budget.");
    const probe = document.createElement("span"); probe.className = classes ?? ""; probe.style.cssText = "position:absolute;visibility:hidden"; source.append(probe);
    try { return property(view.getComputedStyle(probe), "--office-axis-position", fallback); } finally { probe.remove(); }
  };
  const categoryPosition = position(chart.categoryAxisCssClass, "b"), valuePosition = position(chart.valueAxisCssClass, "l");
  if (!["b", "t"].includes(categoryPosition) || !["l", "r"].includes(valuePosition)) throw new DocumentError("InvalidDocument", "Basic chart axis positions are invalid.");
  const categoryTop = categoryPosition === "t", valueRight = valuePosition === "r", pie = chart.kind === "pie";
  const hasTitle = chart.titleVisible !== false && !!(chart.titleNodeId || chart.titleReference);
  const categoryTitle = !pie && chart.showCategoryAxis !== false && !!chart.categoryAxisTitle, valueTitle = !pie && chart.showValueAxis !== false && !!chart.valueAxisTitle;
  const titleHeight = hasTitle ? Math.min(54, heightPt * .2) : 12, axisBand = 24;
  const legendBand = chart.legend === "none" ? 0 : chart.legend === "top" || chart.legend === "bottom" ? 28 : Math.min(widthPt * .25, 100);
  const left = (pie || valueRight ? 12 : 28 + (valueTitle ? axisBand : 0)) + (chart.legend === "left" ? legendBand : 0);
  const right = widthPt - (pie ? 12 : valueRight ? 28 + (valueTitle ? axisBand : 0) : 16) - (chart.legend === "right" ? legendBand : 0);
  const top = titleHeight + (pie ? 15 : 18) + (chart.legend === "top" ? legendBand : 0) + (categoryTop ? 16 + (categoryTitle ? axisBand : 0) : 0);
  const bottom = heightPt - (pie ? 8 : categoryTop ? 12 : 26 + (categoryTitle ? axisBand : 0)) - (chart.legend === "bottom" ? legendBand : 0);
  if (right <= left || bottom <= top || widthPt <= 12) throw new DocumentError("ResourceLimitExceeded", "This chart canvas has no space for its plot, titles and legend.");
  const plot = Object.freeze({ x: left, y: top, width: right - left, height: bottom - top });
  const title = hasTitle ? Object.freeze({ x: 6, y: 6, width: widthPt - 12, height: titleHeight }) : undefined;
  const legend: BasicChartRect | undefined = chart.legend === "none" ? undefined
    : chart.legend === "left" || chart.legend === "right" ? { x: chart.legend === "left" ? 0 : widthPt - legendBand, y: top, width: legendBand, height: plot.height }
      : { x: left, y: chart.legend === "top" ? titleHeight + 12 : heightPt - legendBand, width: plot.width, height: legendBand };
  const layout: BasicChartLayout = Object.freeze({ plot, ...(title ? { title } : {}), ...(legend ? { legend: Object.freeze(legend) } : {}) });
  const boxes: { role: WorkbookChartTextRole; rect: BasicChartRect; rotation: number }[] = [];
  if (title) boxes.push({ role: "title", rect: title, rotation: 0 });
  if (categoryTitle) boxes.push({ role: "category-axis-title", rect: { x: left, y: categoryTop ? top - 16 - axisBand : bottom + 26, width: plot.width, height: axisBand }, rotation: 0 });
  if (valueTitle) boxes.push({ role: "value-axis-title", rect: { x: valueRight ? right + 28 : left - 28 - axisBand, y: bottom, width: plot.height, height: axisBand }, rotation: -90 });
  const inputs: WorkbookChartTitleContentInput[] = boxes.map(box => {
    const definition = box.role === "title" ? { reference: chart.titleReference } : box.role === "category-axis-title" ? chart.categoryAxisTitle! : chart.valueAxisTitle!;
    if (definition.reference && binding.titles[box.role] === undefined) throw new DocumentError("InvalidOperation", "A chart reference title has no current scalar display.");
    return { source, chart, role: box.role, widthPt: box.rect.width, heightPt: box.rect.height, ...(definition.reference ? { value: binding.titles[box.role]! } : {}) };
  });
  const texts = layoutWorkbookChartTitleContents(scope, inputs, { maxNodes, maxCodeUnits: maxUnits });
  let axisNodes = 0, axisUnits = 0;
  for (const [index, object] of texts.objects.entries()) {
    if (object.widthPt > boxes[index]!.rect.width || object.heightPt > boxes[index]!.rect.height) throw new DocumentError("ResourceLimitExceeded", "Auto-sized chart titles exceed their reserved text bands.");
    if (boxes[index]!.role !== "title") { axisNodes += object.nodes; axisUnits += object.element.textContent?.length ?? 0; }
  }
  if (axisNodes >= maxNodes || axisUnits >= maxUnits) throw new DocumentError("ResourceLimitExceeded", "Chart titles exhaust the graphic budget.");
  const main = boxes.findIndex(box => box.role === "title");
  const seriesIndexes = new Map(binding.data.series.map((series, index) => [series.seriesId, index]));
  const format = (value: number, definition?: WorkbookDisplayFormat): string => {
    const display = formatter.format(value, definition);
    if (display.kind === "overflow") return "####";
    if (display.parts.some(part => part.kind !== "text")) throw new DocumentError("UnsupportedFeature", "Chart number-format spacing requires measured label layout.");
    return display.parts.map(part => part.kind === "text" ? part.text : "").join("");
  };
  const linked = (style: CSSStyleDeclaration, kind: "axis" | "label"): boolean => {
    const flag = property(style, `--office-${kind}-number-format-source-linked`, "true");
    if (flag !== "true" && flag !== "false") throw new DocumentError("InvalidDocument", "Chart source-linked number-format flags are invalid.");
    return flag === "true";
  };
  const explicit = (style: CSSStyleDeclaration, kind: "axis" | "label", fallback?: string): WorkbookDisplayFormat | undefined => {
    const raw = property(style, `--office-${kind}-number-format`), code = raw ? cssStringValue(raw) : fallback;
    return code === undefined ? undefined : { code };
  };
  const graphics = createWorkbookChartView(source, chart, binding.data, widthPt, heightPt, id, {
    ...limits, maxNodes: maxNodes - axisNodes, maxCodeUnits: maxUnits - axisUnits, layout,
    ...(main < 0 ? {} : { title: texts.objects[main]!.element }),
    formatNumber(value, code, context) {
      if (!context) return format(value, code === undefined ? undefined : { code });
      if (context.kind === "value-axis") return format(value, linked(context.style, "axis") ? binding.valueAxisFormat : explicit(context.style, "axis", code));
      const sourceFormat = context.content === "percent" ? { code: "0%" } : binding.valueFormats[seriesIndexes.get(context.seriesId)!]?.[context.index];
      return format(value, linked(context.style, "label") ? sourceFormat : explicit(context.style, "label", context.content === "percent" ? "0%" : "General"));
    },
    formatCategory(index, style) {
      const value = binding.categoryNumbers[index];
      return value === null || value === undefined || linked(style, "axis") ? binding.data.categories[index]! : format(value, explicit(style, "axis"));
    },
  });
  for (const [index, object] of texts.objects.entries()) {
    const box = boxes[index]!;
    if (box.role === "title") continue;
    Object.assign(object.element.style, { position: "absolute", left: box.rect.x + "pt", top: box.rect.y + "pt", transformOrigin: "0 0", transform: `rotate(${box.rotation}deg)` });
    graphics.element.append(object.element);
  }
  return Object.freeze({ element: graphics.element, nodes: graphics.nodes + axisNodes, layout, texts: texts.objects, fonts: texts.fonts,
    usedFallbackLineMetrics: texts.usedFallbackLineMetrics, diagnostics: Object.freeze([...binding.diagnostics, ...graphics.diagnostics]),
    overflowingTitles: Object.freeze(boxes.filter((_, index) => texts.objects[index]!.overflow).map(box => box.role)) });
}
