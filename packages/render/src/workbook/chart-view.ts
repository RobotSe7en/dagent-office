import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookChart } from "@dagent-office/core";
import { renderBasicChart } from "../chart.js";
import type { BasicChartData } from "../chart.js";
import { markViewNode } from "../view-nodes.js";

export interface WorkbookChartViewOptions {
  /** Required: the workbook compositor owns number formats, locale and dates.
   * This adapter must never inherit the presentation formatter's policy. */
  readonly formatNumber: (value: number, format?: string) => string;
  /** A fresh, already materialized title view; never a source/active view node. */
  readonly title?: HTMLElement;
  readonly maxNodes?: number; readonly maxPoints?: number; readonly maxCodeUnits?: number;
}
export interface WorkbookChartView { readonly element: HTMLElement; readonly nodes: number; readonly peakNodes: number }

/** Materialize graphics from already bound and formatted plot data. Resolve
 * CSS against the connected source and remove all temporary probes on failure.
 * The detached view carries no source classes or source DOM IDs. Frame effects,
 * title materialization, source projections and pane placement belong to the compositor. */
export function createWorkbookChartView(source: HTMLElement, chart: DeepReadonly<WorkbookChart>, data: BasicChartData, widthPt: number, heightPt: number, id: string, options: WorkbookChartViewOptions): WorkbookChartView {
  const document = source.ownerDocument, view = document.defaultView;
  const maxNodes = options.maxNodes ?? 50_000, maxPoints = options.maxPoints ?? 50_000, maxCodeUnits = options.maxCodeUnits ?? 8 * 1024 * 1024;
  if (!source.isConnected || !view || ![maxNodes, maxPoints, maxCodeUnits].every(value => Number.isSafeInteger(value) && value > 0)
    || ![widthPt, heightPt].every(value => Number.isFinite(value) && value > 0 && value <= 1_000_000) || !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(id)) throw new DocumentError("InvalidOperation", "Chart graphics require a connected source, finite positive dimensions, a local SVG identity and positive budgets.");
  if (source.dataset.nodeId !== chart.nodeId || source.dataset.officeKind !== "chart") throw new DocumentError("InvalidDocument", "Chart graphics require their own source object.");
  const title = options.title;
  if (title && (title.ownerDocument !== document || title.isConnected || title.parentNode)) throw new DocumentError("InvalidOperation", "A chart title must be a fresh detached view in its own document.");
  if (chart.titleVisible !== false && (chart.titleNodeId || chart.titleReference) && !title) throw new DocumentError("InvalidOperation", "A visible chart title requires a materialized title view.");
  if (data.series.length !== chart.series.length || data.series.some((series, index) => series.seriesId !== chart.series[index]!.seriesId || series.values.length !== data.categories.length
    || series.values.some(value => value !== null && !Number.isFinite(value)))) throw new DocumentError("InvalidOperation", "Derived chart values must match their source series and category positions.");
  let units = 0, points = data.categories.length;
  const accountText = (amount: number): void => { units += amount; if (units > maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Chart text exceeds its source and display budget."); };
  for (const name of data.categories) accountText(name.length);
  for (const series of data.series) { accountText(series.name.length); points += series.values.length; }
  if (points > maxPoints) throw new DocumentError("ResourceLimitExceeded", "Derived chart data exceeds its point budget.");
  const element = markViewNode(document.createElement("div")), probes: HTMLElement[] = [];
  element.dataset.officeView = "workbook-chart"; element.dataset.sourceNodeId = chart.nodeId;
  element.style.cssText = `position:absolute;width:${widthPt}pt;height:${heightPt}pt`;
  let nodes = 1, peakNodes = 1;
  const account = (amount: number): void => {
    nodes += amount; peakNodes = Math.max(peakNodes, nodes);
    if (peakNodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Chart graphics exceed their node budget.");
  };
  try {
    if (title) {
      const walker = document.createTreeWalker(title, 0xffff_ffff);
      do { account(1); if (walker.currentNode.nodeType === 3) accountText(walker.currentNode.textContent!.length); } while (walker.nextNode());
      element.append(title);
    }
    renderBasicChart(element, chart, data, widthPt, heightPt, id, account, { formatNumber: options.formatNumber, accountText, ...(title ? { title } : {}), style(classes) {
      account(1);
      const probe = markViewNode(document.createElement("span")); probe.className = classes ?? "";
      probe.style.cssText = "position:absolute;visibility:hidden"; source.append(probe); probes.push(probe); return view.getComputedStyle(probe);
    } });
    return Object.freeze({ element, nodes: nodes - probes.length, peakNodes });
  } catch (error) { element.replaceChildren(); throw error; }
  finally { for (const probe of probes) probe.remove(); }
}
