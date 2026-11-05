import type { ChartAppearance, DeepReadonly } from "@dagent-office/core";
import { DocumentError } from "@dagent-office/core";
import { markViewNode, ownViewAttribute } from "./view-nodes.js";
import { cssStringValue, numeric, property, resolvedColor, strokeAttributes, svgFill } from "./presentation/appearance.js";
import { svgElement as createSvgElement } from "./presentation/svg.js";
import type { DrawingGradientBounds } from "./drawing-gradient.js";
import { chartAxisTicks } from "./chart-axis.js";

/** Derived plot values. The owning format adapter supplies its authoritative
 * source binding, visibility, coercion and explicit number-display policy. */
export interface BasicChartData {
  readonly categories: readonly string[];
  /** Original source positions survive hidden-point compaction. */
  readonly sourceIndexes?: readonly number[];
  readonly series: readonly {
    readonly seriesId: string; readonly name: string; readonly values: readonly (number | null)[];
    readonly hidden?: boolean;
    /** Explicit missing-point behavior overrides the ordinary blank policy. */
    readonly missing?: readonly ("gap" | "span" | undefined)[];
  }[];
}
export interface BasicChartOptions {
  /** Resolve styles against the source ancestry, which may be outside the view.
   * The adapter owns any temporary probes and releases them even on failure. */
  readonly style: (classes?: string) => CSSStyleDeclaration;
  readonly formatNumber: (value: number, format?: string) => string;
  readonly accountText?: (codeUnits: number) => void;
  readonly title?: HTMLElement;
  /** The owning format selects negative-value semantics explicitly. Value
   * labels keep their signed source numbers under the magnitude policy. */
  readonly pieValues?: "nonnegative" | "magnitude";
}
const palette = ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"];
export function renderBasicChart(node: HTMLElement, chart: DeepReadonly<ChartAppearance>, data: BasicChartData, width: number, height: number, serial: string, account: (amount: number) => void, options: BasicChartOptions): void {
  const svgElement: typeof createSvgElement = (document, tag, attributes) => { account(1); return createSvgElement(document, tag, attributes); };
  const document = node.ownerDocument, formatNumber = options.formatNumber, styles = new Map<string, CSSStyleDeclaration>();
  // Validated source CSS has no structural pseudo-classes. One probe per
  // class list is enough for this immutable render, including repeated labels.
  const style = (classes?: string): CSSStyleDeclaration => {
    const key = classes ?? ""; let computed = styles.get(key);
    if (!computed) { computed = options.style(classes); styles.set(key, computed); }
    return computed;
  };
  const seriesDefinitions = new Map((chart.seriesStyles ?? []).map(series => [series.seriesId, series]));
  const pointKey = (seriesId: string, index: number): string => JSON.stringify([seriesId, index]);
  const pointDefinitions = new Map((chart.pointStyles ?? []).map(point => [pointKey(point.seriesId, point.index), point]));
  const seriesIndexes = new Map(data.series.map((series, index) => [series.seriesId, index]));
  const sourceIndex = (index: number): number => data.sourceIndexes?.[index] ?? index;
  data = { ...data, series: data.series.filter(series => !series.hidden) };
  let paintSerial = 0;
  const svg = markViewNode(svgElement(document, "svg", { width: "100%", height: "100%", viewBox: `0 0 ${width} ${height}`, role: "img", "aria-label": chart.kind + " chart", "data-sdk-chart": chart.kind }));
  svg.style.cssText = "position:absolute;inset:0;overflow:visible;pointer-events:none";
  node.prepend(svg);
  const base = style(chart.textCssClass), legendStyle = style([chart.textCssClass, chart.legendCssClass].filter(Boolean).join(" "));
  const font = (s: CSSStyleDeclaration): Record<string, string> => ({ "font-family": s.fontFamily, "font-size": String(parseFloat(s.fontSize) * 0.75), "font-weight": s.fontWeight, "font-style": s.fontStyle, fill: property(s, "--office-text-color") ? resolvedColor(document, s, "--office-text") : s.color });
  const text = (value: string, x: number, y: number, s = base, anchor = "middle"): SVGTextElement => {
    options.accountText?.(value.length);
    const label = svgElement(document, "text", { x, y, "text-anchor": anchor, "dominant-baseline": "middle", ...font(s) });
    if (/\p{Script=Han}/u.test(value)) { label.dataset.officeScript = "east-asia"; label.setAttribute("font-family", property(s, "--office-font-east-asia", property(s, "--office-minor-east-asia", s.fontFamily))); }
    const lines = value.split("\n");
    if (lines.length === 1) { if (value) account(1); label.textContent = value; }
    else lines.forEach((line, index) => { const span = svgElement(document, "tspan", { x, dy: index ? "1.2em" : String(-(lines.length - 1) * parseFloat(s.fontSize) * 0.75 * 0.6) }); if (line) account(1); span.textContent = line; label.append(span); });
    svg.append(label); return label;
  };
  const catStyle = style(chart.categoryAxisCssClass), valStyle = style(chart.valueAxisCssClass), gridStyle = style(chart.gridCssClass);
  const categoryTop = property(catStyle, "--office-axis-position", "b") === "t", valueRight = property(valStyle, "--office-axis-position", "l") === "r";
  const plotStyle = style(chart.plotCssClass);
  const title = options.title;
  const titleHeight = title && chart.titleVisible !== false ? Math.min(54, height * 0.2) : 12;
  if (title) { ownViewAttribute(title, "data-sdk-chart-title"); title.setAttribute("data-sdk-chart-title", ""); }
  if (title) title.style.cssText += `;position:absolute;left:6pt;right:6pt;top:20pt;height:${titleHeight}pt;margin:0;text-align:center;${chart.titleVisible === false ? "display:none" : ""}`;
  const entries = chart.kind === "pie" ? data.categories.map((name, i) => ({ name, series: data.series[0]!, index: i })) : data.series.map(series => ({ name: series.name, series, index: -1 }));
  const legendBand = chart.legend === "none" ? 0 : chart.legend === "bottom" || chart.legend === "top" ? 28 : Math.min(width * 0.25, 100);
  const left = (chart.kind === "pie" || valueRight ? 12 : 28) + (chart.legend === "left" ? legendBand : 0);
  const top = titleHeight + (chart.kind === "pie" ? 15 : 18) + (chart.legend === "top" ? legendBand : 0) + (categoryTop ? 16 : 0);
  const right = width - (chart.kind === "pie" ? 12 : valueRight ? 28 : 16) - (chart.legend === "right" ? legendBand : 0), bottom = height - (chart.kind === "pie" ? 8 : categoryTop ? 12 : 26) - (chart.legend === "bottom" ? legendBand : 0);
  const w = Math.max(1, right - left), h = Math.max(1, bottom - top);
  svg.append(svgElement(document, "rect", { x: left, y: top, width: w, height: h, fill: svgFill(document, plotStyle, svg, serial + "-plot", { x: left, y: top, width: w, height: h }, "--office-fill", "none", account), ...strokeAttributes(document, plotStyle) }));
  if (!data.series.length) return;
  const seriesStyle = new Map(data.series.map(series => [series.seriesId, style(seriesDefinitions.get(series.seriesId)?.cssClass)]));
  const pointStyle = (seriesId: string, index: number): CSSStyleDeclaration => {
    const series = seriesDefinitions.get(seriesId), point = pointDefinitions.get(pointKey(seriesId, sourceIndex(index)));
    return point ? style([series?.cssClass, point.cssClass].filter(Boolean).join(" ")) : seriesStyle.get(seriesId)!;
  };
  const fill = (seriesId: string, index: number, s: CSSStyleDeclaration, bounds: DrawingGradientBounds): string => svgFill(document, s, svg, serial + "-paint-" + ++paintSerial, bounds, "--office-fill", palette[(chart.kind === "pie" ? sourceIndex(index) : seriesIndexes.get(seriesId)!) % palette.length]!, account);
  const label = (series: BasicChartData["series"][number], index: number, value: number, x: number, y: number, percent?: number, baseline?: number): void => {
    if (!chart.showLabels) return;
    const seriesDefinition = seriesDefinitions.get(series.seriesId), point = pointDefinitions.get(pointKey(series.seriesId, sourceIndex(index)));
    const s = style([chart.textCssClass, chart.labelCssClass, seriesDefinition?.labelCssClass, point?.labelCssClass].filter(Boolean).join(" "));
    const contents = chart.labelContents ?? ["value"], separator = cssStringValue(property(s, "--office-label-separator", ", "));
    const values = contents.map(kind => kind === "category" ? data.categories[index]! : kind === "series" ? series.name : kind === "percent" ? formatNumber(percent ?? 0, "0%") : formatNumber(value, cssStringValue(property(s, "--office-label-number-format", chart.numberFormat ?? "General"))));
    const position = property(s, "--office-label-position", chart.kind === "pie" ? "ctr" : "outEnd");
    let anchor = "middle";
    if (chart.kind === "column" && baseline !== undefined) {
      const direction = Math.sign(baseline - y) || 1;
      y = position === "ctr" ? (y + baseline) / 2 : position === "inBase" ? baseline - direction * 8 : position === "inEnd" ? y + direction * 8 : y - direction * 8;
    } else if (chart.kind === "line") {
      if (position === "l" || position === "r") { x += position === "l" ? -8 : 8; anchor = position === "l" ? "end" : "start"; }
      else if (position !== "ctr") y += position === "b" ? 8 : -8;
    } else if (chart.kind === "pie" && ["outEnd", "inEnd"].includes(position)) {
      const cx = left + w / 2, cy = top + h / 2, factor = position === "outEnd" ? 1.7 : 1.3; x = cx + (x - cx) * factor; y = cy + (y - cy) * factor;
    }
    text(values.join(separator), x, y, s, anchor);
  };
  if (chart.kind === "pie") {
    const series = data.series[0]!;
    const magnitude = (value: number | null): number => {
      if (value === null) return 0;
      if (value < 0 && options.pieValues !== "magnitude") throw new DocumentError("InvalidDocument", "Negative pie values require an explicit magnitude policy.");
      return Math.abs(value);
    };
    const normalizer = series.values.reduce<number>((max, value) => Math.max(max, magnitude(value)), 0) || 1;
    const sum = series.values.reduce<number>((total, value) => total + magnitude(value) / normalizer, 0), cx = left + w / 2, cy = top + h / 2, radius = Math.min(w, h) / 2;
    let angle = (chart.firstSliceAngle ?? 0) * Math.PI / 180 - Math.PI / 2;
    series.values.forEach((value, index) => {
      if (value === null || value === 0 || !sum) return;
      const fraction = magnitude(value) / normalizer / sum, sweep = fraction * Math.PI * 2, end = angle + sweep, middle = angle + sweep / 2;
      const points = [[cx, cy], [cx + radius * Math.cos(angle), cy + radius * Math.sin(angle)], [cx + radius * Math.cos(end), cy + radius * Math.sin(end)]];
      for (let quadrant = Math.ceil(angle / (Math.PI / 2)); quadrant * Math.PI / 2 <= end; quadrant++) points.push([cx + radius * Math.cos(quadrant * Math.PI / 2), cy + radius * Math.sin(quadrant * Math.PI / 2)]);
      const xs = points.map(point => point[0]!), ys = points.map(point => point[1]!);
      const bounds = { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) };
      const s = pointStyle(series.seriesId, index), appearance = { fill: fill(series.seriesId, index, s, bounds), ...strokeAttributes(document, s) };
      if (sweep >= Math.PI * 2 - 1e-9) svg.append(svgElement(document, "circle", { cx, cy, r: radius, ...appearance }));
      else svg.append(svgElement(document, "path", { d: `M${cx} ${cy}L${cx + radius * Math.cos(angle)} ${cy + radius * Math.sin(angle)}A${radius} ${radius} 0 ${sweep > Math.PI ? 1 : 0} 1 ${cx + radius * Math.cos(end)} ${cy + radius * Math.sin(end)}Z`, ...appearance }));
      label(series, index, value, cx + radius * 0.65 * Math.cos(middle), cy + radius * 0.65 * Math.sin(middle), fraction); angle = end;
    });
  } else {
    let low = 0, high = 0;
    for (const series of data.series) for (const value of series.values) if (value !== null) { low = Math.min(low, value); high = Math.max(high, value); }
    if (low === 0 && high === 0) high = 1;
    const minimum = chart.minimum ?? low, largest = chart.maximum ?? high;
    const rough = (largest - minimum) / 8 || 1, power = 10 ** Math.floor(Math.log10(rough)), factor = rough / power, unit = chart.majorUnit ?? (factor <= 1 ? 1 : factor <= 2 ? 2 : factor <= 5 ? 5 : 10) * power;
    const maximum = chart.maximum ?? (Math.floor(largest / unit) + 1) * unit;
    const majorTicks = chartAxisTicks(minimum, maximum, unit, 1024, 0.001);
    const y = (v: number): number => {
      const position = top + (chart.reverseValues ? (v - minimum) : (maximum - v)) / (maximum - minimum) * h;
      if (!Number.isFinite(position)) throw new DocumentError("ResourceLimitExceeded", "Chart values exceed the finite preview coordinate scale.");
      return position;
    };
    const axisX = valueRight ? right : left;
    const tick = (x: number, y: number, vertical: boolean, s: CSSStyleDeclaration, minor = false): void => {
      const kind = property(s, "--office-axis-" + (minor ? "minor" : "major") + "-tick", "cross"); if (kind === "none") return;
      const length = minor ? 2 : 4, start = kind === "in" ? 0 : -length, end = kind === "out" ? 0 : length;
      svg.append(svgElement(document, "line", { x1: x + (vertical ? 0 : start), x2: x + (vertical ? 0 : end), y1: y + (vertical ? start : 0), y2: y + (vertical ? end : 0), ...strokeAttributes(document, s, "--office-line", "#898989") }));
    };
    for (const value of majorTicks) {
      const position = y(value);
      if (chart.showValueAxis !== false) {
        const grid = strokeAttributes(document, gridStyle, "--office-line", "#d9d9d9");
        if (chart.gridCssClass) svg.append(svgElement(document, "line", { x1: left, y1: position, x2: right, y2: position, ...grid }));
        tick(axisX, position, false, valStyle);
        if (property(valStyle, "--office-axis-label-position") !== "none") text(formatNumber(value, chart.numberFormat), axisX + (valueRight ? 7 : -7), position, valStyle, valueRight ? "start" : "end");
      }
    }
    const minorUnit = numeric(valStyle, "--office-axis-minor-unit", unit / 5);
    if (minorUnit > 0 && chart.showValueAxis !== false) {
      for (const value of chartAxisTicks(minimum, maximum, minorUnit, 4096)) if (Math.abs((value - minimum) / unit - Math.round((value - minimum) / unit)) > 0.001) tick(axisX, y(value), false, valStyle, true);
    }
    const baseline = y(Math.max(minimum, Math.min(maximum, 0))), axisY = categoryTop ? top : baseline;
    if (chart.showCategoryAxis !== false) svg.append(svgElement(document, "line", { x1: left, y1: axisY, x2: right, y2: axisY, ...strokeAttributes(document, catStyle, "--office-line", "#b7b7b7") }));
    if (chart.showValueAxis !== false) svg.append(svgElement(document, "line", { x1: axisX, y1: top, x2: axisX, y2: bottom, ...strokeAttributes(document, valStyle, "--office-line", "#b7b7b7") }));
    const count = Math.max(1, data.categories.length), group = w / count;
    const x = (index: number): number => left + (chart.reverseCategories ? count - index - 0.5 : index + 0.5) * group;
    if (chart.showCategoryAxis !== false) {
      const skip = Math.max(1, numeric(catStyle, "--office-axis-label-skip", 1)), tickSkip = Math.max(1, numeric(catStyle, "--office-axis-tick-skip", 1));
      const offset = 6 + 10 * numeric(catStyle, "--office-axis-label-offset", 100) / 100;
      data.categories.forEach((category, index) => { if (index % skip === 0 && property(catStyle, "--office-axis-label-position") !== "none") text(category, x(index), categoryTop ? top - offset : bottom + offset, catStyle); });
      if (chart.kind === "line") {
        for (let n = 0; n < count; n += tickSkip) { tick(x(n), axisY, true, catStyle); if (n + 1 < count) tick((x(n) + x(n + 1)) / 2, axisY, true, catStyle, true); }
      } else for (let n = 0; n <= count; n += tickSkip) tick(left + n * group, axisY, true, catStyle);
    }
    data.series.forEach((series, seriesIndex) => {
      const s = seriesStyle.get(series.seriesId)!, segments: string[] = [];
      const overlap = (chart.overlap ?? 0) / 100, slots = data.series.length - (data.series.length - 1) * overlap, bar = group / (slots + (chart.gapWidth ?? 150) / 100), total = slots * bar;
      let connected = false;
      series.values.forEach((raw, index) => {
        if (raw === null && series.missing?.[index]) { if (series.missing[index] === "gap") connected = false; return; }
        const value = raw ?? (chart.blankAs === "zero" ? 0 : null);
        if (value === null) { if (chart.blankAs !== "span") connected = false; return; }
        const px = x(index), py = y(value), point = pointStyle(series.seriesId, index);
        if (chart.kind === "column") {
          const bounds = { x: px - total / 2 + seriesIndex * bar * (1 - overlap), y: Math.min(py, baseline), width: bar, height: Math.abs(py - baseline) };
          svg.append(svgElement(document, "rect", { ...bounds, fill: fill(series.seriesId, index, point, bounds), ...strokeAttributes(document, point) }));
        }
        else {
          segments.push((connected ? "L" : "M") + px + " " + py); connected = true;
          const definition = seriesDefinitions.get(series.seriesId), marker = definition?.marker ?? "none", radius = (definition?.markerSize ?? 5) / 2;
          if (marker !== "none") {
            const extent = marker === "dot" ? 1 : radius;
            const color = fill(series.seriesId, index, point, { x: px - extent, y: py - extent, width: extent * 2, height: extent * 2 });
            if (marker === "circle" || marker === "dot") svg.append(svgElement(document, "circle", { cx: px, cy: py, r: marker === "dot" ? 1 : radius, fill: color }));
            else if (marker === "square") svg.append(svgElement(document, "rect", { x: px - radius, y: py - radius, width: radius * 2, height: radius * 2, fill: color }));
            else {
              const d = marker === "triangle" ? `M${px} ${py - radius}L${px + radius} ${py + radius}L${px - radius} ${py + radius}Z` : marker === "diamond" ? `M${px} ${py - radius}L${px + radius} ${py}L${px} ${py + radius}L${px - radius} ${py}Z` : marker === "dash" ? `M${px - radius} ${py}L${px + radius} ${py}` : marker === "x" ? `M${px - radius} ${py - radius}L${px + radius} ${py + radius}M${px - radius} ${py + radius}L${px + radius} ${py - radius}` : `M${px - radius} ${py}L${px + radius} ${py}M${px} ${py - radius}L${px} ${py + radius}`;
              svg.append(svgElement(document, "path", { d, fill: marker === "triangle" || marker === "diamond" ? color : "none", stroke: color }));
            }
          }
        }
        label(series, index, value, chart.kind === "column" ? px - total / 2 + seriesIndex * bar * (1 - overlap) + bar / 2 : px, py, undefined, baseline);
      });
      if (chart.kind === "line") svg.append(svgElement(document, "path", { d: segments.join(" "), fill: "none", ...strokeAttributes(document, s, "--office-line", palette[seriesIndex % palette.length]!) }));
    });
  }
  if (chart.legend !== "none") {
    const horizontal = chart.legend === "top" || chart.legend === "bottom", slot = 18;
    const context = document.createElement("canvas").getContext("2d")!; context.font = legendStyle.font || legendStyle.fontSize + " " + legendStyle.fontFamily;
    const keyWidth = chart.kind === "line" ? 22 : 7, widths = entries.map(entry => context.measureText(entry.name).width * 0.75 + keyWidth + 8), total = widths.reduce((sum, n) => sum + n, 0);
    let cursor = left + (w - total) / 2;
    entries.forEach((entry, index) => {
      const x = horizontal ? cursor : chart.legend === "left" ? 12 : right + 12; cursor += widths[index]!;
      const y = horizontal ? chart.legend === "top" ? top - 14 : height - 14 : top + h / 2 + (index - (entries.length - 1) / 2) * slot;
      const s = entry.index >= 0 ? pointStyle(entry.series.seriesId, entry.index) : seriesStyle.get(entry.series.seriesId)!;
      if (chart.kind === "line") svg.append(svgElement(document, "line", { x1: x, y1: y, x2: x + keyWidth, y2: y, ...strokeAttributes(document, s, "--office-line", palette[index % palette.length]!) }));
      else svg.append(svgElement(document, "rect", { x, y: y - 3, width: keyWidth, height: 7, fill: fill(entry.series.seriesId, Math.max(0, entry.index), s, { x, y: y - 3, width: keyWidth, height: 7 }) }));
      text(entry.name, x + keyWidth + 3, y + 1, legendStyle, "start");
    });
  }
}
