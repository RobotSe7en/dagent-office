import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, PresentationChart, PresentationDataSet } from "@dagent-office/core";
import { markViewNode, ownViewAttribute } from "../view-nodes.js";
import { cssStringValue, numeric, property, resolvedColor, strokeAttributes, svgFill } from "./appearance.js";
import { svgElement as createSvgElement } from "./svg.js";

const palette = ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"];
export function formatChartNumber(value: number, format = "General"): string {
  if (format === "General" || !format) return String(Number(value.toPrecision(12)));
  const sections = format.split(";"), section = sections[value < 0 && sections.length > 1 ? 1 : value === 0 && sections.length > 2 ? 2 : 0]!;
  if (section === "") return "";
  if (!/^\(?[#,0]+(?:\.[0#]{1,20})?%?\)?$/.test(section)) throw new DocumentError("UnsupportedFeature", "This chart number format is outside the basic numeric format scope.");
  const pattern = section.replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1");
  const match = /^(.*?)([#,0]+)(?:\.([0#]+))?(%?)(.*?)$/.exec(pattern);
  if (!match) throw new DocumentError("UnsupportedFeature", "This chart number format is not supported by the preview.");
  const magnitude = Math.abs(value) * (match[4] ? 100 : 1), decimals = match[3] ?? "";
  const number = magnitude.toLocaleString("en-US", { useGrouping: match[2]!.includes(","), minimumFractionDigits: decimals.replace(/#/g, "").length, maximumFractionDigits: decimals.length });
  return (value < 0 && sections.length === 1 ? "-" : "") + match[1] + number + match[4] + match[5];
}
export function renderChart(node: HTMLElement, chart: DeepReadonly<PresentationChart>, data: DeepReadonly<PresentationDataSet>, width: number, height: number, serial: string, account: (amount: number) => void): void {
  const svgElement: typeof createSvgElement = (document, tag, attributes) => { account(1); return createSvgElement(document, tag, attributes); };
  const document = node.ownerDocument, view = document.defaultView!;
  let paintSerial = 0;
  const svg = markViewNode(svgElement(document, "svg", { width: "100%", height: "100%", viewBox: `0 0 ${width} ${height}`, role: "img", "aria-label": chart.kind + " chart", "data-sdk-chart": chart.kind }));
  svg.style.cssText = "position:absolute;inset:0;overflow:visible;pointer-events:none";
  node.prepend(svg);
  const probes: HTMLElement[] = [];
  const style = (classes?: string): CSSStyleDeclaration => {
    const probe = markViewNode(document.createElement("span")); probe.className = classes ?? ""; probe.style.cssText = "position:absolute;visibility:hidden;font-size:10pt";
    // A declared chart text style must override the native default.
    account(1); probe.style.removeProperty("font-size"); node.append(probe); probes.push(probe); return view.getComputedStyle(probe);
  };
  const base = style(chart.textCssClass), legendStyle = style([chart.textCssClass, chart.legendCssClass].filter(Boolean).join(" "));
  const font = (s: CSSStyleDeclaration): Record<string, string> => ({ "font-family": s.fontFamily, "font-size": String(parseFloat(s.fontSize) * 0.75), "font-weight": s.fontWeight, "font-style": s.fontStyle, fill: property(s, "--office-text-color") ? resolvedColor(document, s, "--office-text") : s.color });
  const text = (value: string, x: number, y: number, s = base, anchor = "middle"): SVGTextElement => {
    const label = svgElement(document, "text", { x, y, "text-anchor": anchor, "dominant-baseline": "middle", ...font(s) });
    if (/\p{Script=Han}/u.test(value)) { label.dataset.officeScript = "east-asia"; label.setAttribute("font-family", property(s, "--office-font-east-asia", property(s, "--office-minor-east-asia", s.fontFamily))); }
    const lines = value.split("\n");
    if (lines.length === 1) label.textContent = value;
    else lines.forEach((line, index) => { const span = svgElement(document, "tspan", { x, dy: index ? "1.2em" : String(-(lines.length - 1) * parseFloat(s.fontSize) * 0.75 * 0.6) }); span.textContent = line; label.append(span); });
    svg.append(label); return label;
  };
  const catStyle = style(chart.categoryAxisCssClass), valStyle = style(chart.valueAxisCssClass), gridStyle = style(chart.gridCssClass);
  const categoryTop = property(catStyle, "--office-axis-position", "b") === "t", valueRight = property(valStyle, "--office-axis-position", "l") === "r";
  const plotStyle = style(chart.plotCssClass);
  const title = chart.titleNodeId ? node.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(chart.titleNodeId) + '"]') : undefined;
  const titleHeight = title && chart.titleVisible !== false ? Math.min(54, height * 0.2) : 12;
  if (title) { ownViewAttribute(title, "data-sdk-chart-title"); title.setAttribute("data-sdk-chart-title", ""); }
  if (title) title.style.cssText += `;position:absolute;left:6pt;right:6pt;top:20pt;height:${titleHeight}pt;margin:0;text-align:center;${chart.titleVisible === false ? "display:none" : ""}`;
  const entries = chart.kind === "pie" ? data.categories.map((name, i) => ({ name, series: data.series[0]!, index: i })) : data.series.map(series => ({ name: series.name, series, index: -1 }));
  const legendBand = chart.legend === "none" ? 0 : chart.legend === "bottom" || chart.legend === "top" ? 28 : Math.min(width * 0.25, 100);
  const left = (chart.kind === "pie" || valueRight ? 12 : 28) + (chart.legend === "left" ? legendBand : 0);
  const top = titleHeight + (chart.kind === "pie" ? 15 : 18) + (chart.legend === "top" ? legendBand : 0) + (categoryTop ? 16 : 0);
  const right = width - (chart.kind === "pie" ? 12 : valueRight ? 28 : 16) - (chart.legend === "right" ? legendBand : 0), bottom = height - (chart.kind === "pie" ? 8 : categoryTop ? 12 : 26) - (chart.legend === "bottom" ? legendBand : 0);
  const w = Math.max(1, right - left), h = Math.max(1, bottom - top);
  svg.append(svgElement(document, "rect", { x: left, y: top, width: w, height: h, fill: svgFill(document, plotStyle, svg, serial + "-plot", "--office-fill", "none", account), ...strokeAttributes(document, plotStyle) }));
  const seriesStyle = new Map(data.series.map(series => [series.seriesId, style(chart.seriesStyles?.find(s => s.seriesId === series.seriesId)?.cssClass)]));
  const pointStyle = (seriesId: string, index: number): CSSStyleDeclaration => {
    const series = chart.seriesStyles?.find(s => s.seriesId === seriesId), point = chart.pointStyles?.find(p => p.seriesId === seriesId && p.index === index);
    return point ? style([series?.cssClass, point.cssClass].filter(Boolean).join(" ")) : seriesStyle.get(seriesId)!;
  };
  const fill = (seriesId: string, index: number, s: CSSStyleDeclaration): string => svgFill(document, s, svg, serial + "-paint-" + ++paintSerial, "--office-fill", palette[(chart.kind === "pie" ? index : data.series.findIndex(d => d.seriesId === seriesId)) % palette.length]!, account);
  const label = (series: DeepReadonly<PresentationDataSet>["series"][number], index: number, value: number, x: number, y: number, percent?: number, baseline?: number): void => {
    if (!chart.showLabels) return;
    const seriesDefinition = chart.seriesStyles?.find(s => s.seriesId === series.seriesId), point = chart.pointStyles?.find(p => p.seriesId === series.seriesId && p.index === index);
    const s = style([chart.textCssClass, chart.labelCssClass, seriesDefinition?.labelCssClass, point?.labelCssClass].filter(Boolean).join(" "));
    const contents = chart.labelContents ?? ["value"], separator = cssStringValue(property(s, "--office-label-separator", ", "));
    const values = contents.map(kind => kind === "category" ? data.categories[index]! : kind === "series" ? series.name : kind === "percent" ? formatChartNumber(percent ?? 0, "0%") : formatChartNumber(value, cssStringValue(property(s, "--office-label-number-format", chart.numberFormat ?? "General"))));
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
    const series = data.series[0]!, normalizer = series.values.reduce<number>((max, v) => Math.max(max, v ?? 0), 0) || 1;
    const sum = series.values.reduce<number>((total, v) => total + (v ?? 0) / normalizer, 0), cx = left + w / 2, cy = top + h / 2, radius = Math.min(w, h) / 2;
    let angle = (chart.firstSliceAngle ?? 0) * Math.PI / 180 - Math.PI / 2;
    series.values.forEach((value, index) => {
      if (value === null || value <= 0 || !sum) return;
      const fraction = value / normalizer / sum, sweep = fraction * Math.PI * 2, end = angle + sweep, middle = angle + sweep / 2;
      const s = pointStyle(series.seriesId, index), appearance = { fill: fill(series.seriesId, index, s), ...strokeAttributes(document, s) };
      if (sweep >= Math.PI * 2 - 1e-9) svg.append(svgElement(document, "circle", { cx, cy, r: radius, ...appearance }));
      else svg.append(svgElement(document, "path", { d: `M${cx} ${cy}L${cx + radius * Math.cos(angle)} ${cy + radius * Math.sin(angle)}A${radius} ${radius} 0 ${sweep > Math.PI ? 1 : 0} 1 ${cx + radius * Math.cos(end)} ${cy + radius * Math.sin(end)}Z`, ...appearance }));
      label(series, index, value, cx + radius * 0.65 * Math.cos(middle), cy + radius * 0.65 * Math.sin(middle), fraction); angle = end;
    });
  } else {
    let low = 0, high = 1;
    for (const series of data.series) for (const value of series.values) if (value !== null) { low = Math.min(low, value); high = Math.max(high, value); }
    const minimum = chart.minimum ?? low, largest = chart.maximum ?? high;
    const rough = (largest - minimum) / 8 || 1, power = 10 ** Math.floor(Math.log10(rough)), factor = rough / power, unit = chart.majorUnit ?? (factor <= 1 ? 1 : factor <= 2 ? 2 : factor <= 5 ? 5 : 10) * power;
    const maximum = chart.maximum ?? (Math.floor(largest / unit) + 1) * unit;
    if (!(maximum > minimum) || !Number.isFinite(unit) || (maximum - minimum) / unit > 1024) throw new DocumentError("ResourceLimitExceeded", "Chart axis ticks exceed the preview budget.");
    const y = (v: number): number => top + (chart.reverseValues ? (v - minimum) : (maximum - v)) / (maximum - minimum) * h;
    const axisX = valueRight ? right : left;
    const tick = (x: number, y: number, vertical: boolean, s: CSSStyleDeclaration, minor = false): void => {
      const kind = property(s, "--office-axis-" + (minor ? "minor" : "major") + "-tick", "cross"); if (kind === "none") return;
      const length = minor ? 2 : 4, start = kind === "in" ? 0 : -length, end = kind === "out" ? 0 : length;
      svg.append(svgElement(document, "line", { x1: x + (vertical ? 0 : start), x2: x + (vertical ? 0 : end), y1: y + (vertical ? start : 0), y2: y + (vertical ? end : 0), ...strokeAttributes(document, s, "--office-line", "#898989") }));
    };
    for (let value = minimum; value <= maximum + unit * 0.001; value += unit) {
      const position = y(value);
      if (chart.showValueAxis !== false) {
        const grid = strokeAttributes(document, gridStyle, "--office-line", "#d9d9d9");
        if (chart.gridCssClass) svg.append(svgElement(document, "line", { x1: left, y1: position, x2: right, y2: position, ...grid }));
        tick(axisX, position, false, valStyle);
        if (property(valStyle, "--office-axis-label-position") !== "none") text(formatChartNumber(value, chart.numberFormat), axisX + (valueRight ? 7 : -7), position, valStyle, valueRight ? "start" : "end");
      }
    }
    const minorUnit = numeric(valStyle, "--office-axis-minor-unit", unit / 5);
    if (minorUnit > 0 && chart.showValueAxis !== false) {
      if ((maximum - minimum) / minorUnit > 4096) throw new DocumentError("ResourceLimitExceeded", "Chart minor ticks exceed the preview budget.");
      for (let value = minimum; value <= maximum; value += minorUnit) if (Math.abs((value - minimum) / unit - Math.round((value - minimum) / unit)) > 0.001) tick(axisX, y(value), false, valStyle, true);
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
        const value = raw ?? (chart.blankAs === "zero" ? 0 : null);
        if (value === null) { if (chart.blankAs !== "span") connected = false; return; }
        const px = x(index), py = y(value), point = pointStyle(series.seriesId, index);
        if (chart.kind === "column") svg.append(svgElement(document, "rect", { x: px - total / 2 + seriesIndex * bar * (1 - overlap), y: Math.min(py, baseline), width: bar, height: Math.abs(py - baseline), fill: fill(series.seriesId, index, point), ...strokeAttributes(document, point) }));
        else {
          segments.push((connected ? "L" : "M") + px + " " + py); connected = true;
          const definition = chart.seriesStyles?.find(d => d.seriesId === series.seriesId), marker = definition?.marker ?? "none", radius = (definition?.markerSize ?? 5) / 2;
          if (marker !== "none") {
            const color = fill(series.seriesId, index, point);
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
      else svg.append(svgElement(document, "rect", { x, y: y - 3, width: keyWidth, height: 7, fill: fill(entry.series.seriesId, Math.max(0, entry.index), s) }));
      text(entry.name, x + keyWidth + 3, y + 1, legendStyle, "start");
    });
  }
  for (const probe of probes) probe.remove();
}
