import type { PresentationChart, PresentationDataSet } from "@dagent-office/core";
import { A, attr, C, check, child, elements, fail, R, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { bool, Context, integer, value } from "./context.js";
import type { Html } from "./context.js";
import { effects, fill, outline } from "./styles.js";
import { importText, paragraphProperties, runProperties } from "./text.js";
import { cachedVector, constants } from "./workbook.js";

function val(node: XmlNode | undefined, name: string, fallback?: string): string | undefined { return value(child(node, name, C)) ?? fallback; }
function flag(node: XmlNode | undefined, name: string, fallback = false): boolean { const source = child(node, name, C); return source ? bool(source, "val", true) : fallback; }
function decimal(node: XmlNode, name: string): number | undefined {
  const property = child(node, name, C);
  if (!property) return undefined;
  const source = value(property), n = Number(source);
  if (!source?.trim() || !Number.isFinite(n)) fail(property, "A chart numeric option is invalid.");
  return n;
}
function appearance(node: XmlNode | undefined): Css {
  if (!node) return {};
  check(node, ["noFill", "solidFill", "gradFill", "ln", "effectLst", "extLst"], A);
  return { ...fill(node), ...outline(child(node, "ln", A)), ...effects(child(node, "effectLst", A)) };
}
function textStyle(node: XmlNode | undefined): Css {
  if (!node) return {};
  const p = child(node, "p", A);
  return { ...paragraphProperties(child(p, "pPr", A)), ...runProperties(child(p, "endParaRPr", A)) };
}
function style(context: Context, source: XmlNode | undefined, prefix: string, extra: Css = {}): string {
  const name = context.id(prefix);
  context.rule(name, { ...appearance(child(source, "spPr", C)), ...textStyle(child(source, "txPr", C)), ...extra }); return name;
}
function noLayout(node: XmlNode | undefined): void {
  const layout = child(node, "layout", C);
  if (layout && elements(layout, undefined, C).length) unsupported(layout, "Manual chart layout is outside the basic chart scope.");
}
function labels(node: XmlNode | undefined, defaults: NonNullable<PresentationChart["labelContents"]> = []): NonNullable<PresentationChart["labelContents"]> {
  if (!node) return defaults;
  check(node, ["dLbl", "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines", "leaderLines", "extLst"], C);
  if (flag(node, "delete")) return [];
  if (flag(node, "showLegendKey") || flag(node, "showBubbleSize")) unsupported(node, "Legend keys and bubble labels are not supported.");
  const result: NonNullable<PresentationChart["labelContents"]> = [];
  for (const [name, kind] of [["showVal", "value"], ["showCatName", "category"], ["showSerName", "series"], ["showPercent", "percent"]] as const) if (flag(node, name, defaults.includes(kind))) result.push(kind);
  for (const point of elements(node, "dLbl", C)) {
    const selected = labels({ ...point, content: point.content.filter(n => typeof n === "string" || n.uri !== C || n.name !== "idx") }, result);
    if (JSON.stringify(selected) !== JSON.stringify(result)) unsupported(point, "Per-point label visibility is outside the basic chart scope.");
  }
  return result;
}
function cssString(value: string): string { return '"' + value.replace(/["\\\u0000-\u001f\u007f]/g, c => c === '"' || c === "\\" ? "\\" + c : "\\" + c.charCodeAt(0).toString(16) + " ") + '"'; }
function labelStyle(context: Context, source: XmlNode | undefined): string | undefined {
  if (!source) return undefined;
  const css: Css = {}, format = child(source, "numFmt", C), separator = child(source, "separator", C), position = val(source, "dLblPos");
  if (format && value(format, "formatCode") !== undefined) css["--office-label-number-format"] = cssString(value(format, "formatCode")!);
  if (separator) css["--office-label-separator"] = cssString(text(separator));
  if (position) css["--office-label-position"] = position;
  return style(context, source, "chart-label-style", css);
}
function axisStyle(context: Context, source: XmlNode, prefix: string): string {
  const css: Css = {};
  for (const [name, key, fallback] of [["axPos", "position", source.name === "catAx" ? "b" : "l"], ["majorTickMark", "major-tick", "none"], ["minorTickMark", "minor-tick", "none"], ["tickLblPos", "label-position", "nextTo"], ["crosses", "crosses", "autoZero"], ["lblAlgn", "label-align", "ctr"], ["lblOffset", "label-offset", "100"], ["tickLblSkip", "label-skip", "1"], ["tickMarkSkip", "tick-skip", "1"], ["crossBetween", "cross-between", "between"], ["minorUnit", "minor-unit", "auto"]]) css["--office-axis-" + key] = val(source, name!, fallback!)!;
  return style(context, source, prefix, css);
}
export function importChart(context: Context, reference: XmlNode, frame: Html): string {
  const path = context.relation(reference.part, attr(reference, "id", R), "chart").target;
  const source = context.part(path, "drawingml.chart", "chartSpace", C);
  check(source, ["date1904", "lang", "roundedCorners", "style", "clrMapOvr", "chart", "spPr", "txPr", "externalData", "printSettings", "extLst"], C);
  if (flag(source, "roundedCorners")) unsupported(source, "Rounded chart frames are not supported.");
  if (child(source, "style", C)) unsupported(source, "Preset chart style IDs require an explicit style mapping.");
  const external = child(source, "externalData", C);
  let book: ReturnType<typeof constants> | undefined, workbookPath: string | undefined;
  if (external) {
    workbookPath = context.relation(path, attr(external, "id", R), "package").target;
    if (context.opc.contentType(workbookPath) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") unsupported(external, "Chart data must be an embedded XLSX constant table.");
    book = constants(context, workbookPath);
    if (flag(external, "autoUpdate")) unsupported(external, "Automatic chart data refresh is not supported.");
  }
  const chart = child(source, "chart", C), plot = child(chart, "plotArea", C);
  if (!chart || !plot) fail(source, "A chart plot is missing.");
  check(chart, ["title", "autoTitleDeleted", "plotArea", "legend", "plotVisOnly", "dispBlanksAs", "showDLblsOverMax", "extLst"], C);
  check(plot, ["layout", "barChart", "lineChart", "pieChart", "catAx", "valAx", "spPr", "extLst"], C);
  noLayout(plot);
  const plots = elements(plot, undefined, C).filter(n => ["barChart", "lineChart", "pieChart"].includes(n.name));
  if (plots.length !== 1) unsupported(plot, "Combination and advanced charts are not supported.");
  const native = plots[0]!, kind = native.name === "barChart" ? "column" : native.name === "lineChart" ? "line" : "pie";
  check(native, kind === "column" ? ["barDir", "grouping", "varyColors", "ser", "dLbls", "gapWidth", "overlap", "axId", "extLst"] : kind === "line" ? ["grouping", "varyColors", "ser", "dLbls", "hiLowLines", "marker", "smooth", "axId", "extLst"] : ["varyColors", "ser", "dLbls", "firstSliceAng", "extLst"], C);
  if (kind === "column" && (val(native, "barDir", "col") !== "col" || val(native, "grouping", "clustered") !== "clustered") || kind === "line" && val(native, "grouping", "standard") !== "standard") unsupported(native, "Only clustered columns, simple lines and single-series pies are supported.");
  if (flag(native, "smooth")) unsupported(native, "Smoothed lines are not supported.");
  const highLow = child(native, "hiLowLines", C);
  if (highLow && !child(child(child(highLow, "spPr", C), "ln", A), "noFill", A)) unsupported(highLow, "High/low chart lines are not supported.");
  const id = context.id("chart"), definition: PresentationChart = { chartId: id, nodeId: frame.attrs["data-node-id"]!, dataSetId: "", kind, legend: "none", showLabels: false };
  const series = elements(native, "ser", C);
  if (!series.length || kind === "pie" && series.length !== 1) unsupported(native, "This chart series arrangement is not supported.");
  context.count(series.length);
  const data: PresentationDataSet = { dataSetId: context.id("data"), categories: [], series: [] };
  const indices = new Set<number>(), orders = new Set<number>();
  const ordered = series.map(s => ({ source: s, order: integer(child(s, "order", C) ?? s, "val") })).sort((a, b) => a.order - b.order);
  definition.seriesStyles = []; definition.pointStyles = [];
  const globalLabels = child(native, "dLbls", C);
  const globalLabelContents = labels(globalLabels), globalLabelStyle = labelStyle(context, globalLabels);
  if (globalLabelStyle) definition.labelCssClass = globalLabelStyle;
  let commonLabels: string | undefined;
  for (const [index, item] of ordered.entries()) {
    const s = item.source;
    check(s, ["idx", "order", "tx", "spPr", "invertIfNegative", "marker", "dPt", "dLbls", "cat", "val", "smooth", "explosion", "extLst"], C);
    const nativeIndex = integer(child(s, "idx", C) ?? s, "val");
    if (nativeIndex < 0 || item.order < 0 || indices.has(nativeIndex) || orders.has(item.order)) fail(s, "Native chart series identities or order are invalid.");
    indices.add(nativeIndex); orders.add(item.order);
    if (flag(s, "smooth") || (decimal(s, "explosion") ?? 0) !== 0) unsupported(s, "Smooth lines and exploded pie slices are not supported.");
    const categories = child(s, "cat", C), values = child(s, "val", C), title = child(s, "tx", C);
    if (!categories || !values) fail(s, "A chart series category/value vector is missing.");
    const numericCategories = !!child(categories, "numRef", C) || !!child(categories, "numLit", C);
    const categoryValues = cachedVector(context, categories, numericCategories, book).map(v => String(v ?? ""));
    const numbers = cachedVector(context, values, true, book) as (number | null)[];
    if (!index) data.categories = categoryValues;
    else if (JSON.stringify(data.categories) !== JSON.stringify(categoryValues)) unsupported(s, "Basic series must share the same category vector.");
    if (numbers.length !== data.categories.length) fail(s, "Chart category/value lengths disagree.");
    let name = "系列 " + (index + 1);
    if (title) {
      if (child(title, "v", C)) name = text(child(title, "v", C));
      else { const names = cachedVector(context, title, false, book); if (names.length !== 1) fail(title, "A chart series name requires one constant."); name = String(names[0] ?? ""); }
    }
    const seriesId = "series-" + index;
    data.series.push({ seriesId, name, values: numbers });
    const marker = child(s, "marker", C), symbol = val(marker, "symbol", kind === "line" ? "none" : "none")!;
    if (!["none", "circle", "square", "diamond", "triangle", "x", "star", "plus", "dot", "dash"].includes(symbol)) unsupported(marker ?? s, "This chart marker is not supported.");
    const markerSize = decimal(marker ?? s, "size") ?? 5;
    const seriesCss = style(context, s, "series-style", { "--office-invert-negative": String(flag(s, "invertIfNegative")) });
    const seriesLabel = child(s, "dLbls", C), labelCssClass = labelStyle(context, seriesLabel);
    definition.seriesStyles.push({ seriesId, cssClass: seriesCss, marker: symbol as NonNullable<PresentationChart["seriesStyles"]>[number]["marker"] & string, markerSize, ...(labelCssClass ? { labelCssClass } : {}) });
    for (const point of elements(s, "dPt", C)) {
      check(point, ["idx", "invertIfNegative", "marker", "bubble3D", "explosion", "spPr", "extLst"], C);
      if ((decimal(point, "explosion") ?? 0) !== 0 || flag(point, "bubble3D") || child(point, "marker", C)) unsupported(point, "This point override is outside the basic chart scope.");
      const pointIndex = integer(child(point, "idx", C) ?? point, "val");
      definition.pointStyles.push({ seriesId, index: pointIndex, cssClass: style(context, point, "point-style") });
    }
    const label = labels(seriesLabel, globalLabelContents);
    for (const pointLabel of elements(seriesLabel ?? globalLabels, "dLbl", C)) {
      const pointIndex = integer(child(pointLabel, "idx", C) ?? pointLabel, "val"), cssClass = labelStyle(context, pointLabel);
      let point = definition.pointStyles.find(p => p.seriesId === seriesId && p.index === pointIndex);
      if (!point) { const empty = context.id("point-style"); context.rule(empty, {}); point = { seriesId, index: pointIndex, cssClass: empty }; definition.pointStyles.push(point); }
      if (cssClass) point.labelCssClass = cssClass;
    }
    const serialized = JSON.stringify(label);
    if (commonLabels !== undefined && commonLabels !== serialized) unsupported(s, "Different label contents per series are not supported.");
    commonLabels = serialized; definition.labelContents = label;
  }
  definition.showLabels = !!definition.labelContents?.length;
  const rangeKey = (workbookPath ?? path) + "\n" + series.map(s => {
    const vectors = [child(s, "tx", C), child(s, "cat", C), child(s, "val", C)];
    return vectors.map(v => elements(v, undefined, C).map(r => text(child(r, "f", C))).join("|")).join(";");
  }).join("\n");
  const shared = context.dataSetsByRange.get(rangeKey);
  if (shared) {
    const previous = context.metadata.dataSets.find(d => d.dataSetId === shared)!;
    if (JSON.stringify({ ...data, dataSetId: "" }) !== JSON.stringify({ ...previous, dataSetId: "" })) fail(native, "Shared chart data references contain conflicting caches.");
    definition.dataSetId = shared;
  } else {
    context.bytes(data.categories.reduce((n, v) => n + v.length * 6 + 4, 0) + data.series.reduce((n, s) => n + s.name.length * 6 + s.values.length * 32 + 128, 0));
    context.metadata.dataSets.push(data); context.dataSetsByRange.set(rangeKey, data.dataSetId); definition.dataSetId = data.dataSetId;
  }
  const title = child(chart, "title", C);
  if (title) {
    definition.titleVisible = !flag(chart, "autoTitleDeleted");
    noLayout(title);
    if (flag(title, "overlay")) unsupported(title, "Overlay chart titles are not supported.");
    const tx = child(title, "tx", C), rich = child(tx, "rich", C), target = context.make("div");
    if (rich) importText(context, rich, target, path);
    else if (tx) { const contents = cachedVector(context, tx, false, book); if (contents.length !== 1) fail(tx, "A chart title must have one constant value."); const p = context.make("p"); context.content(p, String(contents[0] ?? "")); target.children.push(p); }
    else unsupported(title, "Generated chart titles without a source cache are not supported.");
    frame.children.push(target); definition.titleNodeId = target.attrs["data-node-id"]!;
    const titleCss = style(context, title, "chart-title-style"); target.attrs.class = [target.attrs.class, titleCss].filter(Boolean).join(" ");
  }
  const legend = child(chart, "legend", C);
  if (legend) {
    check(legend, ["legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr", "extLst"], C); noLayout(legend);
    if (elements(legend, "legendEntry", C).length || flag(legend, "overlay")) unsupported(legend, "Legend overrides and overlay legends are not supported.");
    const positions: Record<string, PresentationChart["legend"]> = { l: "left", r: "right", t: "top", b: "bottom" }, position = val(legend, "legendPos", "r")!;
    if (!positions[position]) unsupported(legend, "This legend position is not supported.");
    definition.legend = positions[position]!; definition.legendCssClass = style(context, legend, "chart-legend-style");
  }
  for (const [name, key] of [["gapWidth", "gapWidth"], ["overlap", "overlap"], ["firstSliceAng", "firstSliceAngle"]] as const) { const n = decimal(native, name); if (n !== undefined) definition[key] = n; }
  const axes = elements(plot, undefined, C).filter(n => n.name === "catAx" || n.name === "valAx");
  if (kind !== "pie" && (axes.length !== 2 || !axes.some(n => n.name === "catAx") || !axes.some(n => n.name === "valAx"))) unsupported(plot, "A basic chart requires one category axis and one value axis.");
  const axisIds = new Set<string>();
  for (const axis of axes) {
    check(axis, ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "crossBetween", "majorUnit", "minorUnit", "extLst"], C);
    const axisId = val(axis, "axId"); if (!axisId || axisIds.has(axisId)) fail(axis, "Chart axis identities are missing or duplicated."); axisIds.add(axisId);
    if (child(axis, "title", C) || child(axis, "minorGridlines", C) || child(axis, "crossesAt", C)) unsupported(axis, "Axis titles, minor grid lines and custom crossing points are not supported.");
    const scaling = child(axis, "scaling", C);
    check(scaling, ["orientation", "max", "min", "logBase", "extLst"], C);
    if (child(scaling, "logBase", C)) unsupported(axis, "Logarithmic axes are not supported.");
    const reverse = val(scaling, "orientation", "minMax") === "maxMin";
    if (axis.name === "catAx") { definition.categoryAxisCssClass = axisStyle(context, axis, "chart-category-axis"); definition.showCategoryAxis = !flag(axis, "delete"); definition.reverseCategories = reverse; }
    else {
      definition.valueAxisCssClass = axisStyle(context, axis, "chart-value-axis"); definition.showValueAxis = !flag(axis, "delete"); definition.reverseValues = reverse;
      for (const [native, key] of [["min", "minimum"], ["max", "maximum"]] as const) { const n = scaling && decimal(scaling, native); if (n !== undefined) definition[key] = n; }
      const unit = decimal(axis, "majorUnit"); if (unit !== undefined) definition.majorUnit = unit;
      const format = child(axis, "numFmt", C); if (format && value(format, "formatCode") !== undefined) definition.numberFormat = value(format, "formatCode")!;
      const grid = child(axis, "majorGridlines", C); if (grid) definition.gridCssClass = style(context, grid, "chart-grid-style");
    }
  }
  for (const axis of axes) if (!axisIds.has(val(axis, "crossAx") ?? "")) fail(axis, "A chart axis crossing reference is invalid.");
  const blankAs = val(chart, "dispBlanksAs", "gap");
  if (!blankAs || !["gap", "zero", "span"].includes(blankAs)) unsupported(chart, "This blank-value behavior is not supported.");
  definition.blankAs = blankAs as "gap" | "zero" | "span";
  definition.plotCssClass = style(context, plot, "chart-plot-style"); definition.textCssClass = context.id("chart-text-style");
  context.rule(definition.textCssClass, textStyle(child(source, "txPr", C)));
  const frameStyle = style(context, source, "chart-frame-style", { "--office-chart-vary-colors": String(flag(native, "varyColors")) });
  frame.attrs.class = [frame.attrs.class, frameStyle].filter(Boolean).join(" ");
  context.metadata.charts.push(definition); return id;
}
