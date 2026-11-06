import type { WorkbookChart, WorkbookChartVector } from "@dagent-office/core";
import { parseWorkbookFormula, printWorkbookFormula } from "@dagent-office/core";
import type { OpcPackage } from "../ooxml/package.js";
import type { Css } from "../ooxml/styles.js";
import { A, attr, C, check, elements, fail, R, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Html } from "../drawingml/html.js";
import { appearance, axisStyle, cssString, decimal, flag, labels, labelStyle, noLayout, style, val } from "../drawingml/chart-style.js";
import { attributes, boolean, integer, one, spreadsheetText, value } from "./values.js";
import { readChartLabel, readChartReference, readChartVector } from "./chart-data.js";
import type { ChartDataBudget } from "./chart-data.js";
import { drawingTextStyle, readDrawingText } from "./drawing-text.js";
import { chartMarker, chartMarkerFlag } from "../drawingml/chart-marker.js";
import { chartLabelLeaderLines, validateWorksheetChart } from "./chart-validation.js";
import type { XlsxStyles } from "./styles.js";
import type { WorkbookContent } from "./content.js";

export class WorkbookCharts {
  readonly definitions: WorkbookChart[] = [];
  readonly parts: Set<string> = new Set();
  private readonly budget: ChartDataBudget;
  constructor(readonly opc: OpcPackage, readonly styles: XlsxStyles, readonly content: WorkbookContent) {
    this.budget = { items: 0, bytes: 0, maxItems: content.limits.maxNodes, maxBytes: content.limits.maxFileBytes };
  }
  id(_prefix: string): string { return this.content.id(); }
  rule(name: string, css: Css): void { this.styles.drawingRule(name, css); }
  textStyle(source: XmlNode | undefined): Css { return drawingTextStyle(source); }
  private labelStyle(source: XmlNode | undefined): string | undefined {
    if (!source) return undefined;
    const format = one(source, "numFmt", C);
    const normalized = format ? { ...source, content: source.content.map(node => node === format ? { ...format, attrs: new Map([...format.attrs, [" formatCode", spreadsheetText(value(format, "formatCode")!, format, 2048)]]) } : node) } : source;
    const name = labelStyle(this, normalized)!;
    const css: Css = {};
    // A present numFmt defaults sourceLinked to true. Emit that default on
    // every overriding label style so a point cannot inherit a parent's false.
    if (format) css["--office-label-number-format-source-linked"] = String(boolean(format, "sourceLinked", true));
    const leader = one(source, "leaderLines", C);
    const showLeaderLines = chartLabelLeaderLines(source);
    if (showLeaderLines !== undefined) css["--office-label-show-leader-lines"] = String(showLeaderLines);
    for (const [key, value] of Object.entries(appearance(one(leader, "spPr", C)))) css[key.replace("--office-", "--office-label-leader-")] = value;
    if (Object.keys(css).length) this.rule(name, css); return name;
  }
  read(reference: XmlNode, frame: Html, hyperlink: (source: XmlNode | undefined, target: Html, part: string) => void): string {
    attributes(reference, ["id"], R); check(reference, [], C);
    const id = attr(reference, "id", R), relation = id ? this.opc.relationships(reference.part).get(id) : undefined;
    if (!relation || relation.external || relation.type !== R + "/chart" || this.parts.has(relation.target)) fail(reference, "A chart requires its own valid native part relationship.");
    const path = relation.target;
    if (this.opc.contentType(path) !== "application/vnd.openxmlformats-officedocument.drawingml.chart+xml") fail(reference, "A chart part has an invalid content type.");
    const root = this.opc.read(path); this.parts.add(path);
    if (root.name !== "chartSpace" || root.uri !== C) fail(root, "A chart part has an invalid root.");
    const source = validateWorksheetChart(root);
    if (flag(source, "roundedCorners")) unsupported(source, "Rounded chart frames are not supported.");
    const chart = one(source, "chart", C), plot = one(chart, "plotArea", C);
    if (!chart || !plot) fail(source, "A chart has no plot area."); noLayout(plot);
    const plots = elements(plot, undefined, C).filter(node => ["barChart", "lineChart", "pieChart"].includes(node.name));
    if (plots.length !== 1) unsupported(plot, "Only one basic chart type may occupy a plot.");
    const native = plots[0]!, kind = native.name === "barChart" ? "column" : native.name === "lineChart" ? "line" : "pie";
    if (kind === "column" && (val(native, "barDir", "col") !== "col" || val(native, "grouping", "clustered") !== "clustered") || kind === "line" && val(native, "grouping", "standard") !== "standard" || flag(native, "smooth")) unsupported(native, "Only clustered columns, plain lines and single-series pies are supported.");
    const highLow = one(native, "hiLowLines", C);
    if (highLow && !one(one(one(highLow, "spPr", C), "ln", A), "noFill", A)) unsupported(highLow, "High/low chart lines are not supported.");
    const definition: WorkbookChart = { chartId: this.id("chart"), nodeId: frame.attrs["data-node-id"]!, kind, legend: "none", showLabels: false, series: [], seriesStyles: [], pointStyles: [] };
    const markerFlag = kind === "line" ? chartMarkerFlag(one(native, "marker", C)) : undefined;
    if (markerFlag !== undefined) definition.nativeShowMarkers = markerFlag;
    // An omitted element is distinct from a present CT_Boolean with omitted
    // val (true). Preserve that distinction instead of inventing source flags.
    if (one(chart, "plotVisOnly", C)) definition.visibleCellsOnly = flag(chart, "plotVisOnly");
    if (one(chart, "showDLblsOverMax", C)) definition.showLabelsOverMaximum = flag(chart, "showDLblsOverMax");
    if (one(native, "varyColors", C)) definition.varyColors = flag(native, "varyColors");
    const language = val(source, "lang"); if (language) definition.language = language;
    if (one(source, "date1904", C)) definition.dateSystem = flag(source, "date1904") ? "1904" : "1900";
    const series = elements(native, "ser", C);
    if (!series.length || kind === "pie" && series.length !== 1) unsupported(native, "This basic chart series arrangement is not supported.");
    const ordered = series.map(source => { const order = one(source, "order", C); if (!order) fail(source, "A chart series has no order."); return { source, order: integer(order, "val") }; }).sort((a, b) => a.order - b.order);
    const indices = new Set<number>(), orders = new Set<number>(), globalLabels = one(native, "dLbls", C), globalContents = labels(globalLabels), globalStyle = this.labelStyle(globalLabels);
    if (globalStyle) definition.labelCssClass = globalStyle;
    let categoryKey: string | undefined, labelKey: string | undefined;
    const vectorKey = (source: WorkbookChartVector): string => "expression" in source
      ? source.valueType + ":" + printWorkbookFormula(parseWorkbookFormula(source.expression, { row: 0, column: 0 }), { row: 0, column: 0 })
      : JSON.stringify(source);
    for (const [position, entry] of ordered.entries()) {
      const s = entry.source, indexNode = one(s, "idx", C); if (!indexNode) fail(s, "A chart series has no native identity.");
      const index = integer(indexNode, "val");
      if (indices.has(index) || orders.has(entry.order)) fail(s, "Native series identities or order are duplicated."); indices.add(index); orders.add(entry.order);
      if (flag(s, "smooth") || (decimal(s, "explosion") ?? 0) !== 0) unsupported(s, "Smoothed lines and exploded slices are not supported.");
      const categories = one(s, "cat", C), values = one(s, "val", C), name = one(s, "tx", C);
      if (!values) fail(s, "A chart series has no value source.");
      const category = categories ? readChartVector(categories, this.budget) : undefined, key = category ? vectorKey(category) : "automatic";
      if (!position) { categoryKey = key; if (category) definition.categories = category; }
      else if (categoryKey !== key) unsupported(s, "Basic chart series must share one category source.");
      const seriesId = `series-${position}`;
      definition.series.push({ seriesId, nativeIndex: index, values: readChartVector(values, this.budget, true), ...(name ? { name: readChartLabel(name, this.budget) } : {}) });
      const markerSource = one(s, "marker", C), { css: markerCss, ...marker } = chartMarker(markerSource);
      const seriesCss = style(this, s, "series-style", { ...markerCss, "--office-invert-negative": String(flag(s, "invertIfNegative")) });
      const seriesLabels = one(s, "dLbls", C), seriesLabelStyle = this.labelStyle(seriesLabels);
      definition.seriesStyles!.push({ seriesId, cssClass: seriesCss, marker: marker.marker ?? "none", markerSize: marker.markerSize ?? 5, ...(seriesLabelStyle ? { labelCssClass: seriesLabelStyle } : {}) });
      const points = new Set<number>();
      for (const point of elements(s, "dPt", C)) {
        const idx = one(point, "idx", C); if (!idx) fail(point, "A chart point override is missing its index."); const index = integer(idx, "val");
        if (points.has(index)) fail(point, "A chart point override is repeated."); points.add(index);
        if ((decimal(point, "explosion") ?? 0) !== 0 || flag(point, "bubble3D")) unsupported(point, "This point override is outside the basic chart scope.");
        const { css, ...marker } = chartMarker(one(point, "marker", C));
        definition.pointStyles!.push({ seriesId, index, ...marker, cssClass: style(this, point, "point-style", { ...css, "--office-invert-negative": String(flag(point, "invertIfNegative", flag(s, "invertIfNegative"))) }) });
      }
      const contents = labels(seriesLabels, globalContents), selected = JSON.stringify(contents);
      if (labelKey !== undefined && labelKey !== selected) unsupported(s, "Different label contents per series are not supported."); labelKey = selected; definition.labelContents = contents;
      const labelled = new Set<number>();
      for (const label of elements(seriesLabels ?? globalLabels, "dLbl", C)) {
        const idx = one(label, "idx", C); if (!idx) fail(label, "A point label has no index."); const index = integer(idx, "val");
        if (labelled.has(index)) fail(label, "A point label is repeated."); labelled.add(index);
        let point = definition.pointStyles!.find(point => point.seriesId === seriesId && point.index === index);
        if (!point) { point = { seriesId, index, cssClass: style(this, undefined, "point-style") }; definition.pointStyles!.push(point); }
        const cssClass = this.labelStyle(label); if (cssClass) point.labelCssClass = cssClass;
      }
    }
    definition.showLabels = !!definition.labelContents?.length;
    const readTitle = (title: XmlNode): NonNullable<WorkbookChart["categoryAxisTitle"]> => {
      noLayout(title); if (flag(title, "overlay")) unsupported(title, "Overlay titles are not supported.");
      const tx = one(title, "tx", C), rich = one(tx, "rich", C); if (!tx) unsupported(title, "Generated chart titles have no editable source.");
      check(tx, ["rich", "strRef"], C); if (elements(tx, undefined, C).length !== 1) fail(tx, "A chart title must have exactly one source.");
      const cssClass = style(this, title, "chart-title-style");
      if (rich) {
        const body = this.content.make("div"), text = readDrawingText(rich, body, this.content, this.styles, hyperlink);
        body.attrs.class = [body.attrs.class, cssClass].filter(Boolean).join(" "); frame.children.push(body);
        return { cssClass, text, nodeId: body.attrs["data-node-id"]! };
      }
      attributes(tx, []);
      return { cssClass, reference: readChartReference(one(tx, "strRef", C)!, this.budget) };
    };
    const title = one(chart, "title", C);
    if (title) {
      const mapped = readTitle(title); definition.titleVisible = !flag(chart, "autoTitleDeleted");
      definition.titleCssClass = mapped.cssClass;
      if (mapped.nodeId) { definition.titleNodeId = mapped.nodeId; definition.titleText = mapped.text; }
      if (mapped.reference) definition.titleReference = mapped.reference;
    }
    const legend = one(chart, "legend", C);
    if (legend) {
      noLayout(legend); if (flag(legend, "overlay")) unsupported(legend, "Overlay legends are not supported.");
      const positions: Readonly<Record<string, WorkbookChart["legend"]>> = { l: "left", r: "right", t: "top", b: "bottom" }, position = val(legend, "legendPos", "r")!;
      if (!positions[position]) unsupported(legend, "This legend position is not supported."); definition.legend = positions[position]!; definition.legendCssClass = style(this, legend, "chart-legend-style");
    }
    for (const [nativeName, key] of [["gapWidth", "gapWidth"], ["overlap", "overlap"], ["firstSliceAng", "firstSliceAngle"]] as const) { const amount = decimal(native, nativeName); if (amount !== undefined) definition[key] = amount; }
    const axes = elements(plot, undefined, C).filter(node => ["catAx", "valAx"].includes(node.name)), axisIds = new Set<number>();
    if (kind === "pie" ? axes.length !== 0 : axes.length !== 2 || !axes.some(node => node.name === "catAx") || !axes.some(node => node.name === "valAx")) unsupported(plot, "A basic chart requires exactly its category and value axes.");
    for (const axis of axes) {
      const identity = one(axis, "axId", C), scaling = one(axis, "scaling", C); if (!identity || !scaling) fail(axis, "A chart axis identity or scale is missing.");
      const id = integer(identity, "val"); if (axisIds.has(id)) fail(axis, "A chart axis identity is duplicated."); axisIds.add(id);
      const reverse = val(scaling, "orientation", "minMax") === "maxMin", format = one(axis, "numFmt", C);
      const cssClass = axisStyle(this, axis, axis.name === "catAx" ? "category-axis" : "value-axis");
      const title = one(axis, "title", C); if (title) definition[axis.name === "catAx" ? "categoryAxisTitle" : "valueAxisTitle"] = readTitle(title);
      if (axis.name === "catAx") {
        if (one(scaling, "min", C) || one(scaling, "max", C)) unsupported(scaling, "Explicit category-axis scale bounds are not supported.");
        definition.categoryAxisCssClass = cssClass; definition.showCategoryAxis = !flag(axis, "delete"); definition.reverseCategories = reverse;
        const grid = one(axis, "majorGridlines", C); if (grid) definition.categoryGridCssClass = style(this, grid, "chart-category-grid-style");
      } else {
        definition.valueAxisCssClass = cssClass; definition.showValueAxis = !flag(axis, "delete"); definition.reverseValues = reverse;
        for (const [nativeName, key] of [["min", "minimum"], ["max", "maximum"]] as const) { const amount = decimal(scaling, nativeName); if (amount !== undefined) definition[key] = amount; }
        const unit = decimal(axis, "majorUnit"); if (unit !== undefined) definition.majorUnit = unit;
        if (format) definition.numberFormat = spreadsheetText(value(format, "formatCode")!, format, 2048);
        const grid = one(axis, "majorGridlines", C); if (grid) definition.gridCssClass = style(this, grid, "chart-grid-style");
      }
      if (format) {
        const code = spreadsheetText(value(format, "formatCode")!, format, 2048), quoted = cssString(code);
        this.rule(cssClass, { "--office-axis-number-format": quoted, "--office-axis-number-format-source-linked": String(boolean(format, "sourceLinked", true)) });
      }
      for (const name of ["auto", "noMultiLvlLbl"]) if (one(axis, name, C)) this.rule(cssClass, { [`--office-axis-${name}`]: String(flag(axis, name)) });
    }
    for (const axis of axes) {
      const cross = one(axis, "crossAx", C), identity = one(axis, "axId", C);
      if (!cross || !axisIds.has(integer(cross, "val")) || integer(cross, "val") === integer(identity!, "val")) fail(axis, "A chart axis crossing reference is invalid.");
    }
    const references = elements(native, "axId", C).map(node => integer(node, "val"));
    if (references.length !== axisIds.size || new Set(references).size !== references.length || references.some(id => !axisIds.has(id))) fail(native, "A chart plot does not reference its declared axes.");
    definition.blankAs = val(chart, "dispBlanksAs", "gap") as "gap" | "zero" | "span";
    definition.plotCssClass = style(this, plot, "chart-plot-style"); definition.textCssClass = this.id("chart-text-style"); this.rule(definition.textCssClass, this.textStyle(one(source, "txPr", C)));
    frame.attrs.class = [frame.attrs.class, style(this, source, "chart-frame-style")].filter(Boolean).join(" ");
    this.definitions.push(definition); return definition.chartId;
  }
}
