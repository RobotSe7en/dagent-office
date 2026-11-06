import { A, C, check, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, finite, integer, one, spreadsheetText, value } from "./values.js";
import { validateDrawingEffects, validateDrawingFill, validateDrawingLine } from "./drawing-style.js";
import { drawingTextStyle } from "./drawing-text.js";
import { chartMarkerFlag } from "../drawingml/chart-marker.js";

const children: Readonly<Record<string, readonly string[]>> = {
  chartSpace: ["date1904", "lang", "roundedCorners", "chart", "spPr", "txPr"],
  chart: ["title", "autoTitleDeleted", "plotArea", "legend", "plotVisOnly", "dispBlanksAs", "showDLblsOverMax"],
  title: ["tx", "layout", "overlay", "spPr", "txPr"], plotArea: ["layout", "barChart", "lineChart", "pieChart", "catAx", "valAx", "spPr"],
  barChart: ["barDir", "grouping", "varyColors", "ser", "dLbls", "gapWidth", "overlap", "axId"],
  lineChart: ["grouping", "varyColors", "ser", "dLbls", "hiLowLines", "marker", "smooth", "axId"],
  pieChart: ["varyColors", "ser", "dLbls", "firstSliceAng"],
  ser: ["idx", "order", "tx", "spPr", "invertIfNegative", "marker", "dPt", "dLbls", "cat", "val", "smooth", "explosion"],
  tx: ["strRef", "rich", "v"], cat: ["strRef", "numRef", "strLit", "numLit"], val: ["numRef", "numLit"],
  marker: ["symbol", "size", "spPr"], dPt: ["idx", "invertIfNegative", "marker", "bubble3D", "explosion", "spPr"],
  dLbls: ["dLbl", "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines", "leaderLines", "extLst"],
  dLbl: ["idx", "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator"],
  leaderLines: ["spPr"], hiLowLines: ["spPr"], legend: ["legendPos", "layout", "overlay", "spPr", "txPr"],
  catAx: ["axId", "scaling", "delete", "axPos", "majorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl"],
  valAx: ["axId", "scaling", "delete", "axPos", "majorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossBetween", "majorUnit", "minorUnit"],
  scaling: ["orientation", "max", "min"], majorGridlines: ["spPr"], layout: [],
};
const booleans = new Set(["date1904", "roundedCorners", "autoTitleDeleted", "plotVisOnly", "showDLblsOverMax", "overlay", "varyColors", "invertIfNegative", "smooth", "bubble3D", "delete", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "showLeaderLines", "auto", "noMultiLvlLbl"]);
const enums: Readonly<Record<string, readonly string[]>> = {
  barDir: ["bar", "col"], grouping: ["percentStacked", "standard", "stacked", "clustered"],
  symbol: ["circle", "dash", "diamond", "dot", "none", "picture", "plus", "square", "star", "triangle", "x", "auto"],
  dLblPos: ["bestFit", "b", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t"], legendPos: ["b", "tr", "l", "r", "t"],
  orientation: ["minMax", "maxMin"], axPos: ["b", "l", "r", "t"], majorTickMark: ["cross", "in", "none", "out"], minorTickMark: ["cross", "in", "none", "out"],
  tickLblPos: ["high", "low", "nextTo", "none"], crosses: ["autoZero", "max", "min"], lblAlgn: ["ctr", "l", "r"], crossBetween: ["between", "midCat"], dispBlanksAs: ["span", "gap", "zero"],
};
const numeric = new Set(["idx", "order", "axId", "crossAx", "gapWidth", "overlap", "firstSliceAng", "explosion", "size", "lblOffset", "tickLblSkip", "tickMarkSkip", "min", "max", "majorUnit", "minorUnit"]);
const enumDefaults: Readonly<Record<string, string>> = { barDir: "col", legendPos: "r", orientation: "minMax", majorTickMark: "cross", minorTickMark: "cross", tickLblPos: "nextTo", dispBlanksAs: "zero" };
const integerOptions: Readonly<Record<string, readonly [number, number, number?]>> = {
  gapWidth: [0, 500, 150], overlap: [-100, 100, 0], firstSliceAng: [0, 360, 0], size: [2, 72, 5], lblOffset: [0, 1000, 100],
  tickLblSkip: [1, 0xffff_ffff], tickMarkSkip: [1, 0xffff_ffff], explosion: [0, 0xffff_ffff],
};
const withValue = (node: XmlNode, val: string): XmlNode => ({ ...node, attrs: new Map([...node.attrs, [" val", val]]) });
const C15 = "http://schemas.microsoft.com/office/drawing/2012/chart";
/** MS-ODRAWXML 2.2.1.2 / 2.3.60 extends labels on non-pie charts with
 * the same leader-line visibility flag. No other extension is discarded. */
export function chartLabelLeaderLines(source: XmlNode): boolean | undefined {
  const legacy = one(source, "showLeaderLines", C), extensions = one(source, "extLst", C);
  let result = legacy ? boolean(legacy, "val", true) : undefined;
  if (!extensions) return result;
  attributes(extensions, []); check(extensions, ["ext"], C);
  const extension = one(extensions, "ext", C);
  if (!extension) return result;
  attributes(extension, ["uri"]);
  if (value(extension, "uri") !== "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}") unsupported(extension, "This chart-label extension is not supported.");
  check(extension, ["showLeaderLines"], C15);
  const flag = one(extension, "showLeaderLines", C15);
  if (!flag) return result;
  attributes(flag, ["val"]); check(flag, [], C15);
  const selected = boolean(flag, "val", true);
  if (result !== undefined && result !== selected) fail(source, "The native and extended leader-line settings conflict.");
  result = selected; return result;
}
/** Audit every selected chart node before the shared appearance adapter reads
 * it, including attributes on otherwise innocuous leaves and inactive styles. */
export function validateWorksheetChart(source: XmlNode): XmlNode {
  const visit = (node: XmlNode, parent?: XmlNode): XmlNode => {
    if (node.uri !== C) unsupported(node);
    if (["strRef", "numRef", "strLit", "numLit", "rich"].includes(node.name)) return node; // Consumed by vector/text readers, including cumulative budgets.
    if (node.name === "txPr") { drawingTextStyle(node); return node; }
    if (node.name === "spPr") {
      attributes(node, ["bwMode"]); check(node, ["noFill", "solidFill", "gradFill", "ln", "effectLst"], A);
      if (value(node, "bwMode") && value(node, "bwMode") !== "auto") unsupported(node, "Explicit chart black-and-white modes are not supported.");
      let fills = 0;
      for (const item of elements(node, undefined, A)) {
        one(node, item.name, A);
        if (item.name === "ln") validateDrawingLine(item);
        else if (item.name === "effectLst") validateDrawingEffects(item);
        else { fills++; validateDrawingFill(item); }
      }
      if (fills > 1) fail(node, "A chart appearance has multiple fills."); return node;
    }
    if (node.name === "v" || node.name === "separator") {
      attributes(node, []); if (node.content.some(item => typeof item !== "string")) fail(node, "A chart text value has nested markup."); return node;
    }
    if (node.name === "numFmt") {
      attributes(node, ["formatCode", "sourceLinked"]); check(node, [], C);
      if (value(node, "formatCode") === undefined) fail(node, "A chart number format is missing its code.");
      spreadsheetText(value(node, "formatCode")!, node, 2048); boolean(node, "sourceLinked"); return node;
    }
    if (node.name === "lang") { attributes(node, ["val"]); check(node, [], C); const language = value(node, "val"); if (!language || !/^[A-Za-z]{1,8}(?:-[A-Za-z0-9]{1,8})*$/.test(language)) fail(node, "A chart editing language is invalid."); return node; }
    if (node.name === "marker" && parent?.name === "lineChart") return withValue(node, chartMarkerFlag(node) ? "1" : "0");
    if (booleans.has(node.name) || numeric.has(node.name) || enums[node.name]) {
      attributes(node, ["val"]); check(node, [], C);
      if (booleans.has(node.name)) boolean(node, "val", true);
      else if (enums[node.name]) {
        const selected = value(node, "val") ?? (node.name === "grouping" ? parent?.name === "barChart" ? "clustered" : "standard" : enumDefaults[node.name]);
        if (!selected || !enums[node.name]!.includes(selected)) fail(node, "A chart option has an invalid enum value.");
        return withValue(node, selected);
      }
      else if (["idx", "order", "axId", "crossAx"].includes(node.name)) integer(node, "val");
      else if (integerOptions[node.name]) {
        const [min, max, fallback] = integerOptions[node.name]!, raw = value(node, "val");
        // Transitional accepts both integral amounts and percentage lexemes.
        const normalized = raw?.endsWith("%") && ["gapWidth", "overlap", "lblOffset"].includes(node.name) ? withValue(node, raw.slice(0, -1)) : node;
        return withValue(node, String(integer(normalized, "val", min, max, fallback)));
      } else {
        const raw = value(node, "val"); if (raw === undefined) fail(node, "A chart numeric option is missing its value.");
        const amount = finite(raw, node); if (["majorUnit", "minorUnit"].includes(node.name) && amount <= 0) fail(node, "Chart axis units must be positive.");
      }
      return node;
    }
    const allowed = children[node.name]; if (!allowed) unsupported(node, "This native chart feature is outside the selected chart reader.");
    attributes(node, []); check(node, allowed, C);
    if (node.name === "dLbls") chartLabelLeaderLines(node);
    if (node.name === "ser") {
      const selected = parent?.name === "barChart" ? ["invertIfNegative"] : parent?.name === "lineChart" ? ["marker", "smooth"] : ["explosion"];
      for (const item of elements(node, undefined, C)) if (["invertIfNegative", "marker", "smooth", "explosion"].includes(item.name) && !selected.includes(item.name)) fail(item, "A series option does not belong to its chart type.");
    }
    return { ...node, content: node.content.map(item => {
      if (typeof item === "string") return item;
      if (!["ser", "dPt", "dLbl", "axId"].includes(item.name)) one(node, item.name, C);
      if (node.name === "dLbls" && item.name === "extLst") return item;
      return visit(item, node);
    }) };
  };
  return visit(source);
}
