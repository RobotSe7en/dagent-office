import type { ChartAppearance } from "@doffi/core";
import { A, C, check, child, elements, fail, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { bool, value } from "./values.js";
import { effects, fill, outline } from "./styles.js";
import { paragraphProperties, runProperties } from "./text.js";
import { chartAxisCrossing } from "./chart-axis.js";

export interface ChartStyleContext {
  id(prefix: string): string;
  rule(name: string, css: Css): void;
  textStyle?(source: XmlNode | undefined): Css;
}

export function val(node: XmlNode | undefined, name: string, fallback?: string): string | undefined { return value(child(node, name, C)) ?? fallback; }
export function flag(node: XmlNode | undefined, name: string, fallback = false): boolean { const source = child(node, name, C); return source ? bool(source, "val", true) : fallback; }
export function decimal(node: XmlNode, name: string): number | undefined {
  const property = child(node, name, C);
  if (!property) return undefined;
  const source = value(property), n = Number(source);
  if (!source?.trim() || !Number.isFinite(n)) fail(property, "A chart numeric option is invalid.");
  return n;
}
export function appearance(node: XmlNode | undefined): Css {
  if (!node) return {};
  check(node, ["noFill", "solidFill", "gradFill", "ln", "effectLst", "extLst"], A);
  return { ...fill(node), ...outline(child(node, "ln", A)), ...effects(child(node, "effectLst", A)) };
}
export function textStyle(node: XmlNode | undefined): Css {
  if (!node) return {};
  const p = child(node, "p", A);
  return { ...paragraphProperties(child(p, "pPr", A)), ...runProperties(child(p, "endParaRPr", A)) };
}
export function style(context: ChartStyleContext, source: XmlNode | undefined, prefix: string, extra: Css = {}): string {
  const name = context.id(prefix);
  context.rule(name, { ...appearance(child(source, "spPr", C)), ...(context.textStyle ?? textStyle)(child(source, "txPr", C)), ...extra }); return name;
}
export function noLayout(node: XmlNode | undefined): void {
  const layout = child(node, "layout", C);
  if (layout && elements(layout, undefined, C).length) unsupported(layout, "Manual chart layout is outside the basic chart scope.");
}
export function labels(node: XmlNode | undefined, defaults: NonNullable<ChartAppearance["labelContents"]> = []): NonNullable<ChartAppearance["labelContents"]> {
  if (!node) return defaults;
  check(node, ["dLbl", "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines", "leaderLines", "extLst"], C);
  if (flag(node, "delete")) return [];
  if (flag(node, "showLegendKey") || flag(node, "showBubbleSize")) unsupported(node, "Legend keys and bubble labels are not supported.");
  const result: NonNullable<ChartAppearance["labelContents"]> = [];
  for (const [name, kind] of [["showVal", "value"], ["showCatName", "category"], ["showSerName", "series"], ["showPercent", "percent"]] as const) if (flag(node, name, defaults.includes(kind))) result.push(kind);
  for (const point of elements(node, "dLbl", C)) {
    const selected = labels({ ...point, content: point.content.filter(n => typeof n === "string" || n.uri !== C || n.name !== "idx") }, result);
    if (JSON.stringify(selected) !== JSON.stringify(result)) unsupported(point, "Per-point label visibility is outside the basic chart scope.");
  }
  return result;
}
export function cssString(value: string): string { return '"' + value.replace(/["\\\u0000-\u001f\u007f]/g, c => c === '"' || c === "\\" ? "\\" + c : "\\" + c.charCodeAt(0).toString(16) + " ") + '"'; }
export function labelStyle(context: ChartStyleContext, source: XmlNode | undefined): string | undefined {
  if (!source) return undefined;
  const css: Css = {}, format = child(source, "numFmt", C), separator = child(source, "separator", C), position = val(source, "dLblPos");
  if (format && value(format, "formatCode") !== undefined) css["--office-label-number-format"] = cssString(value(format, "formatCode")!);
  if (separator) css["--office-label-separator"] = cssString(text(separator));
  if (position) css["--office-label-position"] = position;
  return style(context, source, "chart-label-style", css);
}
export function axisStyle(context: ChartStyleContext, source: XmlNode, prefix: string): string {
  const css: Css = chartAxisCrossing(source);
  for (const [name, key, fallback] of [["axPos", "position", source.name === "catAx" ? "b" : "l"], ["majorTickMark", "major-tick", "cross"], ["minorTickMark", "minor-tick", "cross"], ["tickLblPos", "label-position", "nextTo"], ["lblAlgn", "label-align", "ctr"], ["lblOffset", "label-offset", "100"], ["tickLblSkip", "label-skip", "1"], ["tickMarkSkip", "tick-skip", "1"], ["crossBetween", "cross-between", "between"], ["minorUnit", "minor-unit", "auto"]]) if (child(source, name!, C)) css["--office-axis-" + key] = val(source, name!, fallback!)!;
  return style(context, source, prefix, css);
}
