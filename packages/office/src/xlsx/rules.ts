import { parseWorkbookFormula, WorkbookComparisonSchema } from "@doffi/core";
import type { WorkbookColorThreshold, WorkbookConditionalRule, WorkbookCoordinate, WorkbookDataValidation, WorkbookRange, WorkbookSheet } from "@doffi/core";
import { check, elements, fail, S, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { XlsxStyles } from "./styles.js";
import { attributes, boolean, finite, integer, one, range, spreadsheetText, value } from "./values.js";

function regions(node: XmlNode): WorkbookRange[] {
  const source = value(node, "sqref")?.trim(); if (!source) fail(node, "A worksheet rule has no target ranges.");
  return source.split(/\s+/).map(item => range(item, node));
}
/** The pinned native producer uses the leftmost actual range start, then the
 * topmost at that column, independently of sqref order. Do not synthesize the
 * enclosing rectangle's corner; that can shift every relative rule reference. */
function origin(ranges: WorkbookRange[]): WorkbookCoordinate {
  return ranges.reduce((point, area) => area.from.column < point.column || area.from.column === point.column && area.from.row < point.row ? { ...area.from } : point, { ...ranges[0]!.from });
}
function formula(source: string, point: WorkbookCoordinate, node: XmlNode): string {
  const expression = "=" + source; if (!source) fail(node, "A rule formula is empty.");
  parseWorkbookFormula(expression, point); return expression;
}
function formulaText(node: XmlNode): string {
  attributes(node, []); if (node.content.some(item => typeof item !== "string")) fail(node, "A rule formula contains unexpected elements.");
  return text(node);
}
function comparison(node: XmlNode, fallback?: string): Extract<WorkbookConditionalRule, { kind: "cellIs" }>["operator"] {
  const result = WorkbookComparisonSchema.safeParse(value(node, "operator") ?? fallback);
  if (!result.success) fail(node, "A rule comparison operator is invalid."); return result.data;
}
function threshold(node: XmlNode, color: XmlNode, point: WorkbookCoordinate, styles: XlsxStyles): WorkbookColorThreshold {
  attributes(node, ["type", "val", "gte"]); check(node, [], S);
  // gte applies to icon sets; it has no effect in a color scale.
  if (value(node, "gte") !== undefined) boolean(node, "gte");
  const kind = value(node, "type"), raw = value(node, "val"), cssClass = styles.scaleColor(color);
  if (kind === "min" || kind === "max") return { kind, cssClass };
  if (raw === undefined) fail(node, "A color scale threshold has no value.");
  if (kind === "formula") return { kind, expression: formula(raw, point, node), cssClass };
  if (kind === "num" || kind === "percent" || kind === "percentile") {
    const amount = finite(raw, node);
    if (kind !== "num" && (amount < 0 || amount > 100)) fail(node, "A color scale percentage is out of bounds.");
    return { kind: kind === "num" ? "number" : kind, value: amount, cssClass };
  }
  return unsupported(node, "This color scale threshold is not supported.");
}
export function readConditionalRules(root: XmlNode, sheetId: string, styles: XlsxStyles): WorkbookConditionalRule[] {
  const result: WorkbookConditionalRule[] = [];
  for (const container of elements(root, "conditionalFormatting", S)) {
    attributes(container, ["sqref", "pivot"]); check(container, ["cfRule"], S);
    if (boolean(container, "pivot")) unsupported(container, "Pivot conditional formatting is not supported.");
    const ranges = regions(container), point = origin(ranges), rules = elements(container, "cfRule", S);
    if (!rules.length) fail(container, "A conditional formatting container has no rules.");
    for (const node of rules) {
      attributes(node, ["type", "dxfId", "priority", "stopIfTrue", "operator", "aboveAverage", "equalAverage", "bottom", "percent", "rank", "text", "stdDev", "timePeriod"]);
      check(node, ["formula", "colorScale"], S);
      for (const key of ["aboveAverage", "equalAverage", "bottom", "percent"]) if (value(node, key) !== undefined) boolean(node, key);
      if (value(node, "rank") !== undefined) integer(node, "rank");
      // These attributes are explicitly ignored by native semantics for the
      // three supported rule types; Calc writes several on every scalar rule.
      const kind = value(node, "type"), common = { ruleId: `${sheetId}-cf-${result.length + 1}`, ranges, origin: point, priority: integer(node, "priority", 1, 2147483647), stopIfTrue: boolean(node, "stopIfTrue") };
      const formulas = elements(node, "formula", S).map(item => formula(formulaText(item), point, item)), scale = one(node, "colorScale");
      if (kind === "colorScale") {
        if (!scale || formulas.length || value(node, "dxfId") !== undefined || common.stopIfTrue) fail(node, "A color scale has incompatible rule properties.");
        attributes(scale, []); check(scale, ["cfvo", "color"], S);
        const values = elements(scale, "cfvo", S), colors = elements(scale, "color", S);
        if (values.length < 2 || values.length > 3 || colors.length !== values.length) fail(scale, "A color scale requires two or three matched stops.");
        result.push({ ...common, kind, stops: values.map((item, i) => threshold(item, colors[i]!, point, styles)) });
      } else if (kind === "cellIs" || kind === "expression") {
        if (scale) fail(node, "A scalar conditional rule contains a color scale.");
        const style = value(node, "dxfId") === undefined ? {} : styles.differential(integer(node, "dxfId"), node);
        if (kind === "expression") {
          if (formulas.length !== 1) fail(node, "An expression rule requires one formula.");
          result.push({ ...common, ...style, kind, expression: formulas[0]! });
        } else {
          const operator = comparison(node), count = ["between", "notBetween"].includes(operator) ? 2 : 1;
          if (formulas.length !== count) fail(node, "A comparison rule has the wrong number of formulas.");
          result.push({ ...common, ...style, kind, operator, formulas });
        }
      } else unsupported(node, "This conditional formatting rule is outside the declared reading scope.");
    }
  }
  return result;
}
export function readDataValidations(root: XmlNode, sheetId: string): Pick<WorkbookSheet, "dataValidations" | "disableValidationPrompts"> {
  const container = one(root, "dataValidations"); if (!container) return {};
  attributes(container, ["count", "disablePrompts", "xWindow", "yWindow"]); check(container, ["dataValidation"], S);
  for (const key of ["xWindow", "yWindow"]) if (value(container, key) !== undefined) integer(container, key);
  const result: WorkbookDataValidation[] = [];
  for (const node of elements(container, "dataValidation", S)) {
    attributes(node, ["type", "errorStyle", "imeMode", "operator", "allowBlank", "showDropDown", "showInputMessage", "showErrorMessage", "errorTitle", "error", "promptTitle", "prompt", "sqref"]);
    check(node, ["formula1", "formula2"], S);
    if (value(node, "imeMode") !== undefined && value(node, "imeMode") !== "noControl") unsupported(node, "Forced IME modes in data validation are not implemented.");
    const ranges = regions(node), point = origin(ranges), kind = value(node, "type") ?? "none", errorStyle = value(node, "errorStyle") ?? "stop";
    if (errorStyle !== "stop" && errorStyle !== "warning" && errorStyle !== "information") fail(node, "A data validation error style is invalid.");
    const common = { ruleId: `${sheetId}-dv-${result.length + 1}`, ranges, origin: point, errorStyle: errorStyle as WorkbookDataValidation["errorStyle"],
      allowBlank: boolean(node, "allowBlank"), showErrorMessage: boolean(node, "showErrorMessage"), showInputMessage: boolean(node, "showInputMessage"),
      errorTitle: spreadsheetText(value(node, "errorTitle") ?? "", node, 32), error: spreadsheetText(value(node, "error") ?? "", node, 225),
      promptTitle: spreadsheetText(value(node, "promptTitle") ?? "", node, 32), prompt: spreadsheetText(value(node, "prompt") ?? "", node, 255),
    };
    const first = one(node, "formula1"), second = one(node, "formula2"), dropdown = !boolean(node, "showDropDown");
    if (first) formulaText(first); if (second) formulaText(second);
    if (kind === "none") { result.push({ ...common, kind }); continue; }
    if (!first) fail(node, "A data validation rule requires a formula.");
    const expression = formula(formulaText(first), point, first);
    if (kind === "list") result.push({ ...common, kind, expression, dropdown });
    else if (kind === "custom") result.push({ ...common, kind, expression });
    else if (["whole", "decimal", "date", "time", "textLength"].includes(kind)) {
      const operator = comparison(node, "between"), formulas = [expression];
      if (["between", "notBetween"].includes(operator)) {
        if (!second) fail(node, "A between validation rule requires two formulas.");
        formulas.push(formula(formulaText(second), point, second));
      }
      result.push({ ...common, kind: kind as "whole" | "decimal" | "date" | "time" | "textLength", operator, formulas });
    } else unsupported(node, "This data validation type is not supported.");
    // formula2 is inert for list/custom/unary comparisons; Calc writes zero.
  }
  if (value(container, "count") !== undefined && integer(container, "count") !== result.length) fail(container, "The data validation count is inconsistent.");
  return { dataValidations: result, disableValidationPrompts: boolean(container, "disablePrompts") };
}
