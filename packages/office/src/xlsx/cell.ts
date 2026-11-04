import { WorkbookErrorValueSchema } from "@dagent-office/core";
import type { SharedWorkbookFormulas, WorkbookCoordinate, WorkbookValue } from "@dagent-office/core";
import { check, fail, S, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, finite, integer, one, range, spreadsheetText, textHtml, value } from "./values.js";
import { richString } from "./strings.js";
import type { RichTextStyles, SharedStrings } from "./strings.js";
import type { SheetStyles } from "./sheet.js";

/** First pass registers shared origins, including origins after followers. The
 * second pass resolves only after the entire group table has been validated. */
export function cellFormula(node: XmlNode, coordinate: WorkbookCoordinate, shared: SharedWorkbookFormulas, pass: 0 | 1): string | undefined {
  const formula = one(node, "f"); if (!formula) return undefined;
  attributes(formula, ["t", "si", "ref", "ca", "aca", "bx"]);
  if (formula.content.some(item => typeof item !== "string")) fail(formula, "A formula must contain only text.");
  if (boolean(formula, "bx")) unsupported(formula, "Cell formulas assigning defined names are not supported.");
  for (const key of ["ca", "aca"]) if (value(formula, key) !== undefined) boolean(formula, key);
  const type = value(formula, "t") ?? "normal", source = text(formula);
  if (type !== "normal" && type !== "shared") unsupported(formula, "Array and data-table formulas are not supported.");
  if (type === "normal") {
    if (!source || value(formula, "si") !== undefined || value(formula, "ref") !== undefined) fail(formula, "A normal cell formula has invalid grouping attributes or empty content.");
    return "=" + source;
  }
  const id = integer(formula, "si"), reference = value(formula, "ref");
  if (pass === 0) {
    if (source) {
      if (!reference) fail(formula, "A shared formula origin has no range.");
      shared.define(id, coordinate, range(reference, formula), "=" + source);
    } else if (reference) fail(formula, "A shared formula follower declares an origin range.");
    return undefined;
  }
  return shared.resolve(id, coordinate);
}

export interface CellSource { readonly html: string; readonly formula: boolean; readonly sharedString: boolean }
export function readCell(node: XmlNode, coordinate: WorkbookCoordinate, styles: SheetStyles & RichTextStyles, strings: SharedStrings, shared: SharedWorkbookFormulas): CellSource {
  attributes(node, ["r", "s", "t", "cm", "vm", "ph"]); check(node, ["f", "v", "is"], S);
  if (integer(node, "cm", 0, 0xffff_ffff, 0) || integer(node, "vm", 0, 0xffff_ffff, 0) || boolean(node, "ph")) unsupported(node, "Cell metadata extensions and phonetic display are not supported.");
  const raw = one(node, "v"), inline = one(node, "is"), type = value(node, "t") ?? "n", expression = cellFormula(node, coordinate, shared, 1);
  if (raw) { attributes(raw, []); if (raw.content.some(item => typeof item !== "string")) fail(raw, "A cell value must contain only text."); }
  if (inline && type !== "inlineStr" || inline && raw || type === "inlineStr" && expression) fail(node, "A cell has incompatible value containers.");
  let kind: WorkbookValue["type"] = "blank", literal: string | undefined, html = "", hasCache = false;
  const content = text(raw);
  if (type === "inlineStr") {
    if (raw) fail(node, "An inline string has a scalar value container.");
    if (inline) { kind = "text"; html = richString(inline, styles).html; hasCache = true; }
  } else if (type === "s") {
    if (!raw || !/^\d+$/.test(content)) fail(node, "A shared string cell has an invalid index.");
    kind = "text"; html = strings.get(Number(content), node).html; hasCache = true;
  } else if (type === "str") {
    if (raw) { kind = "text"; html = textHtml(spreadsheetText(content, node)); hasCache = true; }
  } else if (type === "n") {
    if (raw && content.trim()) { kind = "number"; literal = String(finite(content, node)); html = literal; hasCache = true; }
  } else if (type === "b") {
    if (raw && content) {
      if (content !== "0" && content !== "1") fail(node, "A boolean cell must contain zero or one.");
      kind = "boolean"; literal = content === "1" ? "true" : "false"; html = content === "1" ? "TRUE" : "FALSE"; hasCache = true;
    }
  } else if (type === "e") {
    if (raw && content) {
      if (!WorkbookErrorValueSchema.safeParse(content).success) unsupported(node, "This native cell error value is not supported.");
      kind = "error"; literal = content; html = textHtml(content); hasCache = true;
    }
  } else unsupported(node, "This native cell value type is not supported.");
  const index = value(node, "s"), style = index === undefined ? undefined : styles.cell(integer(node, "s"), node);
  const attrs = ` data-column="${coordinate.column}" data-value-type="${kind}"`
    + (literal === undefined ? "" : ` data-value="${textHtml(literal)}"`)
    + (style ? ` class="${textHtml(style.cssClass)}" data-format-id="${textHtml(style.numberFormatId)}"` : "")
    + (expression === undefined ? "" : ` data-formula="${textHtml(expression)}" data-cache-state="${hasCache ? "imported" : "unavailable"}"`);
  return { html: `<td${attrs}>${html}</td>`, formula: expression !== undefined, sharedString: type === "s" };
}
