import { check, fail, S, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { cssFontName } from "../ooxml/css-values.js";
import { attributes, boolean, integer, number, one, spreadsheetText, value } from "./values.js";
import { colorDeclarations } from "./colors.js";
import type { SpreadsheetColors } from "./colors.js";

export function fontDeclarations(source: XmlNode, colors: SpreadsheetColors, full: boolean): Css {
  attributes(source, []);
  check(source, ["name", "rFont", "charset", "family", "b", "i", "strike", "outline", "shadow", "condense", "extend", "color", "sz", "u", "vertAlign", "scheme"], S);
  const css: Css = full ? { "font-weight": "normal", "font-style": "normal", "vertical-align": "baseline", "--office-font-underline": "none", "--office-font-strike": "false" } : {};
  const element = (name: string): XmlNode | undefined => {
    const node = one(source, name); if (node) { attributes(node, ["val"]); check(node, [], S); } return node;
  };
  for (const name of ["outline", "shadow", "condense", "extend"]) {
    const node = element(name); if (node && boolean(node, "val", true)) unsupported(node, "This native font effect is not implemented.");
  }
  const family = element("family"), charset = element("charset");
  if (family) css["--office-font-family-class"] = String(integer(family, "val", 0, 5));
  if (charset) css["--office-font-charset"] = String(integer(charset, "val", 0, 255));
  const size = element("sz"); if (size) { const pt = number(size, "val"); if (pt <= 0 || pt > 409) fail(size, "A font size is outside native bounds."); css["font-size"] = String(pt) + "pt"; }
  for (const [tag, property, yes, no] of [["b", "font-weight", "bold", "normal"], ["i", "font-style", "italic", "normal"]] as const) {
    const node = element(tag); if (node) css[property] = boolean(node, "val", true) ? yes : no;
  }
  const underline = element("u"), strike = element("strike");
  if (underline) {
    const kind = value(underline, "val") ?? "single";
    if (!["none", "single", "double", "singleAccounting", "doubleAccounting"].includes(kind)) fail(underline!, "A native underline style is invalid.");
    css["--office-font-underline"] = kind;
  }
  if (strike) css["--office-font-strike"] = String(boolean(strike, "val", true));
  const baseline = element("vertAlign");
  if (baseline) {
    const position = value(baseline, "val");
    if (!position || !["baseline", "superscript", "subscript"].includes(position)) fail(baseline, "A native baseline position is invalid.");
    css["vertical-align"] = position === "superscript" ? "super" : position === "subscript" ? "sub" : "baseline";
  }
  const nameNode = element("name"), runName = element("rFont"), scheme = element("scheme");
  if (nameNode && runName) fail(source, "A font has duplicate typeface definitions.");
  const nativeName = value(nameNode ?? runName, "val"), font = nativeName === undefined ? undefined : cssFontName(spreadsheetText(nativeName, source, 255));
  const role = scheme ? value(scheme, "val") : undefined;
  if (role !== undefined && !["major", "minor", "none"].includes(role)) fail(scheme!, "A font scheme is invalid.");
  if (scheme && role === undefined) fail(scheme, "A font scheme has no value.");
  if (font) css["--office-font-source-name"] = font;
  if (font || role) {
    for (const key of ["ascii", "h-ansi", "east-asia", "complex-script"]) css[`--office-font-${key}`] = role && role !== "none" ? `var(--office-${role}-${key}${font ? "," + font : ""})` : font ?? "var(--office-font-source-name)";
    css["font-family"] = "var(--office-font-ascii),var(--office-font-east-asia)";
  }
  if (role !== undefined || font || full) css["--office-font-scheme"] = role ?? "none";
  const color = one(source, "color");
  if (color || full) {
    Object.assign(css, colorDeclarations("--office-cell-foreground", color ? colors.read(color) : { css: "var(--office-system-foreground)", tint: 0 }));
    css.color = "var(--office-cell-foreground-color)";
  }
  return css;
}
