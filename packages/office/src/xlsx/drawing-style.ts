import { A, check, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, integer, one, value } from "./values.js";

function empty(node: XmlNode, names: readonly string[] = []): void { attributes(node, names); check(node, [], A); }
function enumeration(node: XmlNode, name: string, values: readonly string[]): void {
  const raw = value(node, name); if (raw !== undefined && !values.includes(raw)) fail(node, "A DrawingML appearance value is invalid.");
}
function choice(node: XmlNode, names: readonly string[]): void {
  if (elements(node, undefined, A).filter(item => names.includes(item.name)).length > 1) fail(node, "A DrawingML appearance choice is repeated.");
}
/** The shared appearance source represents these transforms in this order.
 * Preserve that meaning instead of silently reordering arbitrary native lists. */
export function validateDrawingColor(node: XmlNode): void {
  if (node.uri !== A || !["srgbClr", "sysClr", "schemeClr"].includes(node.name)) unsupported(node, "This drawing color is not supported.");
  attributes(node, node.name === "sysClr" ? ["val", "lastClr"] : ["val"]);
  const order = ["lumMod", "lumOff", "shade", "tint"];
  check(node, [...order, "alpha"], A); let previous = -1;
  for (const transform of elements(node, undefined, A)) {
    one(node, transform.name, A); empty(transform, ["val"]); integer(transform, "val", 0, 100000);
    if (transform.name === "alpha") continue;
    const index = order.indexOf(transform.name);
    if (index < previous) unsupported(node, "This ordered color-transform sequence cannot be represented by the drawing appearance source.");
    previous = index;
  }
}
export function validateDrawingColorContainer(node: XmlNode, optional = false): void {
  check(node, ["srgbClr", "sysClr", "schemeClr"], A);
  const colors = elements(node, undefined, A);
  if (colors.length > 1 || !optional && colors.length !== 1) fail(node, "A drawing appearance requires exactly one color.");
  colors.forEach(validateDrawingColor);
}
export function validateDrawingFill(node: XmlNode): void {
  if (node.uri !== A) unsupported(node);
  if (node.name === "noFill") { empty(node); return; }
  if (node.name === "solidFill") { attributes(node, []); validateDrawingColorContainer(node); return; }
  if (node.name !== "gradFill") unsupported(node, "This drawing fill is not supported.");
  attributes(node, ["flip", "rotWithShape"]); check(node, ["gsLst", "lin", "tileRect"], A);
  if (value(node, "flip") !== undefined && value(node, "flip") !== "none") unsupported(node, "Flipped gradient tiles are not supported.");
  boolean(node, "rotWithShape", true);
  const stops = one(node, "gsLst", A), linear = one(node, "lin", A), tile = one(node, "tileRect", A);
  if (!stops || !linear) unsupported(node, "Only explicit linear drawing gradients are supported.");
  attributes(stops, []); check(stops, ["gs"], A);
  const values = elements(stops, "gs", A);
  if (values.length < 2 || values.length > 1024) fail(stops, "A drawing gradient requires between 2 and 1024 stops.");
  for (const stop of values) { attributes(stop, ["pos"]); integer(stop, "pos", 0, 100000); validateDrawingColorContainer(stop); }
  empty(linear, ["ang", "scaled"]); integer(linear, "ang", 0, 21600000 - 1, 0); boolean(linear, "scaled", true);
  if (tile) { empty(tile, ["l", "r", "t", "b"]); for (const side of ["l", "r", "t", "b"]) if (integer(tile, side, -2147483648, 2147483647, 0)) unsupported(tile, "Offset gradient tiles are not supported."); }
}
export function validateDrawingLine(node: XmlNode): void {
  attributes(node, ["w", "cap", "cmpd", "algn"]);
  check(node, ["noFill", "solidFill", "prstDash", "round", "bevel", "miter", "headEnd", "tailEnd"], A);
  integer(node, "w", 0, 20116800, 12700);
  enumeration(node, "cap", ["rnd", "sq", "flat"]); enumeration(node, "cmpd", ["sng", "dbl", "thickThin", "thinThick", "tri"]); enumeration(node, "algn", ["ctr", "in"]);
  choice(node, ["noFill", "solidFill"]); choice(node, ["round", "bevel", "miter"]);
  for (const item of elements(node, undefined, A)) {
    one(node, item.name, A);
    if (["noFill", "solidFill"].includes(item.name)) validateDrawingFill(item);
    else if (item.name === "prstDash") empty(item, ["val"]);
    else if (item.name === "miter") { empty(item, ["lim"]); integer(item, "lim", 0, 2147483647, 800000); }
    else if (["headEnd", "tailEnd"].includes(item.name)) empty(item, ["type", "w", "len"]);
    else empty(item);
  }
}
export function validateDrawingEffects(node: XmlNode): void {
  attributes(node, []); check(node, ["outerShdw"], A);
  const shadow = one(node, "outerShdw", A); if (!shadow) return;
  attributes(shadow, ["blurRad", "dist", "dir", "sx", "sy", "kx", "ky", "algn", "rotWithShape"]);
  validateDrawingColorContainer(shadow);
  integer(shadow, "dir", 0, 21600000 - 1, 0);
  for (const name of ["blurRad", "dist"]) integer(shadow, name, 0, 27273042316900, 0);
  enumeration(shadow, "algn", ["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"]);
  boolean(shadow, "rotWithShape", true);
}
