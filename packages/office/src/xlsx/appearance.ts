import { check, fail, S, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { attributes, boolean, integer, one, value } from "./values.js";
import { colorDeclarations } from "./colors.js";
import type { SpreadsheetColors } from "./colors.js";

export function alignmentDeclarations(node: XmlNode | undefined, full: boolean): Css {
  const css: Css = {};
  if (node) {
    attributes(node, ["horizontal", "vertical", "textRotation", "wrapText", "shrinkToFit", "indent", "relativeIndent", "justifyLastLine", "readingOrder"]); check(node, [], S);
    if (integer(node, "textRotation", 0, 255, 0)) unsupported(node, "Rotated and stacked cell text is not implemented.");
  }
  for (const [key, choices, fallback] of [
    ["horizontal", ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"], "general"],
    ["vertical", ["top", "center", "bottom", "justify", "distributed"], "bottom"],
  ] as const) {
    const setting = value(node, key);
    if (setting !== undefined && !choices.some(choice => choice === setting)) fail(node!, "A native cell alignment is invalid.");
    if (setting !== undefined || full) css[`--office-cell-${key}`] = setting ?? fallback;
  }
  for (const [key, property] of [["wrapText", "--office-cell-wrap"], ["shrinkToFit", "--office-cell-shrink-to-fit"], ["justifyLastLine", "--office-cell-justify-last-line"]] as const) {
    if (value(node, key) !== undefined || full) css[property] = String(boolean(node, key));
  }
  if (node && value(node, "indent") !== undefined || full) css["--office-cell-indent"] = node ? String(integer(node, "indent", 0, 255, 0)) : "0";
  if (node && value(node, "relativeIndent") !== undefined) css["--office-cell-relative-indent"] = String(integer(node, "relativeIndent", -15, 15));
  if (node && value(node, "readingOrder") !== undefined || full) css["--office-cell-reading-order"] = ["context", "ltr", "rtl"][node ? integer(node, "readingOrder", 0, 2, 0) : 0]!;
  return css;
}

export function fillDeclarations(node: XmlNode, colors: SpreadsheetColors): Css {
  attributes(node, []); check(node, ["patternFill", "gradientFill"], S);
  const pattern = one(node, "patternFill"), gradient = one(node, "gradientFill");
  if (pattern && gradient || !pattern && !gradient) fail(node, "A fill requires one definition.");
  if (gradient) unsupported(gradient, "Gradient cell fills are not implemented.");
  attributes(pattern!, ["patternType"]); check(pattern, ["fgColor", "bgColor"], S);
  const kind = value(pattern, "patternType") ?? "none", foreground = one(pattern, "fgColor"), background = one(pattern, "bgColor"), css: Css = { "--office-cell-fill-kind": kind };
  if (kind !== "none" && kind !== "solid") unsupported(pattern!, "Patterned cell fills are not implemented.");
  if (foreground) Object.assign(css, colorDeclarations("--office-cell-fill", colors.read(foreground)));
  if (background) Object.assign(css, colorDeclarations("--office-cell-fill-background", colors.read(background)));
  if (kind === "solid" && !foreground) fail(pattern!, "A solid cell fill requires a foreground color.");
  return css;
}
const borderKinds = new Set(["none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"]);
export function borderDeclarations(node: XmlNode, colors: SpreadsheetColors, full: boolean): Css {
  attributes(node, ["diagonalUp", "diagonalDown", "outline"]); check(node, ["left", "right", "top", "bottom", "start", "end", "diagonal", "vertical", "horizontal"], S);
  if (boolean(node, "diagonalUp") || boolean(node, "diagonalDown")) unsupported(node, "Diagonal cell borders are not implemented.");
  if (value(node, "outline") !== undefined) boolean(node, "outline");
  const css: Css = {};
  for (const side of ["left", "right", "top", "bottom", "start", "end", "diagonal", "vertical", "horizontal"]) {
    const border = one(node, side);
    if (!border) { if (full && ["left", "right", "top", "bottom"].includes(side)) css[`--office-cell-border-${side}-kind`] = "none"; continue; }
    attributes(border, ["style"]); check(border, ["color"], S);
    const kind = value(border, "style") ?? "none", color = one(border, "color");
    if (!borderKinds.has(kind)) fail(border, "A native border style is invalid.");
    if (["diagonal", "vertical", "horizontal"].includes(side)) {
      if (kind !== "none") unsupported(border, "This internal or diagonal border is not implemented.");
      if (color) colors.read(color); continue;
    }
    css[`--office-cell-border-${side}-kind`] = kind;
    Object.assign(css, colorDeclarations(`--office-cell-border-${side}`, color ? colors.read(color) : { css: "var(--office-system-foreground)", tint: 0 }));
  }
  return css;
}
