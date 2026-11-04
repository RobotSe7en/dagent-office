import type { DrawingTextProperties } from "@dagent-office/core";
import { A, check, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Html } from "../drawingml/html.js";
import { importText, paragraphProperties, runProperties, textBodyProperties } from "../drawingml/text.js";
import type { Css } from "../ooxml/styles.js";
import { attributes, boolean, integer, one, value } from "./values.js";
import { validateDrawingColor, validateDrawingFill, validateDrawingLine } from "./drawing-style.js";
import type { WorkbookContent } from "./content.js";
import type { XlsxStyles } from "./styles.js";

const attributesByTag: Readonly<Record<string, readonly string[]>> = {
  txBody: [], bodyPr: ["rot", "vert", "wrap", "lIns", "rIns", "tIns", "bIns", "numCol", "spcCol", "rtlCol", "anchor", "anchorCtr"],
  lstStyle: [], p: [], r: [], br: [],
  rPr: ["lang", "sz", "b", "i", "u", "strike", "cap", "spc", "baseline", "dirty", "err", "smtClean", "smtId", "noProof"],
  pPr: ["marL", "marR", "lvl", "indent", "algn", "defTabSz", "rtl"],
  noAutofit: [], normAutofit: ["fontScale", "lnSpcReduction"], spAutoFit: [],
  latin: ["typeface", "panose", "pitchFamily", "charset"], ea: ["typeface", "panose", "pitchFamily", "charset"], cs: ["typeface", "panose", "pitchFamily", "charset"],
  effectLst: [], uFillTx: [], uLnTx: [],
  lnSpc: [], spcBef: [], spcAft: [], spcPct: ["val"], spcPts: ["val"],
  buClrTx: [], buSzTx: [], buFontTx: [], buNone: [], buClr: [], buSzPct: ["val"], buSzPts: ["val"],
  buFont: ["typeface", "panose", "pitchFamily", "charset"], buAutoNum: ["type", "startAt"], buChar: ["char"], tabLst: [], tab: ["pos", "algn"],
};
function validateTextTree(body: XmlNode): void {
  const walk = (node: XmlNode, root = false, parent?: XmlNode): void => {
    if (!root && node.uri !== A) unsupported(node);
    if (["srgbClr", "sysClr", "schemeClr"].includes(node.name)) { validateDrawingColor(node); return; }
    if (["solidFill", "noFill"].includes(node.name)) { validateDrawingFill(node); return; }
    if (node.name === "ln") { validateDrawingLine(node); return; }
    if (node.name === "hlinkClick") {
      if (parent?.name !== "rPr") unsupported(node, "Default-style and paragraph-end hyperlinks are not supported.");
      return; // Consumed by the relationship-aware callback below.
    }
    if (node.name === "t") {
      attributes(node, ["space"], "http://www.w3.org/XML/1998/namespace");
      const space = node.attrs.get("http://www.w3.org/XML/1998/namespace space");
      if (space !== undefined && !["default", "preserve"].includes(space)) fail(node, "A drawing text whitespace mode is invalid.");
      if (node.content.some(child => typeof child !== "string")) fail(node, "Drawing text must contain only characters."); return;
    }
    const key = root ? "txBody" : node.name === "endParaRPr" || node.name === "defRPr" ? "rPr" : node.name === "defPPr" || /^lvl[1-9]pPr$/.test(node.name) ? "pPr" : node.name;
    const allowed = attributesByTag[key]; if (!allowed) unsupported(node, "This drawing text property is not supported."); attributes(node, allowed);
    if (key === "rPr") {
      for (const name of ["dirty", "err", "smtClean", "noProof"]) boolean(node, name);
      if (value(node, "smtId") !== undefined) integer(node, "smtId");
      if (value(node, "sz") !== undefined) integer(node, "sz", 100, 400000);
    }
    if (["effectLst", "uFillTx", "uLnTx", "noAutofit", "normAutofit", "spAutoFit", "latin", "ea", "cs", "buFont", "buChar", "buAutoNum", "buClrTx", "buSzTx", "buFontTx", "buNone", "buSzPct", "buSzPts", "spcPct", "spcPts", "tab"].includes(key)) check(node, [], A);
    for (const item of elements(node, undefined, A)) {
      if (!["p", "r", "br", "tab"].includes(item.name)) one(node, item.name, A);
      walk(item, false, node);
    }
    if (node.content.some(item => typeof item === "string" && item.trim())) fail(node, "Unexpected characters outside a drawing text run.");
    if (node.content.some(item => typeof item !== "string" && item.uri !== A)) unsupported(node);
  };
  walk(body, true);
}
export function drawingTextStyle(source: XmlNode | undefined): Css {
  if (!source) return {};
  validateTextTree(source); check(source, ["bodyPr", "lstStyle", "p"], A);
  const properties = one(source, "bodyPr", A), list = one(source, "lstStyle", A), paragraph = one(source, "p", A);
  if (!properties || !paragraph) fail(source, "A chart text style requires body and paragraph properties.");
  if (list && elements(list, undefined, A).length) unsupported(list, "Chart text style lists are not supported.");
  check(paragraph, ["pPr", "endParaRPr"], A);
  const body = textBodyProperties(properties), css = { ...body.css, ...paragraphProperties(one(paragraph, "pPr", A)), ...runProperties(one(paragraph, "endParaRPr", A)) };
  if (body.text.wrap !== undefined) css["--office-text-wrap"] = String(body.text.wrap);
  if (body.text.autoFit !== undefined) css["--office-text-auto-fit"] = body.text.autoFit;
  if (body.text.fontScale !== undefined) css["--office-text-font-scale"] = String(body.text.fontScale);
  if (body.text.lineSpacingReduction !== undefined) css["--office-text-line-spacing-reduction"] = String(body.text.lineSpacingReduction);
  return css;
}
export function readDrawingText(body: XmlNode, target: Html, content: WorkbookContent, styles: XlsxStyles, hyperlink: (source: XmlNode | undefined, target: Html, part: string) => void): DrawingTextProperties {
  validateTextTree(body);
  return importText({
    id(prefix) {
      const id = content.id();
      // The common mapper refers to every level, including implicit defaults.
      if (prefix === "text-style") { styles.drawingRule(id, {}); for (let level = 0; level < 9; level++) styles.drawingRule(`${id}-level-${level}`, {}); }
      return id;
    },
    make: (tag, attrs) => content.make(tag, attrs), bytes: amount => content.reserve(amount),
    rule: (name, css) => styles.drawingRule(name, css),
    style(node, css) { const name = styles.objectStyle(css); if (name) node.attrs.class = [node.attrs.class, name].filter(Boolean).join(" "); },
    content(node, text) { content.reserve(text.length * 6); node.children.push(text); },
    hyperlink, field: source => unsupported(source, "Worksheet drawing fields are not supported."),
  }, body, target, body.part);
}
