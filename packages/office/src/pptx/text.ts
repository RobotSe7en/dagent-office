import { A, attr, check, child, elements, fail, R, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import type { PresentationObject } from "@dagent-office/core";
import { bool, Context, integer, points, value } from "./context.js";
import type { Html } from "./context.js";
import { fill, fontCss, onlyColor, paint } from "./styles.js";

export function importHyperlink(context: Context, hyperlink: XmlNode | undefined, targetNode: Html, part: string): void {
  if (!hyperlink) return;
  if (value(hyperlink, "action") && value(hyperlink, "action") !== "ppaction://hlinksldjump") unsupported(hyperlink, "Active presentation actions are not supported.");
  if (child(hyperlink, "snd", A)) unsupported(hyperlink, "Hyperlink sound actions are not supported.");
  const id = attr(hyperlink, "id", R), relation = id && context.opc.relationships(part).get(id);
  if (!relation) fail(hyperlink, "A hyperlink relationship is missing.");
  if (relation.type === R + "/slide" && !relation.external) {
    const slideId = context.slidePaths.get(relation.target);
    if (!slideId) fail(hyperlink, "A hyperlink targets an unknown slide.");
    context.metadata.hyperlinks.push({ nodeId: targetNode.attrs["data-node-id"]!, slideId });
  } else {
    const target = context.relation(part, id, "hyperlink", true);
    if (!target.external || !/^(https?:\/\/|mailto:)/i.test(target.target) || /[\u0000-\u0020\u007f]/.test(target.target)) unsupported(hyperlink, "This hyperlink target is not supported.");
    context.metadata.hyperlinks.push({ nodeId: targetNode.attrs["data-node-id"]!, target: target.target });
  }
}

export function runProperties(node: XmlNode | undefined): Css {
  const css: Css = {};
  if (!node) return css;
  check(node, ["ln", "noFill", "solidFill", "effectLst", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "uLnTx", "uFillTx", "extLst"], A);
  for (const name of ["b", "i"]) if (value(node, name) !== undefined) css[name === "b" ? "font-weight" : "font-style"] = bool(node, name) ? name === "b" ? "bold" : "italic" : "normal";
  if (value(node, "sz") !== undefined) {
    const size = integer(node, "sz") / 100;
    if (size <= 0) fail(node, "A font size must be positive.");
    css["font-size"] = size + "pt";
  }
  if (value(node, "spc") !== undefined) css["letter-spacing"] = integer(node, "spc") / 100 + "pt";
  if (value(node, "baseline") !== undefined) css["vertical-align"] = integer(node, "baseline") / 1000 + "%";
  const underline = value(node, "u");
  if (underline && !["none", "sng", "dbl", "heavy", "dotted", "dash", "wavy"].includes(underline)) unsupported(node, "This text underline is not supported.");
  const strike = value(node, "strike");
  if (strike && !["noStrike", "sngStrike", "dblStrike"].includes(strike)) unsupported(node);
  if (underline || strike) css["text-decoration-line"] = [underline && underline !== "none" ? "underline" : "", strike && strike !== "noStrike" ? "line-through" : ""].filter(Boolean).join(" ") || "none";
  if (underline === "dbl" || strike === "dblStrike") css["text-decoration-style"] = "double";
  if (underline === "dotted") css["text-decoration-style"] = "dotted";
  if (underline === "dash") css["text-decoration-style"] = "dashed";
  if (underline === "wavy") css["text-decoration-style"] = "wavy";
  if (underline) css["--office-underline"] = underline;
  if (strike) css["--office-strike"] = strike;
  if (value(node, "cap")) {
    const cap = value(node, "cap");
    if (!["none", "all", "small"].includes(cap!)) unsupported(node);
    css["text-transform"] = cap === "all" ? "uppercase" : "none";
    css["font-variant"] = cap === "small" ? "small-caps" : "normal";
  }
  for (const name of ["rtl", "kumimoji"]) if (bool(node, name)) unsupported(node, "Vertical or right-to-left text is outside the current text scope.");
  const solid = child(node, "solidFill", A);
  if (solid) { const c = onlyColor(solid); paint(css, "--office-text", c); css.color = c.css; }
  if (child(node, "noFill", A)) { css.color = "transparent"; css["--office-text-color"] = "transparent"; }
  for (const [tag, role] of [["latin", "ascii"], ["ea", "east-asia"], ["cs", "complex-script"]] as const) {
    const font = fontCss(child(node, tag, A), role);
    if (font) css["--office-font-" + role] = font;
    if (font && role === "ascii") { css["--office-font-h-ansi"] = font; css["font-family"] = font + ",var(--office-font-east-asia,var(--office-minor-east-asia))"; }
  }
  if (child(node, "sym", A)) unsupported(node, "Symbol-font remapping is not supported.");
  const outline = child(node, "ln", A), effects = child(node, "effectLst", A);
  if (outline && !(child(outline, "noFill", A)) || effects && elements(effects, undefined, A).length) unsupported(node, "Text outlines and text effects are not supported.");
  if (child(node, "hlinkMouseOver", A)) unsupported(node, "Hover actions are not supported.");
  return css;
}

export function paragraphProperties(node: XmlNode | undefined, listDefaults = false): Css {
  const css: Css = {};
  if (!node) return css;
  check(node, ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "tabLst", "defRPr", "extLst"], A);
  if (bool(node, "rtl")) unsupported(node, "Right-to-left paragraphs are not supported.");
  const align = value(node, "algn");
  if (align) {
    const mapped: Record<string, string> = { l: "left", ctr: "center", r: "right", just: "justify", dist: "justify", thaiDist: "justify", justLow: "justify" };
    if (!mapped[align]) unsupported(node);
    css["text-align"] = mapped[align]!;
  }
  for (const [native, property] of [["marL", "margin-left"], ["marR", "margin-right"], ["indent", "text-indent"]])
    if (value(node, native!) !== undefined) css[property!] = points(node, native!) + "pt";
  for (const [name, property] of [["lnSpc", "line-height"], ["spcBef", "margin-top"], ["spcAft", "margin-bottom"]]) {
    const parent = child(node, name!, A);
    if (!parent) continue;
    check(parent, ["spcPct", "spcPts"], A);
    const percent = child(parent, "spcPct", A), absolute = child(parent, "spcPts", A);
    if (!!percent === !!absolute) fail(parent, "Paragraph spacing requires exactly one value.");
    const n = percent ? integer(percent, "val") / 100000 : integer(absolute!, "val") / 100;
    if (n < 0) fail(parent, "Paragraph spacing cannot be negative.");
    // Before/after percentage spacing depends on the inherited font size.
    css[property!] = percent ? property === "line-height" ? String(n) : n + "em" : n + "pt";
  }
  const tabs = child(node, "tabLst", A);
  if (tabs && elements(tabs, undefined, A).length) unsupported(tabs, "Explicit presentation tab stops are not yet supported.");
  if (value(node, "defTabSz") !== undefined) {
    const size = points(node, "defTabSz");
    if (size <= 0) fail(node, "A default tab interval must be positive.");
    css["tab-size"] = size + "pt";
  }
  const bulletColor = child(node, "buClr", A);
  if (bulletColor) paint(css, "--office-bullet", onlyColor(bulletColor));
  const bulletFont = child(node, "buFont", A);
  if (bulletFont) css["--office-bullet-font"] = fontCss(bulletFont, "ascii") ?? "inherit";
  const bulletPercent = child(node, "buSzPct", A), bulletPoints = child(node, "buSzPts", A);
  if (bulletPercent) css["--office-bullet-size"] = integer(bulletPercent, "val") / 100000 + "em";
  if (bulletPoints) css["--office-bullet-size"] = integer(bulletPoints, "val") / 100 + "pt";
  if (listDefaults) {
    const bullet = child(node, "buChar", A), numbering = child(node, "buAutoNum", A), none = child(node, "buNone", A);
    if ([bullet, numbering, none].filter(Boolean).length > 1) fail(node, "Paragraph list defaults conflict.");
    if (bullet) { css["--office-list-kind"] = "bullet"; css["--office-list-character"] = '"' + (value(bullet, "char") ?? fail(bullet, "A bullet character is missing.")).replace(/["\\]/g, "\\$&") + '"'; }
    if (numbering) { css["--office-list-kind"] = "number"; css["--office-list-numbering"] = value(numbering, "type") ?? fail(numbering, "A numbering type is missing."); css["--office-list-start"] = String(integer(numbering, "startAt", 1)); }
    if (none) css["--office-list-kind"] = "none";
  }
  return { ...css, ...runProperties(child(node, "defRPr", A)) };
}

export function importTextStyle(context: Context, source: XmlNode | undefined, className: string): void {
  if (!source) return;
  check(source, ["defPPr", ...Array.from({ length: 9 }, (_, index) => "lvl" + (index + 1) + "pPr")], A);
  context.rule(className, paragraphProperties(child(source, "defPPr", A), true));
  for (let level = 0; level < 9; level++) {
    const properties = child(source, "lvl" + (level + 1) + "pPr", A);
    if (properties) context.rule(className + "-level-" + level, paragraphProperties(properties, true));
  }
}

export function importText(context: Context, body: XmlNode, target: Html, part: string): Extract<PresentationObject, { kind: "text" }>["text"] {
  check(body, ["bodyPr", "lstStyle", "p"], A);
  const properties = child(body, "bodyPr", A);
  if (!properties) fail(body, "A DrawingML text body requires body properties.");
  check(properties, ["noAutofit", "normAutofit", "spAutoFit", "prstTxWarp", "scene3d", "sp3d", "extLst"], A);
  for (const name of ["rot", "numCol", "spcCol"]) if (integer(properties, name, name === "numCol" ? 1 : 0) !== (name === "numCol" ? 1 : 0)) unsupported(properties, "Rotated or multicolumn text bodies are not supported.");
  if (value(properties, "vert") && value(properties, "vert") !== "horz" || bool(properties, "rtlCol")) unsupported(properties, "Only horizontal text is supported.");
  const warp = child(properties, "prstTxWarp", A);
  if (warp && value(warp, "prst") !== "textNoShape" || child(properties, "scene3d", A) || child(properties, "sp3d", A)) unsupported(properties, "Text warping and 3D are not supported.");
  const kinds = ["noAutofit", "normAutofit", "spAutoFit"].filter(name => child(properties, name, A));
  if (kinds.length > 1) fail(properties, "Text auto-fit is repeated.");
  const normal = child(properties, "normAutofit", A);
  const wrap = value(properties, "wrap") ?? "square";
  if (!["square", "none"].includes(wrap)) fail(properties, "Text wrapping is invalid.");
  const result: NonNullable<Extract<PresentationObject, { kind: "text" }>["text"]> = { autoFit: kinds[0] === "normAutofit" ? "normal" : kinds[0] === "spAutoFit" ? "shape" : "none", wrap: wrap === "square" };
  if (normal) {
    result.fontScale = integer(normal, "fontScale", 100000) / 100000;
    result.lineSpacingReduction = integer(normal, "lnSpcReduction", 0) / 100000;
    if (result.fontScale <= 0 || result.fontScale > 1 || result.lineSpacingReduction < 0 || result.lineSpacingReduction > 1) fail(normal, "Text auto-fit scales are invalid.");
  }
  const css: Css = { ...fill(properties, "--office-text-fill"), "tab-size": points(properties, "defTabSz", 914400) + "pt" };
  for (const [name, side, fallback] of [["lIns", "left", 91440], ["rIns", "right", 91440], ["tIns", "top", 45720], ["bIns", "bottom", 45720]] as const) {
    const amount = points(properties, name, fallback);
    if (amount < 0) fail(properties, "Text insets cannot be negative.");
    css["padding-" + side] = amount + "pt";
  }
  const anchor = value(properties, "anchor") ?? "t";
  if (!["t", "ctr", "b", "just", "dist"].includes(anchor)) unsupported(properties);
  css["--office-text-anchor"] = anchor; css["white-space"] = result.wrap ? "pre-wrap" : "pre";
  css["--office-text-anchor-center"] = String(bool(properties, "anchorCtr"));
  target.attrs.style = (target.attrs.style ? target.attrs.style + ";" : "") + Object.entries(css).map(([k, v]) => k + ":" + v).join(";");
  const listStyle = context.id("text-style");
  importTextStyle(context, child(body, "lstStyle", A), listStyle);
  target.attrs.class = [target.attrs.class, listStyle].filter(Boolean).join(" ");
  const paragraphs = elements(body, "p", A);
  if (!paragraphs.length) fail(body, "A text body requires at least one paragraph.");
  for (const source of paragraphs) {
    check(source, ["pPr", "r", "br", "fld", "endParaRPr"], A);
    const pPr = child(source, "pPr", A), level = pPr ? integer(pPr, "lvl", 0) : 0;
    if (level < 0 || level > 8) fail(source, "A paragraph outline level is invalid.");
    const p = context.make("p", { "data-office-level": String(level), class: listStyle + "-level-" + level });
    context.style(p, paragraphProperties(pPr));
    const bullet = child(pPr, "buChar", A), numbering = child(pPr, "buAutoNum", A);
    if (bullet && numbering) fail(source, "A paragraph has conflicting list types.");
    if (bullet) { const character = value(bullet, "char") ?? fail(bullet, "A bullet character is missing."); context.bytes(character.length * 6); p.attrs["data-office-bullet"] = character; }
    if (numbering) {
      const format = value(numbering, "type");
      if (!format || !["arabicPeriod", "arabicParenR", "arabicParenBoth", "romanUcPeriod", "romanLcPeriod", "alphaLcPeriod", "alphaUcPeriod", "arabicPlain"].includes(format)) unsupported(numbering, "This list numbering format is not supported.");
      p.attrs["data-office-numbering"] = format; p.attrs["data-office-numbering-start"] = String(integer(numbering, "startAt", 1));
    }
    if (child(pPr, "buNone", A)) p.attrs["data-office-list-none"] = "true";
    for (const sourceRun of elements(source, undefined, A)) {
      if (sourceRun.name === "pPr") continue;
      if (sourceRun.name === "endParaRPr") { const c = context.id("paragraph-end"); context.rule(c, runProperties(sourceRun)); p.attrs["data-office-end-style"] = c; continue; }
      if (sourceRun.name === "br") { p.children.push(context.make("br")); continue; }
      const field = sourceRun.name === "fld";
      check(sourceRun, field ? ["rPr", "pPr", "t"] : ["rPr", "t"], A);
      const run = context.make("span");
      const rPr = child(sourceRun, "rPr", A);
      context.style(run, runProperties(rPr));
      const language = rPr && value(rPr, "lang");
      if (language) run.attrs.lang = language;
      if (field) {
        const type = value(sourceRun, "type");
        if (!type || !/^(slidenum|datetime(?:[1-9]|1[0-3])?)$/.test(type)) unsupported(sourceRun, "This presentation field type is not supported.");
        run.attrs["data-office-field"] = type;
        if (value(sourceRun, "id")) run.attrs["data-office-field-id"] = value(sourceRun, "id")!;
      }
      const t = child(sourceRun, "t", A);
      if (!t) fail(sourceRun, "A text run is missing its text value.");
      context.content(run, text(t));
      importHyperlink(context, child(rPr, "hlinkClick", A), run, part);
      p.children.push(run);
    }
    target.children.push(p);
  }
  return result;
}
