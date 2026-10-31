import { A, check, child, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { bool, Context, fontName, integer, points, value } from "./context.js";

const roles = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
export interface Color { css: string; alpha?: number; lumMod?: number; lumOff?: number; tint?: number; shade?: number }
export function color(node: XmlNode): Color {
  let css: string;
  const v = value(node);
  if (node.uri !== A) return unsupported(node);
  if (node.name === "srgbClr" || node.name === "sysClr") {
    const rgb = node.name === "sysClr" ? value(node, "lastClr") : v;
    if (!rgb || !/^[0-9a-f]{6}$/i.test(rgb)) fail(node, "A DrawingML RGB color is invalid.");
    css = "#" + rgb.toLowerCase();
  } else if (node.name === "schemeClr") {
    if (!v || ![...roles, "bg1", "bg2", "tx1", "tx2", "phClr"].includes(v)) fail(node, "A theme color role is invalid.");
    css = v === "phClr" ? "var(--office-placeholder-color,var(--office-theme-dk1))" : "var(--office-theme-" + v + ")";
  } else return unsupported(node, "This DrawingML color representation is not supported.");
  check(node, ["alpha", "lumMod", "lumOff", "tint", "shade"], A);
  const result: Color = { css };
  for (const name of ["alpha", "lumMod", "lumOff", "tint", "shade"] as const) {
    const property = child(node, name, A);
    if (property) {
      const amount = integer(property, "val") / 100000;
      if (amount < 0 || amount > 1) fail(property, "A color transform is outside its supported range.");
      result[name] = amount;
    }
  }
  return result;
}
export function onlyColor(node: XmlNode): Color {
  const list = elements(node, undefined, A);
  if (list.length !== 1) fail(node, "A DrawingML color container requires exactly one color.");
  return color(list[0]!);
}
export function paint(css: Css, prefix: string, source: Color): void {
  css[prefix + "-color"] = source.css;
  for (const key of ["alpha", "lumMod", "lumOff", "tint", "shade"] as const)
    css[prefix + "-" + key.replace(/[A-Z]/g, c => "-" + c.toLowerCase())] = String(source[key] ?? (key === "alpha" || key === "lumMod" || key === "shade" ? 1 : 0));
}
export function fill(source: XmlNode | undefined, prefix = "--office-fill"): Css {
  const css: Css = {};
  if (!source) return css;
  const choices = elements(source, undefined, A).filter(n => ["noFill", "solidFill", "gradFill", "pattFill", "blipFill", "grpFill"].includes(n.name));
  if (choices.length > 1) fail(source, "A DrawingML fill is repeated.");
  const item = choices[0];
  if (!item) return css;
  if (item.name === "noFill") css[prefix + "-kind"] = "none";
  else if (item.name === "solidFill") { css[prefix + "-kind"] = "solid"; paint(css, prefix, onlyColor(item)); }
  else if (item.name === "gradFill") {
    check(item, ["gsLst", "lin", "tileRect"], A);
    const linear = child(item, "lin", A), stops = child(item, "gsLst", A);
    if (!linear || !stops) unsupported(item, "Only explicit linear gradients are supported.");
    const list = elements(stops, "gs", A);
    if (list.length < 2) fail(stops, "A gradient requires at least two stops.");
    css[prefix + "-kind"] = "linear"; css[prefix + "-angle"] = String(integer(linear, "ang", 0) / 60000);
    css[prefix + "-scaled"] = String(bool(linear, "scaled", true));
    css[prefix + "-rotate-with-shape"] = String(bool(item, "rotWithShape", true));
    css[prefix + "-stop-count"] = String(list.length);
    let previous = -1;
    list.forEach((stop, index) => {
      const position = integer(stop, "pos") / 100000;
      if (position < previous || position < 0 || position > 1) fail(stop, "Gradient stop positions must be ordered fractions.");
      previous = position;
      css[prefix + "-stop-" + index + "-position"] = String(position);
      paint(css, prefix + "-stop-" + index, onlyColor(stop));
    });
    const tile = child(item, "tileRect", A);
    if (tile && ["l", "r", "t", "b"].some(name => integer(tile, name, 0) !== 0)) unsupported(tile, "Offset gradient tiles are not supported.");
  } else unsupported(item, "Pattern, picture and group fills are not supported.");
  return css;
}
export function outline(source: XmlNode | undefined, prefix = "--office-line"): Css {
  if (!source) return {};
  check(source, ["noFill", "solidFill", "prstDash", "round", "bevel", "miter", "headEnd", "tailEnd"], A);
  const css = fill(source, prefix), width = points(source, "w", 12700);
  if (width < 0) fail(source, "A line width cannot be negative.");
  css[prefix + "-width"] = width + "pt";
  const dash = child(source, "prstDash", A);
  if (dash) {
    const kind = value(dash);
    if (!kind || !["solid", "dash", "dot", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"].includes(kind)) unsupported(dash, "This dash pattern is not supported.");
    css[prefix + "-dash"] = kind;
  }
  for (const [tag, side] of [["headEnd", "start"], ["tailEnd", "end"]]) {
    const arrow = child(source, tag!, A);
    if (!arrow) continue;
    const kind = value(arrow, "type") ?? "none";
    if (!["none", "triangle", "arrow", "stealth", "diamond", "oval"].includes(kind)) unsupported(arrow);
    css[prefix + "-" + side + "-arrow"] = kind;
    for (const name of ["w", "len"]) if (value(arrow, name)) {
      if (!["sm", "med", "lg"].includes(value(arrow, name)!)) fail(arrow, "An arrow size is invalid.");
      css[prefix + "-" + side + "-" + name] = value(arrow, name)!;
    }
  }
  for (const name of ["cap", "cmpd", "algn"]) if (value(source, name)) css[prefix + "-" + name] = value(source, name)!;
  const join = elements(source, undefined, A).find(n => ["round", "bevel", "miter"].includes(n.name));
  if (join) { css[prefix + "-join"] = join.name; if (join.name === "miter") css[prefix + "-miter-limit"] = String(integer(join, "lim", 800000) / 100000); }
  return css;
}
export function effects(source: XmlNode | undefined): Css {
  if (!source) return {};
  check(source, ["outerShdw"], A);
  const shadow = child(source, "outerShdw", A);
  if (!shadow) return { "--office-shadow-kind": "none" };
  const blur = points(shadow, "blurRad", 0), distance = points(shadow, "dist", 0);
  if (blur < 0 || distance < 0) fail(shadow, "Shadow dimensions cannot be negative.");
  for (const name of ["sx", "sy"]) if (integer(shadow, name, 100000) !== 100000) unsupported(shadow, "Scaled shadows are not supported.");
  for (const name of ["kx", "ky"]) if (integer(shadow, name, 0) !== 0) unsupported(shadow, "Skewed shadows are not supported.");
  const css: Css = { "--office-shadow-kind": "outer", "--office-shadow-blur": blur + "pt", "--office-shadow-distance": distance + "pt",
    "--office-shadow-direction": String(integer(shadow, "dir", 0) / 60000), "--office-shadow-rotate-with-shape": String(bool(shadow, "rotWithShape", true)),
    "--office-shadow-align": value(shadow, "algn") ?? "b" };
  const colors = elements(shadow, undefined, A);
  if (colors.length !== 1) fail(shadow, "A shadow requires one color.");
  paint(css, "--office-shadow", color(colors[0]!));
  return css;
}
export function fontCss(source: XmlNode | undefined, role: "ascii" | "east-asia" | "complex-script"): string | undefined {
  const name = source && value(source, "typeface");
  if (!name) return undefined;
  const match = /^\+(mj|mn)-(lt|ea|cs)$/.exec(name);
  if (match) return "var(--office-" + (match[1] === "mj" ? "major" : "minor") + "-" + ({ lt: "ascii", ea: "east-asia", cs: "complex-script" }[match[2]!] ?? role) + ")";
  if (name.startsWith("+")) unsupported(source!, "This theme font reference is not supported.");
  return fontName(name);
}
export function colorMap(node: XmlNode | undefined): Css {
  if (!node) return {};
  const css: Css = {};
  for (const role of ["bg1", "bg2", "tx1", "tx2", ...roles.slice(4)]) {
    const mapped = value(node, role);
    if (!mapped || !roles.includes(mapped)) fail(node, "A color mapping is incomplete or invalid.");
    css["--office-theme-" + role] = "var(--office-palette-" + mapped + ")";
  }
  return css;
}
export function importTheme(context: Context, path: string): string {
  const known = context.surfaces.get(path);
  if (known) return known.id;
  const source = context.part(path, "theme", "theme", A);
  check(source, ["themeElements", "objectDefaults", "extraClrSchemeLst", "custClrLst", "extLst"], A);
  const theme = child(source, "themeElements", A);
  if (!theme) fail(source, "Theme elements are missing.");
  check(theme, ["clrScheme", "fontScheme", "fmtScheme"], A);
  const colors = child(theme, "clrScheme", A), fonts = child(theme, "fontScheme", A), format = child(theme, "fmtScheme", A);
  if (!colors || !fonts || !format) fail(source, "A presentation theme is incomplete.");
  const css: Css = {}, id = context.id("theme");
  for (const role of roles) {
    const item = child(colors, role, A);
    if (!item) fail(colors, "A required theme color is missing.");
    const v = onlyColor(item);
    if (!v.css.startsWith("#") || Object.keys(v).length !== 1) unsupported(item, "Theme base colors must be untransformed RGB values.");
    css["--office-palette-" + role] = v.css;
    css["--office-theme-" + role] = "var(--office-palette-" + role + ")";
  }
  for (const kind of ["major", "minor"] as const) {
    const parent = child(fonts, kind + "Font", A);
    if (!parent) fail(fonts, "A theme font role is missing.");
    for (const [tag, role] of [["latin", "ascii"], ["ea", "east-asia"], ["cs", "complex-script"]] as const) {
      const name = value(child(parent, tag, A)!, "typeface");
      if (name === undefined) fail(parent, "A theme font name is missing.");
      // An empty East Asian font uses the declared Hans supplemental role.
      const supplement = elements(parent, "font", A).find(n => value(n, "script") === (role === "east-asia" ? "Hans" : "Arab"));
      css["--office-" + kind + "-" + role] = fontName(name || (supplement && value(supplement, "typeface")) || "sans-serif");
      if (role === "ascii") css["--office-" + kind + "-h-ansi"] = css["--office-" + kind + "-" + role]!;
    }
    for (const item of elements(parent, "font", A)) {
      const script = value(item, "script"), name = value(item, "typeface");
      if (!script || !/^[A-Za-z]{4}$/.test(script) || !name) fail(item, "A supplemental theme font is invalid.");
      css["--office-" + kind + "-script-" + script] = fontName(name);
    }
  }
  css.color = "var(--office-theme-dk1)"; css["font-family"] = "var(--office-minor-ascii),var(--office-minor-east-asia)"; css["font-size"] = "18pt";
  context.rule(id, css);
  for (const [name, prefix, map] of [["fillStyleLst", "fill", (n: XmlNode) => fill({ ...n, content: [n] })],
    ["bgFillStyleLst", "background", (n: XmlNode) => fill({ ...n, content: [n] })],
    ["lnStyleLst", "line", outline], ["effectStyleLst", "effect", (n: XmlNode) => { check(n, ["effectLst", "scene3d", "sp3d"], A); if (child(n, "scene3d", A) || child(n, "sp3d", A)) unsupported(n); return effects(child(n, "effectLst", A)); }]] as const) {
    const list = child(format, name, A);
    if (!list) fail(format, "A theme format list is missing.");
    elements(list, undefined, A).forEach((item, index) => {
      const css = map(item);
      for (const key of Object.keys(css)) css[key] = css[key]!.replaceAll("--office-placeholder-color", "--office-" + prefix + "-reference-color");
      context.rule(id + "-" + prefix + "-" + (index + 1), css);
    });
  }
  // Object defaults affect visible source and cannot be silently discarded.
  const defaults = child(source, "objectDefaults", A);
  if (defaults && elements(defaults, undefined, A).length) unsupported(defaults, "Theme object defaults are not yet supported.");
  context.metadata.themes.push({ themeId: id, cssClass: id });
  context.surfaces.set(path, { id, root: { tag: "div", attrs: {}, children: [] }, source, themeId: id });
  return id;
}
