import { lexer } from "css-tree";
import type { Css } from "../ooxml/styles.js";
import { attr, check, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { one, value } from "./values.js";
import { VML, VML_EXCEL, VML_OFFICE, vmlAttributes, vmlBoolean, vmlStyle } from "./vml-notes.js";
import type { NativeVmlNote } from "./vml-notes.js";

function color(source: string, node: XmlNode): string {
  if (!/^(?:#[\da-f]{3}|#[\da-f]{6}|[a-z]+)$/i.test(source) || lexer.matchProperty("color", source).error
    || /^(?:inherit|initial|unset|revert|currentcolor|transparent)$/i.test(source)) unsupported(node, "This VML note color is outside the solid color boundary.");
  return source.toLowerCase();
}
function length(source: string, node: XmlNode, unitlessEmu = false): number {
  const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(pt|px|in|cm|mm|pc|pi)?$/.exec(source.trim());
  if (!match || !match[2] && !unitlessEmu && Number(match[1]) !== 0) unsupported(node, "This VML note length requires supported explicit units.");
  const factors: Record<string, number> = { pt: 1, px: 0.75, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, pc: 12, pi: 12 };
  const result = Number(match[1]) * (match[2] ? factors[match[2]]! : 1 / 12700);
  if (!Number.isFinite(result) || Math.abs(result) > 2147483647) fail(node, "A VML note length exceeds its supported bounds.");
  return result;
}
function opacity(source: string | undefined, node: XmlNode): string {
  if (source === undefined) return "1";
  const match = /^([+]?(?:\d+(?:\.\d*)?|\.\d+))(f)?$/.exec(source);
  if (!match) fail(node, "A VML note opacity is invalid.");
  const result = Number(match[1]) / (match[2] ? 65536 : 1);
  if (result < 0 || result > 1) fail(node, "A VML note opacity exceeds its range."); return String(result);
}
function path(node: XmlNode | undefined, definition: boolean): void {
  if (!node) return;
  vmlAttributes(node, ["gradientshapeok"], [[VML_OFFICE, "connecttype"]]); check(node, [], VML);
  if (value(node, "gradientshapeok") !== undefined) vmlBoolean(value(node, "gradientshapeok"), node, true);
  const connection = attr(node, "connecttype", VML_OFFICE);
  if (connection !== undefined && connection !== (definition ? "rect" : "none")) unsupported(node, "This VML note connection geometry is not supported.");
}
/** Validate even unused note type definitions before discarding their native
 * rectangular path encoding. Their only permitted appearance is the default
 * miter border, which is represented by the note CSS below. */
export function validateVmlNoteType(node: XmlNode): void {
  vmlAttributes(node, ["id", "coordsize", "path"], [[VML_OFFICE, "spt"]]); check(node, ["stroke", "path"], VML);
  const coordinates = value(node, "coordsize")?.replace(/\s/g, ""), source = value(node, "path")?.replace(/\s/g, "").toLowerCase();
  if (attr(node, "spt", VML_OFFICE) !== "202" || coordinates !== "21600,21600"
    || !["m,l,21600r21600,l21600,xe", "m,l,21600l21600,21600l21600,xe"].includes(source ?? "")) unsupported(node, "Only the native rectangular note shape type is supported.");
  const stroke = one(node, "stroke", VML);
  if (stroke) { vmlAttributes(stroke, ["joinstyle"]); check(stroke, [], VML); if (value(stroke, "joinstyle") !== undefined && value(stroke, "joinstyle") !== "miter") unsupported(stroke); }
  path(one(node, "path", VML), true);
}

/** Convert bounded solid note appearance into source CSS. VML is never copied
 * into executable/foreign HTML or used as an alternate rendering payload. */
export function vmlNoteAppearance(note: NativeVmlNote): Css {
  const shape = note.shape;
  if (!note.shapeType) unsupported(shape, "A note requires its supported rectangular type definition.");
  validateVmlNoteType(note.shapeType);
  vmlAttributes(shape, ["id", "type", "style", "fillcolor", "filled", "strokecolor", "strokeweight", "stroked", "alt"], [[VML_OFFICE, "insetmode"], [VML_OFFICE, "allowincell"], [VML_OFFICE, "spid"]]);
  for (const node of shape.content) {
    if (typeof node === "string") { if (node.trim()) fail(shape, "Unexpected text in a VML note shape."); }
    else if (!(node.uri === VML && ["fill", "stroke", "shadow", "path", "textbox"].includes(node.name) || node.uri === VML_EXCEL && node.name === "ClientData")) unsupported(node);
  }
  const insetMode = attr(shape, "insetmode", VML_OFFICE);
  if (insetMode !== undefined && !["auto", "custom"].includes(insetMode)) unsupported(shape);
  if (attr(shape, "allowincell", VML_OFFICE) !== undefined) vmlBoolean(attr(shape, "allowincell", VML_OFFICE), shape, false);
  const allowedStyle = new Set(["position", "margin-left", "margin-top", "left", "top", "width", "height", "z-index", "visibility", "mso-wrap-style", "v-text-anchor"]);
  for (const [name, content] of note.shapeStyle) {
    if (!allowedStyle.has(name)) unsupported(shape, "This VML note style is not supported.");
    if (["margin-left", "margin-top", "left", "top", "width", "height"].includes(name)) {
      const size = length(content, shape); if (["width", "height"].includes(name) && size < 0) fail(shape, "A VML note cache has a negative size.");
    }
  }
  if (note.shapeStyle.get("position") !== undefined && note.shapeStyle.get("position") !== "absolute") unsupported(shape);
  const wrap = note.shapeStyle.get("mso-wrap-style"); if (wrap !== undefined && !["none", "tight"].includes(wrap)) unsupported(shape);
  const zIndex = note.shapeStyle.get("z-index"); if (zIndex !== undefined && (!/^-?\d+$/.test(zIndex) || !Number.isSafeInteger(Number(zIndex)))) fail(shape, "A VML note stacking order is invalid.");
  const vertical = note.shapeStyle.get("v-text-anchor") ?? "top";
  if (!["top", "middle", "bottom"].includes(vertical)) unsupported(shape, "This VML note text direction is not supported.");
  const css: Css = { "--office-note-inset-mode": insetMode ?? "custom", "--office-note-vertical": note.textVertical ?? (vertical === "middle" ? "center" : vertical),
    "text-align": note.textHorizontal === "distributed" ? "justify" : note.textHorizontal ?? "left", "--office-note-horizontal": note.textHorizontal ?? "left",
    "--office-note-wrap-style": wrap ?? "none", ...(zIndex === undefined ? {} : { "--office-note-z-index": zIndex }) };
  const fill = one(shape, "fill", VML);
  if (fill) {
    vmlAttributes(fill, ["color", "color2", "type", "on", "opacity"], [[VML_OFFICE, "detectmouseclick"]]); check(fill, [], VML);
    if (value(fill, "type") !== undefined && value(fill, "type") !== "solid") unsupported(fill, "Only solid VML note fills are supported.");
    if (attr(fill, "detectmouseclick", VML_OFFICE) !== undefined) vmlBoolean(attr(fill, "detectmouseclick", VML_OFFICE), fill, false);
    if (value(fill, "color2") !== undefined) color(value(fill, "color2")!, fill); // Inactive in a solid fill.
  }
  css["--office-note-fill"] = vmlBoolean(fill ? value(fill, "on") : undefined, fill ?? shape, vmlBoolean(value(shape, "filled"), shape, true)) ? color(value(fill, "color") ?? value(shape, "fillcolor") ?? "white", shape) : "none";
  css["--office-note-fill-opacity"] = opacity(value(fill, "opacity"), fill ?? shape);
  const stroke = one(shape, "stroke", VML);
  if (stroke) {
    vmlAttributes(stroke, ["color", "weight", "on", "opacity", "joinstyle", "endcap", "dashstyle", "startarrow", "startarrowwidth", "startarrowlength", "endarrow", "endarrowwidth", "endarrowlength"]); check(stroke, [], VML);
    for (const [name, allowed, fallback] of [["joinstyle", ["round", "bevel", "miter"], "miter"], ["endcap", ["round", "square", "flat"], "flat"], ["dashstyle", ["solid", "shortdash", "shortdot", "shortdashdot", "shortdashdotdot", "dot", "dash", "longdash", "dashdot", "longdashdot", "longdashdotdot"], "solid"]] as const) {
      const selected = value(stroke, name) ?? fallback; if (!(allowed as readonly string[]).includes(selected)) unsupported(stroke); css[`--office-note-stroke-${name}`] = selected;
    }
    for (const end of ["start", "end"]) for (const suffix of ["arrow", "arrowwidth", "arrowlength"]) {
      const selected = value(stroke, end + suffix) ?? (suffix === "arrow" ? "none" : "medium"), allowed = suffix === "arrow" ? ["none", "block", "classic", "oval", "diamond", "open"] : suffix === "arrowwidth" ? ["narrow", "medium", "wide"] : ["short", "medium", "long"];
      if (!allowed.includes(selected)) unsupported(stroke); css[`--office-note-${end}-${suffix}`] = selected;
    }
  }
  css["--office-note-stroke"] = vmlBoolean(stroke ? value(stroke, "on") : undefined, stroke ?? shape, vmlBoolean(value(shape, "stroked"), shape, true)) ? color(value(stroke, "color") ?? value(shape, "strokecolor") ?? "black", shape) : "none";
  const weight = length(value(stroke, "weight") ?? value(shape, "strokeweight") ?? "1px", stroke ?? shape, true);
  if (weight < 0) fail(shape, "A VML note stroke has a negative width."); css["--office-note-stroke-width"] = `${weight}pt`; css["--office-note-stroke-opacity"] = opacity(value(stroke, "opacity"), stroke ?? shape);
  const shadow = one(shape, "shadow", VML);
  css["--office-note-shadow"] = "none";
  if (shadow) {
    vmlAttributes(shadow, ["on", "color", "obscured", "offset", "opacity", "type"]); check(shadow, [], VML);
    if (value(shadow, "type") !== undefined && value(shadow, "type") !== "single") unsupported(shadow, "Only a single VML note shadow is supported.");
    const offsets = (value(shadow, "offset") ?? "2pt,2pt").split(","); if (offsets.length !== 2) fail(shadow, "A VML note shadow requires two offsets.");
    css["--office-note-shadow"] = vmlBoolean(value(shadow, "on"), shadow, false) ? color(value(shadow, "color") ?? "#808080", shadow) : "none";
    css["--office-note-shadow-x"] = `${length(offsets[0]!, shadow)}pt`; css["--office-note-shadow-y"] = `${length(offsets[1]!, shadow)}pt`;
    css["--office-note-shadow-opacity"] = opacity(value(shadow, "opacity"), shadow); css["--office-note-shadow-obscured"] = String(vmlBoolean(value(shadow, "obscured"), shadow, false));
  }
  path(one(shape, "path", VML), false);
  const textbox = one(shape, "textbox", VML), insets = ["7.2pt", "3.6pt", "7.2pt", "3.6pt"];
  if (textbox) {
    vmlAttributes(textbox, ["style", "inset"]);
    for (const [name, content] of vmlStyle(textbox)) if (name !== "mso-direction-alt" || content !== "auto") unsupported(textbox, "This VML textbox direction or automatic sizing is not supported.");
    for (const node of textbox.content) {
      if (typeof node === "string") { if (node.trim()) unsupported(textbox, "Note text must belong to its classic comment part."); continue; }
      if (node.uri !== "" || node.name !== "div") unsupported(node);
      vmlAttributes(node, ["style"]); check(node, [], "");
      for (const [name, content] of vmlStyle(node)) {
        if (name !== "text-align" || !["left", "center", "right", "justify"].includes(content)) unsupported(node);
        if (note.textHorizontal === undefined) { css["text-align"] = content; css["--office-note-horizontal"] = content; }
      }
    }
    const source = value(textbox, "inset");
    if (source !== undefined) {
      const entries = source.split(/[,\s]+/).filter(Boolean); if (entries.length > 4) fail(textbox, "A VML note textbox has too many inset values.");
      for (const [index, entry] of entries.entries()) { const points = length(entry, textbox); if (points < 0) unsupported(textbox, "Negative VML note insets are not supported."); insets[index] = `${points}pt`; }
    }
  }
  for (const [index, side] of ["left", "top", "right", "bottom"].entries()) css[`padding-${side}`] = insets[index]!;
  return css;
}
