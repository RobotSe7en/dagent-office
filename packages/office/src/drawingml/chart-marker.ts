import type { ChartAppearance } from "@dagent-office/core";
import { C, check, child, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { appearance, decimal, val } from "./chart-style.js";

/** Missing point properties inherit the series; never invent a source size or
 * symbol on a partial override. Native CT_MarkerSize is integral, 2–72 pt. */
export function chartMarker(source: XmlNode | undefined): {
  marker?: NonNullable<ChartAppearance["seriesStyles"]>[number]["marker"];
  markerSize?: number; css: Css;
} {
  if (!source) return { css: {} };
  if (source.attrs.size) unsupported(source, "A chart marker has unknown attributes.");
  check(source, ["symbol", "size", "spPr"], C);
  for (const name of ["symbol", "size", "spPr"]) if (elements(source, name, C).length > 1) fail(source, "A chart marker property is repeated.");
  for (const name of ["symbol", "size"]) {
    const leaf = child(source, name, C); if (!leaf) continue;
    if ([...leaf.attrs.keys()].some(key => key !== " val")) unsupported(leaf, "A chart marker property has unknown attributes.");
    check(leaf, [], C);
  }
  const marker = val(source, "symbol"), size = child(source, "size", C);
  const markerSize = size && val(source, "size") === undefined ? 5 : decimal(source, "size");
  if (child(source, "symbol", C) && marker === undefined) fail(source, "A chart marker symbol is missing its value.");
  if (marker !== undefined && !["none", "circle", "square", "diamond", "triangle", "x", "star", "plus", "dot", "dash"].includes(marker)) unsupported(source, "This chart marker shape is not supported.");
  if (markerSize !== undefined && (!Number.isInteger(markerSize) || markerSize < 2 || markerSize > 72)) fail(source, "A chart marker size exceeds native bounds.");
  const css: Css = {};
  for (const [key, value] of Object.entries(appearance(child(source, "spPr", C)))) css[key.replace("--office-", "--office-marker-")] = value;
  return { ...(marker === undefined ? {} : { marker: marker as NonNullable<ChartAppearance["seriesStyles"]>[number]["marker"] & string }), ...(markerSize === undefined ? {} : { markerSize }), css };
}
