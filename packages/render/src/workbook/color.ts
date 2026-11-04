import { DocumentError } from "@dagent-office/core";

export type WorkbookRgb = readonly [number, number, number];

/** SpreadsheetML tint changes HLS luminance, unlike DrawingML RGB tint/shade.
 * Work in normalized HLS and round only the final 8-bit display channels. */
export function tintWorkbookColor(rgb: WorkbookRgb, tint: number): WorkbookRgb {
  if (rgb.length !== 3 || rgb.some(value => !Number.isInteger(value) || value < 0 || value > 255)
    || !Number.isFinite(tint) || tint < -1 || tint > 1) throw new DocumentError("InvalidDocument", "A worksheet color or luminance tint is invalid.");
  if (tint === 0) return Object.freeze([...rgb]) as WorkbookRgb;
  const [r, g, b] = rgb.map(value => value / 255) as [number, number, number];
  const max = Math.max(r, g, b), min = Math.min(r, g, b), delta = max - min, lightness = (max + min) / 2;
  const luminance = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  const saturation = delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  let hue = delta === 0 ? 0 : max === r ? (g - b) / delta : max === g ? (b - r) / delta + 2 : (r - g) / delta + 4;
  hue = (hue + 6) % 6;
  const chroma = (1 - Math.abs(2 * luminance - 1)) * saturation, x = chroma * (1 - Math.abs(hue % 2 - 1)), m = luminance - chroma / 2;
  const channels = hue < 1 ? [chroma, x, 0] : hue < 2 ? [x, chroma, 0] : hue < 3 ? [0, chroma, x] : hue < 4 ? [0, x, chroma] : hue < 5 ? [x, 0, chroma] : [chroma, 0, x];
  return Object.freeze(channels.map(channel => Math.round(Math.max(0, Math.min(1, channel + m)) * 255))) as unknown as WorkbookRgb;
}

const contexts = new WeakMap<Document, CanvasRenderingContext2D>();
/** Resolve already-computed source CSS. The native ARGB high byte is retained
 * for export in CSS but is not an opacity for SpreadsheetML cell formatting. */
export function resolvedWorkbookColor(document: Document, style: CSSStyleDeclaration, prefix: string): string | undefined {
  let source = style.getPropertyValue(prefix + "-color").trim(); if (!source) return undefined;
  if (source.toLowerCase() === "currentcolor") source = style.color;
  const rawTint = style.getPropertyValue(prefix + "-tint").trim(), tint = rawTint ? Number(rawTint) : 0;
  return workbookDisplayColor(document, source, tint);
}

/** Normalize a computed CSS color for first-party paint layers. The caller
 * resolves contextual values such as currentColor before crossing this edge. */
export function workbookDisplayColor(document: Document, source: string, tint = 0): string {
  let context = contexts.get(document);
  if (!context) {
    const created = document.createElement("canvas").getContext("2d");
    if (!created) throw new DocumentError("UnsupportedFeature", "This browser cannot resolve worksheet colors.");
    context = created; contexts.set(document, created);
  }
  // Invalid Canvas colors leave the previous value intact. Two different
  // sentinels distinguish a rejected value from a valid sentinel-colored input.
  context.fillStyle = "#010203"; context.fillStyle = source; const first = context.fillStyle;
  context.fillStyle = "#040506"; context.fillStyle = source;
  if (first !== context.fillStyle) throw new DocumentError("InvalidDocument", "A worksheet source color cannot be resolved.");
  const hex = /^#([0-9a-f]{6})([0-9a-f]{2})?$/i.exec(first), rgba = /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)(?:\s*,\s*([\d.]+))?\s*\)$/i.exec(first);
  if (!hex && !rgba) throw new DocumentError("UnsupportedFeature", "A worksheet color requires an sRGB display value.");
  const rgb = hex ? [0, 2, 4].map(offset => parseInt(hex[1]!.slice(offset, offset + 2), 16)) : rgba!.slice(1, 4).map(Number);
  const alpha = hex ? hex[2] ? parseInt(hex[2], 16) / 255 : 1 : Number(rgba![4] ?? 1);
  return `rgba(${tintWorkbookColor(rgb as unknown as WorkbookRgb, tint).join(",")},${alpha})`;
}
