import type { WorkbookDisplayFormat, WorkbookNumberFormatter } from "@doffi/core";
import type { WorkbookCellTextInput } from "./cell-text-layout.js";
import { workbookTextAlignment } from "./text-alignment.js";
import { officeFontSegments } from "../font-scripts.js";

export interface WorkbookRowTypography { readonly eligible: boolean; readonly signature: readonly string[]; readonly cacheId?: number }
const properties = ["font-family", "font-size", "font-style", "font-weight", "font-variant", "line-height", "vertical-align", "text-transform",
  "letter-spacing", "word-spacing", "--office-font-ascii", "--office-minor-ascii", "--office-font-h-ansi", "--office-minor-h-ansi",
  "--office-font-east-asia", "--office-minor-east-asia", "--office-font-complex-script", "--office-minor-complex-script", "--office-font-underline", "--office-font-strike"];

function rowTypography(input: WorkbookCellTextInput, typography?: Map<string, WorkbookRowTypography>): WorkbookRowTypography {
  const styleKey = input.source.format.cssClass + "\n" + input.source.element.style.cssText;
  let resolved = typography?.get(styleKey);
  if (!resolved) {
    const style = input.source.element.ownerDocument.defaultView!.getComputedStyle(input.source.element), alignment = workbookTextAlignment(style, input.value?.type ?? "blank", "ltr");
    resolved = { eligible: !alignment.wrap && !alignment.rotationDegrees && !alignment.stacked && !alignment.shrink, signature: properties.map(property => style.getPropertyValue(property)) };
    if (typography && typography.size < 1024 && styleKey.length + resolved.signature.reduce((sum, part) => sum + part.length, 0) <= 4096) {
      resolved = { ...resolved, cacheId: typography.size }; typography.set(styleKey, resolved);
    }
  }
  return resolved;
}
export function workbookRowFontKey(input: WorkbookCellTextInput, typography?: Map<string, WorkbookRowTypography>): string {
  const resolved = rowTypography(input, typography); return JSON.stringify(resolved.cacheId ?? resolved.signature);
}

/** Sharing is limited to unwrapped, unrotated, plain displays with the exact
 * same computed typography, script sequence and glyphs. Non-ASCII segments
 * preserve their complete text (including variation sequences and emoji).
 * Rich runs, newlines and geometry-dependent layouts get their own measurement.
 * Coordinate-sensitive CSS is resolved before a cache key is constructed. */
export function workbookRowMetricKey(input: WorkbookCellTextInput, formatter: WorkbookNumberFormatter, formats: ReadonlyMap<string, WorkbookDisplayFormat>, typography?: Map<string, WorkbookRowTypography>): string | undefined {
  if (!input.value || input.value.type === "blank" || input.source.element.children.length) return undefined;
  const resolved = rowTypography(input, typography);
  if (!resolved.eligible) return undefined;
  const value = input.value;
  let text: string;
  if (value.type === "number" || value.type === "text") {
    const display = formatter.format(value.value, input.source.format.numberFormatId ? formats.get(input.source.format.numberFormatId) : undefined);
    text = display.kind === "parts" ? display.parts.map(part => part.kind === "text" ? part.text : part.character).join("") : "#";
  } else text = value.type === "boolean" ? value.value ? "TRUE" : "FALSE" : value.value;
  if (/[\r\n\t\u0000-\u001f]/.test(text)) return undefined;
  const segments = officeFontSegments(text).map(segment => [segment.role, segment.role === "ascii" ? [...new Set(segment.text)].sort().join("") : segment.text]);
  // The pass-local typography map already owns the exact signature. Refer to
  // that immutable entry instead of serializing its twenty properties again
  // for every source value; uncached/coordinate-sensitive styles keep the full
  // signature and non-ASCII glyph sequences remain complete.
  return JSON.stringify([segments, value.type === "number" ? input.widthPt : 0, resolved.cacheId ?? resolved.signature]);
}
