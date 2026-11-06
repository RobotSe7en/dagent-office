import { DocumentError } from "@doffi/core";
import type { WorkbookNumberDisplayPart } from "@doffi/core";

export type MeasuredNumberPart = { readonly kind: "text"; readonly text: string; readonly width: number } | { readonly kind: "space"; readonly width: number };
export interface MeasuredNumberDisplay { readonly parts: readonly MeasuredNumberPart[]; readonly width: number; readonly overflow: boolean }

/** Caller supplies one consistent font measurement and unit system. Spacers are
 * measured from their actual source character; fill is measured as a complete
 * shaped string, never estimated by repeating a single-glyph advance. */
export function measureNumberDisplay(source: readonly WorkbookNumberDisplayPart[], availableWidth: number, measure: (text: string) => number): MeasuredNumberDisplay {
  if (!Number.isFinite(availableWidth) || availableWidth < 0) throw new DocumentError("InvalidOperation", "Number display width must be finite and nonnegative.");
  if (source.length > 32767) throw new DocumentError("ResourceLimitExceeded", "Number display has too many parts.");
  const measured = (text: string): number => {
    const width = measure(text);
    if (!Number.isFinite(width) || width < 0) throw new DocumentError("InvalidOperation", "Number display measurements must be finite and nonnegative.");
    return width;
  };
  const normalized: WorkbookNumberDisplayPart[] = []; let units = 0, lastFill = -1;
  for (let index = 0; index < source.length; index++) if (source[index]!.kind === "fill") lastFill = index;
  for (const [index, part] of source.entries()) {
    if (part.kind === "fill") {
      if ([...part.character].length !== 1) throw new DocumentError("InvalidOperation", "A numeric fill field needs one Unicode character.");
      // ECMA-376 §18.8.31 uses only the final repetition marker in a section.
      if (index !== lastFill) continue;
    }
    units += part.kind === "text" ? part.text.length : part.kind === "space" ? part.character.length : 0;
    if (units > 32767) throw new DocumentError("ResourceLimitExceeded", "Number display exceeds its text budget.");
    const previous = normalized.at(-1);
    if (previous?.kind === "text" && part.kind === "text") normalized[normalized.length - 1] = { kind: "text", text: previous.text + part.text };
    else normalized.push(part);
  }
  const parts: MeasuredNumberPart[] = []; let fixedWidth = 0, fill: { index: number; character: string } | undefined;
  for (const part of normalized) {
    if (part.kind === "fill") {
      fill = { index: parts.length, character: part.character };
      parts.push({ kind: "text", text: "", width: 0 }); continue;
    }
    const text = part.kind === "text" ? part.text : part.character;
    const width = measured(text); fixedWidth += width;
    if (!Number.isFinite(fixedWidth)) throw new DocumentError("InvalidOperation", "Number display measurements exceed their numeric range.");
    parts.push(part.kind === "text" ? { kind: "text", text, width } : { kind: "space", width });
  }
  let width = fixedWidth;
  if (fill) {
    const result = fitWorkbookFill(fill.character, Math.max(0, availableWidth - fixedWidth), units, measured);
    parts[fill.index] = { kind: "text", ...result }; width += result.width;
  }
  return { parts, width, overflow: fixedWidth > availableWidth };
}

/** Shared measured repetition for numeric and source-aware text sections. The
 * caller accounts for all fixed displayed code units before admitting a fill. */
export function fitWorkbookFill(character: string, availableWidth: number, fixedCodeUnits: number, measure: (text: string) => number): { readonly text: string; readonly width: number } {
  if ([...character].length !== 1 || !Number.isFinite(availableWidth) || availableWidth < 0 || !Number.isSafeInteger(fixedCodeUnits) || fixedCodeUnits < 0 || fixedCodeUnits > 32767)
    throw new DocumentError("InvalidOperation", "Text repetition needs one Unicode character, finite space and a bounded fixed text size.");
  const measured = (text: string): number => {
    const width = measure(text);
    if (!Number.isFinite(width) || width < 0) throw new DocumentError("InvalidOperation", "Text repetition measurements must be finite and nonnegative.");
    return width;
  };
  if (!availableWidth || measured(character) === 0) return { text: "", width: 0 };
  const maximum = Math.floor((32767 - fixedCodeUnits) / character.length);
  // Exponential admission keeps a normal cell from shaping thousands of
  // characters. One extra character distinguishes an exact fit from truncation.
  let low = 0, high = 1;
  while (high <= maximum && measured(character.repeat(high)) <= availableWidth) {
    low = high; high = Math.min(maximum + 1, high * 2);
  }
  while (low < high) {
    const middle = Math.ceil((low + high) / 2), sample = character.repeat(middle);
    if (measured(sample) <= availableWidth) low = middle; else high = middle - 1;
  }
  if (low > maximum) throw new DocumentError("ResourceLimitExceeded", "Text repetition exceeds the rendered text budget.");
  const text = character.repeat(low); return { text, width: measured(text) };
}
