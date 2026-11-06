import { checkLimit, invalid, unsupported } from "../document/validation.js";
import { formatGeneralNumber } from "./format-general.js";
import type { NumberDisplayPart } from "./format-numeric.js";
import type { NumberFormatSection } from "./number-format.js";

export interface WorkbookNumberMeasure { readonly availableWidth: number; readonly measure: (text: string) => number }
/** General fields are fitted as a whole section, including literal affixes and
 * measured spacers. Fill is expanded later by the render package. */
export function formatGeneralSection(value: number, section: NumberFormatSection, automaticNegative: boolean,
  decimal: string, measurement?: WorkbookNumberMeasure): readonly NumberDisplayPart[] | undefined {
  if (measurement && (!Number.isFinite(measurement.availableWidth) || measurement.availableWidth < 0 || typeof measurement.measure !== "function")) invalid("General display requires a finite width and text measurement function.");
  if (!section.tokens.some(token => token.kind === "general") || section.tokens.some(token => !["general", "literal", "percent", "space", "fill"].includes(token.kind))) unsupported("This General section contains unsupported mixed numeric or calendar fields.");
  const assemble = (number: string): readonly NumberDisplayPart[] => {
    const parts: NumberDisplayPart[] = []; let units = 0;
    const text = (text: string): void => {
      units += text.length; checkLimit(units, 32767); if (!text) return;
      const previous = parts.at(-1);
      if (previous?.kind === "text") parts[parts.length - 1] = { kind: "text", text: previous.text + text };
      else parts.push({ kind: "text", text });
    };
    if (automaticNegative && value < 0) text("-");
    for (const token of section.tokens) {
      if (token.kind === "literal") text(token.text);
      else if (token.kind === "general") text(number);
      else if (token.kind === "percent") text("%");
      else if (token.kind === "space" || token.kind === "fill") { units += token.character.length; checkLimit(units, 32767); parts.push(token); }
    }
    return parts;
  };
  const measured = (number: string): number => {
    if (!measurement) return 0;
    let width = 0;
    for (const part of assemble(number)) if (part.kind !== "fill") {
      const value = measurement.measure(part.kind === "text" ? part.text : part.character);
      if (!Number.isFinite(value) || value < 0) invalid("General display measurements must be finite and nonnegative.");
      width += value;
    }
    if (!Number.isFinite(width)) invalid("General display measurement exceeds its numeric range.");
    return width;
  };
  const result = formatGeneralNumber(Math.abs(value), decimal, measurement?.availableWidth ?? 0, measured,
    section.tokens.some(token => token.kind === "percent") ? 2 : 0);
  return result.kind === "overflow" ? undefined : assemble(result.text);
}
