import { formatDecimalSection } from "./format-numeric.js";
import type { NumberDisplayPart, NumberSeparators } from "./format-numeric.js";
import { formatScientificSection } from "./format-scientific.js";
import { formatFractionSection } from "./format-fraction.js";
import { formatLiteralSection, selectNumberFormat } from "./format-select.js";
import type { NumberFormatColor, NumberFormatSection } from "./number-format.js";
import { checkLimit } from "../document/validation.js";

export type WorkbookValueDisplay =
  | { readonly kind: "parts"; readonly parts: readonly NumberDisplayPart[]; readonly color?: NumberFormatColor; readonly locale?: number }
  | { readonly kind: "general" | "date"; readonly value: number; readonly section: NumberFormatSection; readonly automaticNegative: boolean }
  | { readonly kind: "unmatchedCondition" };

/** Internal scalar dispatch. General width fitting and calendar rendering have
 * distinct consumers; returning their source fields is not a formatted value
 * or a claim that a native cached string has been recomputed. */
export function formatWorkbookValue(value: number | string, sections: readonly NumberFormatSection[], separators: NumberSeparators): WorkbookValueDisplay {
  if (typeof value === "string") checkLimit(value.length, 32767);
  const selection = selectNumberFormat(value, sections);
  if (selection.kind === "unmatchedCondition") return selection;
  if (selection.kind === "plainText") return { kind: "parts", parts: [{ kind: "text", text: value as string }] };
  if (selection.kind === "general") return { kind: "general", value: value as number, section: { tokens: [{ kind: "general" }] }, automaticNegative: (value as number) < 0 };
  const section = selection.section, metadata = { ...(section.color ? { color: section.color } : {}), ...(section.locale === undefined ? {} : { locale: section.locale }) };
  if (selection.kind === "text") return { kind: "parts", parts: formatLiteralSection(value as string, section), ...metadata };
  const scalar = value as number;
  if (section.tokens.some(token => ["date", "elapsed", "dayPeriod"].includes(token.kind))) return { kind: "date", value: scalar, section, automaticNegative: selection.automaticNegative };
  if (section.tokens.some(token => token.kind === "general")) return { kind: "general", value: scalar, section, automaticNegative: selection.automaticNegative };
  if (!section.tokens.some(token => token.kind === "digit")) return { kind: "parts", parts: formatLiteralSection(undefined, section), ...metadata };
  const display = section.tokens.some(token => token.kind === "exponent") ? formatScientificSection(scalar, section, separators)
    : section.tokens.some(token => token.kind === "slash") ? formatFractionSection(scalar, section) : formatDecimalSection(scalar, section, separators);
  return { kind: "parts", parts: selection.automaticNegative && !display.zero ? [{ kind: "text", text: "-" }, ...display.parts] : display.parts, ...metadata };
}
