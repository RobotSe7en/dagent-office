import { invalid, unsupported } from "../document/validation.js";
import type { WorkbookDisplayLocale } from "./display-locale.js";
import { workbookLocaleDatePattern } from "./locale-date-pattern.js";

// ECMA-376 §18.8.30, with the documented Excel corrections in MS-OI29500
// §2.1.712 (padding in 37–40 and the time separator in 47). Definitions not
// assigned by the standard require an explicit source code; do not invent a
// currency for IDs such as 5–8 or silently substitute General for unknown IDs.
const common: Readonly<Record<number, string>> = Object.freeze({
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00",
  12: "# ?/?", 13: "# ??/??", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM",
  20: "h:mm", 21: "h:mm:ss", 37: "#,##0_);(#,##0)", 38: "#,##0_);[Red](#,##0)",
  39: "#,##0.00_);(#,##0.00)", 40: "#,##0.00_);[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
});
const simplifiedChinese: Readonly<Record<number, string>> = Object.freeze({
  27: 'yyyy"年"m"月"', 28: 'm"月"d"日"', 29: 'm"月"d"日"', 30: "m-d-yy", 31: 'yyyy"年"m"月"d"日"',
  32: 'h"时"mm"分"', 33: 'h"时"mm"分"ss"秒"', 34: 'AM/PM h"时"mm"分"', 35: 'AM/PM h"时"mm"分"ss"秒"',
  36: 'yyyy"年"m"月"', 50: 'yyyy"年"m"月"', 51: 'm"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"',
  54: 'm"月"d"日"', 55: 'AM/PM h"时"mm"分"', 56: 'AM/PM h"时"mm"分"ss"秒"', 57: 'yyyy"年"m"月"', 58: 'm"月"d"日"',
});

export function workbookBuiltinNumberFormat(id: number, locale: WorkbookDisplayLocale): string {
  if (!Number.isSafeInteger(id) || id < 0 || id > 163) invalid("A workbook built-in number format must lie in 0–163.");
  if (common[id] !== undefined) return common[id]!;
  if (id === 14 || id === 22) return workbookLocaleDatePattern(locale, "shortDate") + (id === 22 ? " h:mm" : "");
  const language = new Intl.Locale(locale.locale).maximize();
  if (language.language === "zh" && language.script === "Hans" && simplifiedChinese[id] !== undefined) return simplifiedChinese[id]!;
  unsupported("This built-in format needs an explicit source definition or a regional calendar adapter.");
}
