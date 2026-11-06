import { DEFAULT_WORKBOOK_INDEXED_COLORS } from "@doffi/core";
import { check, elements, fail, S, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { attributes, boolean, integer, number, one, value } from "./values.js";

export const themeSlots: readonly string[] = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
export interface SpreadsheetColor { readonly css: string; readonly tint: number; readonly nativeAlpha?: number }
function argb(source: string | undefined, node: XmlNode): { css: string; nativeAlpha: number } {
  if (!source || !/^[0-9a-f]{8}$/i.test(source)) fail(node, "A SpreadsheetML ARGB color is invalid.");
  // Cell formatting uses RGB; the reserved/native alpha byte is retained for
  // export, rather than making default 00-prefixed indexed colors invisible.
  return { css: "#" + source.slice(2).toLowerCase(), nativeAlpha: parseInt(source.slice(0, 2), 16) };
}
export class SpreadsheetColors {
  readonly css: Css = {};
  private readonly indexed: readonly string[];
  constructor(colors: XmlNode | undefined, private readonly theme: ReadonlySet<string>) {
    if (colors) { attributes(colors, []); check(colors, ["indexedColors", "mruColors"], S); }
    const indexed = one(colors, "indexedColors");
    if (indexed) {
      attributes(indexed, []); check(indexed, ["rgbColor"], S);
      this.indexed = elements(indexed, "rgbColor", S).map((node, index) => {
        attributes(node, ["rgb"]); check(node, [], S); const parsed = argb(value(node, "rgb"), node);
        this.css[`--office-indexed-${index}`] = parsed.css; this.css[`--office-indexed-${index}-native-alpha`] = String(parsed.nativeAlpha);
        return parsed.css;
      });
      if (!this.indexed.length || this.indexed.length > 64) fail(indexed, "An indexed color palette has an invalid size.");
    } else {
      this.indexed = DEFAULT_WORKBOOK_INDEXED_COLORS;
      this.indexed.forEach((color, index) => { this.css[`--office-indexed-${index}`] = color; this.css[`--office-indexed-${index}-native-alpha`] = "0"; });
    }
    const recent = one(colors, "mruColors");
    if (recent) { attributes(recent, []); check(recent, ["color"], S); for (const node of elements(recent, "color", S)) this.read(node); }
    this.css["--office-system-foreground"] = "#000000"; this.css["--office-system-background"] = "#ffffff";
  }
  read(node: XmlNode): SpreadsheetColor {
    attributes(node, ["auto", "indexed", "rgb", "theme", "tint"]); check(node, [], S);
    const tint = number(node, "tint", 0); if (tint < -1 || tint > 1) fail(node, "A SpreadsheetML tint is outside its range.");
    // Office's documented precedence is theme > RGB > indexed > automatic.
    // Validate every supplied attribute, including lower-priority values.
    const automatic = value(node, "auto") === undefined ? undefined : boolean(node, "auto");
    const index = value(node, "indexed") === undefined ? undefined : integer(node, "indexed", 0, 255);
    const rgb = value(node, "rgb") === undefined ? undefined : argb(value(node, "rgb"), node);
    const theme = value(node, "theme") === undefined ? undefined : integer(node, "theme", 0, 255);
    if (theme !== undefined) {
      const role = themeSlots[theme]; if (!role || !this.theme.has(role)) unsupported(node, "A SpreadsheetML theme color cannot be resolved.");
      return { css: `var(--office-theme-${role})`, tint };
    }
    if (rgb) return { ...rgb, tint };
    if (index !== undefined) {
      if (index === 64 || index === 65) return { css: `var(--office-system-${index === 64 ? "foreground" : "background"})`, tint };
      if (!this.indexed[index]) fail(node, "An indexed color cannot be resolved.");
      return { css: `var(--office-indexed-${index})`, tint };
    }
    if (automatic) return { css: "var(--office-system-foreground)", tint };
    return fail(node, "A SpreadsheetML color has no effective source.");
  }
}
export function colorDeclarations(prefix: string, color: SpreadsheetColor): Css {
  return { [prefix + "-color"]: color.css, [prefix + "-tint"]: String(color.tint), [prefix + "-native-alpha"]: color.nativeAlpha === undefined ? "initial" : String(color.nativeAlpha) };
}
