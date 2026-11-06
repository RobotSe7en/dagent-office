import type { Css } from "../ooxml/styles.js";
import { A, check, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { effects, fill, outline } from "../drawingml/styles.js";
import { attributes, one } from "./values.js";
import { validateDrawingEffects, validateDrawingFill, validateDrawingLine } from "./drawing-style.js";

export interface WorkbookDrawingTheme {
  readonly fill: readonly string[];
  readonly background: readonly string[];
  readonly line: readonly string[];
  readonly effect: readonly string[];
}
/** Consume the complete native style matrix. Definitions remain shared CSS
 * classes; per-object style colors are supplied by reference declarations. */
export function readDrawingTheme(source: XmlNode | undefined, rule: (name: string, css: Css) => void): WorkbookDrawingTheme {
  const result: { [K in keyof WorkbookDrawingTheme]: string[] } = { fill: [], background: [], line: [], effect: [] };
  if (!source) return result;
  const theme = one(source, "themeElements", A), format = one(theme, "fmtScheme", A);
  if (!format) fail(source, "A drawing theme format scheme is missing.");
  attributes(format, ["name"]); check(format, ["fillStyleLst", "bgFillStyleLst", "lnStyleLst", "effectStyleLst"], A);
  const fillStyle = (node: XmlNode): Css => { validateDrawingFill(node); return fill({ ...node, content: [node] }); };
  const lineStyle = (node: XmlNode): Css => { validateDrawingLine(node); return outline(node); };
  const effectStyle = (node: XmlNode): Css => {
    attributes(node, []); check(node, ["effectLst"], A);
    const list = one(node, "effectLst", A); if (!list) fail(node, "A theme effect style has no effect list.");
    validateDrawingEffects(list); return effects(list);
  };
  for (const [name, kind, children, map] of [
    ["fillStyleLst", "fill", ["noFill", "solidFill", "gradFill"], fillStyle],
    ["bgFillStyleLst", "background", ["noFill", "solidFill", "gradFill"], fillStyle],
    ["lnStyleLst", "line", ["ln"], lineStyle], ["effectStyleLst", "effect", ["effectStyle"], effectStyle],
  ] as const) {
    const list = one(format, name, A); if (!list) fail(format, "A theme format list is missing.");
    attributes(list, []); check(list, children, A);
    const items = elements(list, undefined, A); if (items.length < 3) fail(list, "A native theme format list requires at least three entries.");
    for (const [index, item] of items.entries()) {
      const css = map(item), prefix = `--office-${kind}-reference`;
      for (const key of Object.keys(css)) {
        css[key] = css[key]!.replaceAll("--office-placeholder-color", prefix + "-color");
        if (key.endsWith("-source") && css[key] === "placeholder") { css[key] = "reference"; css[key.slice(0, -7) + "-reference-prefix"] = prefix; }
      }
      const name = `xlsx-theme-${kind}-${index + 1}`; rule(name, css); result[kind].push(name);
    }
  }
  for (const name of ["objectDefaults", "extraClrSchemeLst", "custClrLst", "extLst"]) {
    const node = one(source, name, A); if (!node) continue;
    attributes(node, []);
    if (node.content.some(item => typeof item !== "string" || item.trim())) unsupported(node, "Nonempty supplementary drawing theme definitions are not yet supported.");
  }
  return result;
}
