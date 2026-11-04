import type { Css } from "../ooxml/styles.js";
import { A, check, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { OpcPackage } from "../ooxml/package.js";
import { cssFontName } from "../ooxml/css-values.js";
import { attributes, one, value } from "./values.js";
import { themeSlots } from "./colors.js";

export interface WorkbookTheme {
  readonly css: Css;
  readonly colors: ReadonlySet<string>;
  readonly source?: XmlNode;
}
/** Common cell roles are CSS definitions. Drawing format lists remain in the
 * transient import tree for the drawing adapter, never in the saved package. */
export function readWorkbookTheme(opc: OpcPackage, workbookPart: string): WorkbookTheme {
  const part = opc.related(workbookPart, "theme"); if (!part) return { css: {}, colors: new Set() };
  if (opc.contentType(part) !== "application/vnd.openxmlformats-officedocument.theme+xml") fail({ part, name: "theme" }, "The theme content type is invalid.");
  const root = opc.read(part);
  if (root.uri !== A || root.name !== "theme") fail(root, "The theme root is invalid.");
  attributes(root, ["name"]); check(root, ["themeElements", "objectDefaults", "extraClrSchemeLst", "custClrLst", "extLst"], A);
  const theme = one(root, "themeElements", A), colors = one(theme, "clrScheme", A), fonts = one(theme, "fontScheme", A);
  if (!theme || !colors || !fonts || !one(theme, "fmtScheme", A)) fail(root, "A native theme is incomplete.");
  attributes(theme, []); check(theme, ["clrScheme", "fontScheme", "fmtScheme"], A);
  attributes(colors, ["name"]); check(colors, themeSlots, A);
  const css: Css = {};
  for (const role of themeSlots) {
    const container = one(colors, role, A); if (!container) fail(colors, "A native theme color is missing.");
    attributes(container, []); check(container, ["srgbClr", "sysClr"], A);
    const values = elements(container, undefined, A); if (values.length !== 1) fail(container, "A theme color requires one definition.");
    const source = values[0]!; check(source, [], A);
    attributes(source, source.name === "sysClr" ? ["val", "lastClr"] : ["val"]);
    const rgb = value(source, source.name === "sysClr" ? "lastClr" : "val");
    if (!rgb || !/^[0-9a-f]{6}$/i.test(rgb)) fail(source, "A theme base color must be an untransformed RGB value.");
    css[`--office-theme-${role}`] = "#" + rgb.toLowerCase();
  }
  for (const [role, base] of [["bg1", "lt1"], ["bg2", "lt2"], ["tx1", "dk1"], ["tx2", "dk2"]]) css[`--office-theme-${role}`] = `var(--office-theme-${base})`;
  attributes(fonts, ["name"]); check(fonts, ["majorFont", "minorFont"], A);
  for (const kind of ["major", "minor"] as const) {
    const source = one(fonts, kind + "Font", A); if (!source) fail(fonts, "A theme font family is missing.");
    attributes(source, []); check(source, ["latin", "ea", "cs", "font"], A);
    const scripts = new Set<string>();
    for (const item of elements(source, "font", A)) {
      attributes(item, ["script", "typeface"]); check(item, [], A);
      const script = value(item, "script"), name = value(item, "typeface");
      if (!script || !/^[A-Za-z]{4}$/.test(script) || scripts.has(script) || !name) fail(item, "A supplemental theme font is invalid or duplicated.");
      scripts.add(script); css[`--office-${kind}-script-${script}`] = cssFontName(name);
    }
    for (const [tag, role, fallbackScript] of [["latin", "ascii", undefined], ["ea", "east-asia", "Hans"], ["cs", "complex-script", "Arab"]] as const) {
      const item = one(source, tag, A); if (!item) fail(source, "A theme font role is missing.");
      attributes(item, ["typeface", "panose", "pitchFamily", "charset"]); check(item, [], A);
      const name = value(item, "typeface"); if (name === undefined) fail(item, "A theme font role has no typeface attribute.");
      if (tag === "latin" && !name) unsupported(item, "An empty Latin theme font cannot be resolved.");
      css[`--office-${kind}-${role}`] = name ? cssFontName(name) : scripts.has(fallbackScript!) ? `var(--office-${kind}-script-${fallbackScript})` : `var(--office-${kind}-ascii)`;
    }
    css[`--office-${kind}-h-ansi`] = `var(--office-${kind}-ascii)`;
  }
  return { css, colors: new Set(themeSlots), source: root };
}
