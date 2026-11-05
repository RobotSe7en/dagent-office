import { DocumentError } from "@dagent-office/core";
import { builtinTableDifferentials, builtinTableElements } from "./table-style-data.js";
import type { WorkbookBuiltinTableDifferential, WorkbookTableRegion } from "./table-style-types.js";

export interface WorkbookBuiltinTableStyle {
  readonly name: string;
  /** Low to high precedence, from ECMA-376 Part 1 §18.8.41. */
  readonly elements: readonly { readonly region: WorkbookTableRegion; readonly size: number; readonly style: WorkbookBuiltinTableDifferential }[];
}
// These are static standard definitions, never document-dependent cache entries.
function freezeStandard<Value>(value: Value): Value {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    for (const child of Object.values(value)) freezeStandard(child);
    Object.freeze(value);
  }
  return value;
}
freezeStandard(builtinTableDifferentials); freezeStandard(builtinTableElements);
const styles = new Map<string, WorkbookBuiltinTableStyle>();
for (const [name, elements] of Object.entries(builtinTableElements)) styles.set(name, freezeStandard({ name,
  elements: elements.map(element => ({ region: element.region, size: element.size, style: builtinTableDifferentials[element.differential]! })) }));

export function workbookBuiltinTableStyle(name: string): WorkbookBuiltinTableStyle {
  const style = styles.get(name);
  if (!style) throw new DocumentError("UnsupportedFeature", "This predefined workbook table style is not supported.");
  return style;
}
