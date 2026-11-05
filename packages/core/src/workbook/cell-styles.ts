import type { WorkbookCellStyles } from "../contracts/workbook.js";
import { checkLimit, invalid } from "../document/validation.js";

export function validateWorkbookCellStyles(styles: WorkbookCellStyles | undefined, classExists: (name: string) => void, formatExists: (id: string) => void, maxItems: number): { classes: Set<string>; items: number } {
  const classes = new Set<string>();
  if (!styles) return { classes, items: 0 };
  const items = styles.definitions.length + styles.names.length + styles.bindings.length; checkLimit(items, maxItems);
  const definitions = new Set<string>(), names = new Set<string>(); let normal = false;
  const styleExists = (id: string): void => { if (!definitions.has(id)) invalid("A named cell style refers to an undefined definition."); };
  const bindClass = (name: string): void => {
    classExists(name); if (classes.has(name)) invalid("A cell style class must have one definition or binding."); classes.add(name);
  };
  for (const definition of styles.definitions) {
    if (definitions.has(definition.styleId)) invalid("Cell style definition identities must be unique."); definitions.add(definition.styleId);
    bindClass(definition.cssClass); formatExists(definition.numberFormatId);
  }
  for (const name of styles.names) {
    styleExists(name.styleId);
    if (names.has(name.name.toUpperCase())) invalid("Named cell styles must have unique names."); names.add(name.name.toUpperCase());
    if (name.builtinId === 0) { if (normal) invalid("A workbook must not declare competing built-in Normal styles."); normal = true; }
  }
  for (const binding of styles.bindings) { styleExists(binding.styleId); bindClass(binding.cssClass); }
  return { classes, items };
}
