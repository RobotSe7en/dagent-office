import type { WorkbookSheet } from "../contracts/workbook.js";
import { checkLimit, invalid } from "../document/validation.js";
import { parseFormula } from "./formula.js";
import { checkRange, checkNonOverlappingRanges } from "./ranges.js";

export function validateWorkbookRules(sheet: WorkbookSheet, ids: Set<string>, classExists: (name: string) => void, formatExists: (name: string | undefined) => void, maxItems: number): number {
  let items = 0;
  const priorities = new Set<number>();
  for (const rule of [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []]) {
    items += 1 + rule.ranges.length; checkLimit(items, maxItems);
    if (ids.has(rule.ruleId)) invalid("Workbook rule identities must be unique."); ids.add(rule.ruleId);
    for (const range of rule.ranges) checkRange(range);
    if ("priority" in rule) {
      if (priorities.has(rule.priority)) invalid("Conditional rule priorities must be unique within a worksheet."); priorities.add(rule.priority);
    }
    if ("cssClass" in rule && rule.cssClass !== undefined) classExists(rule.cssClass);
    if ("numberFormatId" in rule) formatExists(rule.numberFormatId);
    if ("expression" in rule) parseFormula(rule.expression, rule.origin);
    if ("formulas" in rule) {
      const count = ["between", "notBetween"].includes(rule.operator) ? 2 : 1;
      if (rule.formulas.length !== count) invalid("A comparison rule has an invalid number of formulas.");
      for (const formula of rule.formulas) parseFormula(formula, rule.origin);
    }
    if (rule.kind === "colorScale") {
      if (rule.stopIfTrue) invalid("A color scale cannot stop subsequent conditional rules.");
      items += rule.stops.length;
      for (const [index, stop] of rule.stops.entries()) {
        classExists(stop.cssClass);
        if (stop.kind === "min" && index !== 0 || stop.kind === "max" && index !== rule.stops.length - 1) invalid("A color scale minimum or maximum must be at its corresponding end.");
        if (stop.kind === "formula") parseFormula(stop.expression, rule.origin);
      }
    }
  }
  // Overlapping validation definitions would make enforcement ambiguous.
  checkNonOverlappingRanges((sheet.dataValidations ?? []).flatMap(rule => rule.ranges));
  checkLimit(items, maxItems); return items;
}
