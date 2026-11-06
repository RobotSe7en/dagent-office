import { WorkbookDynamicFilterSchema, WorkbookFilterComparisonSchema } from "@doffi/core";
import type { WorkbookAutoFilter, WorkbookFilterCriterion, WorkbookFilterDate, WorkbookSort } from "@doffi/core";
import { check, elements, fail, S, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, integer, number, one, range, spreadsheetText, value } from "./values.js";

function date(node: XmlNode): WorkbookFilterDate {
  attributes(node, ["year", "month", "day", "hour", "minute", "second", "dateTimeGrouping"]); check(node, [], S);
  const grouping = value(node, "dateTimeGrouping"), names = ["year", "month", "day", "hour", "minute", "second"] as const;
  const end = names.indexOf(grouping as typeof names[number]); if (end < 0) fail(node, "A date filter grouping is invalid.");
  const result: WorkbookFilterDate = { grouping: grouping as WorkbookFilterDate["grouping"], year: integer(node, "year", 1, 9999) };
  for (const [index, key] of names.entries()) if (index > 0 && index <= end) {
    result[key] = integer(node, key, index < 3 ? 1 : 0, index === 1 ? 12 : index === 2 ? 31 : index === 3 ? 23 : 59);
  }
  // Lower-order attributes are ignored by native dateTimeGrouping semantics.
  return result;
}
function criterion(node: XmlNode): WorkbookFilterCriterion | undefined {
  if (node.name === "filters") {
    attributes(node, ["blank", "calendarType"]); check(node, ["filter", "dateGroupItem"], S);
    if (value(node, "calendarType") !== undefined && !["none", "gregorian"].includes(value(node, "calendarType")!)) unsupported(node, "This filter calendar is not implemented.");
    const values = elements(node, "filter", S).map(item => {
      attributes(item, ["val"]); check(item, [], S); const raw = value(item, "val"); if (raw === undefined) fail(item, "A filter value is missing."); return spreadsheetText(raw, item);
    });
    return { kind: "values", blank: boolean(node, "blank"), values, dates: elements(node, "dateGroupItem", S).map(date) };
  }
  if (node.name === "customFilters") {
    attributes(node, ["and"]); check(node, ["customFilter"], S);
    const tests = elements(node, "customFilter", S).map(item => {
      attributes(item, ["operator", "val"]); check(item, [], S);
      const operator = WorkbookFilterComparisonSchema.safeParse(value(item, "operator") ?? "equal"), raw = value(item, "val");
      if (!operator.success || raw === undefined) fail(item, "A custom filter comparison is invalid.");
      return { operator: operator.data, value: spreadsheetText(raw, item) };
    });
    if (!tests.length || tests.length > 2) fail(node, "Custom filters require one or two comparisons.");
    return { kind: "custom", and: boolean(node, "and"), tests };
  }
  if (node.name === "top10") {
    attributes(node, ["top", "percent", "val", "filterVal"]); check(node, [], S);
    const count = number(node, "val"), percent = boolean(node, "percent");
    if (count < 1 || count > (percent ? 100 : 500)) fail(node, "A top/bottom filter count is out of bounds.");
    if (value(node, "filterVal") !== undefined) number(node, "filterVal"); // Derived threshold, not authoritative source.
    return { kind: "top", top: boolean(node, "top", true), percent, count };
  }
  if (node.name === "dynamicFilter") {
    attributes(node, ["type", "val", "maxVal"]); check(node, [], S);
    for (const key of ["val", "maxVal"]) if (value(node, key) !== undefined) number(node, key); // Derived date/average caches.
    if (value(node, "type") === "null") return undefined;
    const type = WorkbookDynamicFilterSchema.safeParse(value(node, "type")); if (!type.success) unsupported(node, "This dynamic filter type is not implemented.");
    return { kind: "dynamic", type: type.data };
  }
  return unsupported(node, "This native filter criterion is outside the declared reading scope.");
}
export function readSort(node: XmlNode | undefined): WorkbookSort | undefined {
  if (!node) return undefined;
  attributes(node, ["columnSort", "caseSensitive", "sortMethod", "ref"]); check(node, ["sortCondition"], S);
  if (boolean(node, "columnSort")) unsupported(node, "Sorting whole columns is not implemented.");
  const method = value(node, "sortMethod") ?? "none";
  if (method !== "none" && method !== "pinYin" && method !== "stroke") fail(node, "Invalid native sort method.");
  const area = range(value(node, "ref") ?? "", node), keys = elements(node, "sortCondition", S).map(item => {
    attributes(item, ["descending", "sortBy", "ref", "customList", "dxfId", "iconSet", "iconId"]); check(item, [], S);
    if ((value(item, "sortBy") ?? "value") !== "value" || value(item, "customList") !== undefined) unsupported(item, "Native color, icon and custom-list sort keys are not implemented.");
    const region = range(value(item, "ref") ?? "", item);
    if (region.from.column !== region.to.column || region.from.row !== area.from.row || region.to.row !== area.to.row) fail(item, "A sort key must cover one column across the sort range.");
    return { column: region.from.column, descending: boolean(item, "descending") };
  });
  if (!keys.length) return undefined; // Native empty sort state has no operation to preserve.
  if (keys.length > 64) fail(node, "A native sort definition exceeds its key count.");
  return { range: area, caseSensitive: boolean(node, "caseSensitive"), method, keys };
}
export function readAutoFilter(node: XmlNode | undefined): WorkbookAutoFilter | undefined {
  if (!node) return undefined;
  attributes(node, ["ref"]); check(node, ["filterColumn", "sortState"], S);
  const area = range(value(node, "ref") ?? "", node), columns = elements(node, "filterColumn", S).map(item => {
    attributes(item, ["colId", "hiddenButton", "showButton"]); check(item, ["filters", "customFilters", "dynamicFilter", "top10"], S);
    const children = elements(item, undefined, S); if (children.length > 1) fail(item, "A filter column has multiple competing criteria.");
    const condition = children[0] ? criterion(children[0]) : undefined;
    return { column: area.from.column + integer(item, "colId", 0, area.to.column - area.from.column), hiddenButton: boolean(item, "hiddenButton"), showButton: boolean(item, "showButton", true), ...(condition ? { criterion: condition } : {}) };
  });
  const sort = readSort(one(node, "sortState")); return { range: area, columns, ...(sort ? { sort } : {}) };
}
