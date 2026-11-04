import { parseWorkbookFormula, WorkbookTableStyleReferenceSchema, WorkbookTableStyleRegionSchema } from "@dagent-office/core";
import type { WorkbookTable, WorkbookTableStyle, WorkbookTableStyleReference } from "@dagent-office/core";
import type { OpcPackage } from "../ooxml/package.js";
import { attr, check, elements, fail, R, S, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { readAutoFilter, readSort } from "./filter.js";
import type { DifferentialStyleReference, XlsxStyles } from "./styles.js";
import { attributes, boolean, integer, one, range, spreadsheetText, value } from "./values.js";

export interface NativeTableStyles {
  readonly definitions: WorkbookTableStyle[];
  readonly defaultStyle?: WorkbookTableStyleReference;
  resolve(name: string, node: XmlNode): WorkbookTableStyleReference;
}
export function readTableStyles(styles: XlsxStyles): NativeTableStyles {
  const root = styles.tableStyles, definitions: WorkbookTableStyle[] = [], names = new Map<string, string>();
  const resolve = (name: string, node: XmlNode): WorkbookTableStyleReference => {
    const id = names.get(name.toUpperCase()); if (id) return { kind: "custom", styleId: id };
    const builtin = WorkbookTableStyleReferenceSchema.safeParse({ kind: "builtin", name });
    if (!builtin.success) fail(node, "A table style definition is missing or unknown."); return builtin.data;
  };
  if (root) {
    attributes(root, ["count", "defaultTableStyle", "defaultPivotStyle"]); check(root, ["tableStyle"], S);
    for (const node of elements(root, "tableStyle", S)) {
      attributes(node, ["name", "pivot", "table", "count"]); check(node, ["tableStyleElement"], S);
      if (!boolean(node, "table", true)) unsupported(node, "Custom styles exclusively for pivot tables are not implemented.");
      if (value(node, "pivot") !== undefined) boolean(node, "pivot");
      const name = spreadsheetText(value(node, "name") ?? "", node, 255), styleId = `table-style-${definitions.length + 1}`;
      if (!name || names.has(name.toUpperCase())) fail(node, "Custom table style names are empty or duplicated."); names.set(name.toUpperCase(), styleId);
      const regions = elements(node, "tableStyleElement", S).map(item => {
        attributes(item, ["type", "size", "dxfId"]); check(item, [], S);
        const type = WorkbookTableStyleRegionSchema.safeParse(value(item, "type")); if (!type.success) unsupported(item, "This table style region is not implemented.");
        const style = value(item, "dxfId") === undefined ? undefined : styles.differential(integer(item, "dxfId"), item);
        const size = integer(item, "size", 1, 9, 1);
        return { region: type.data, size: type.data.endsWith("Stripe") ? size : 1, ...(style ? { style } : {}) };
      });
      if (value(node, "count") !== undefined && integer(node, "count") !== regions.length) fail(node, "A custom table style element count is inconsistent.");
      definitions.push({ styleId, name, pivotCompatible: boolean(node, "pivot", true), elements: regions });
    }
    if (value(root, "count") !== undefined && integer(root, "count") !== definitions.length) fail(root, "The custom table style count is inconsistent.");
  }
  const defaultName = value(root, "defaultTableStyle");
  return { definitions, resolve, ...(defaultName ? { defaultStyle: resolve(spreadsheetText(defaultName, root!, 255), root!) } : {}) };
}
function regionStyles(node: XmlNode, styles: XlsxStyles): Pick<WorkbookTable, "headerStyle" | "dataStyle" | "totalStyle"> {
  const result: { headerStyle?: DifferentialStyleReference; dataStyle?: DifferentialStyleReference; totalStyle?: DifferentialStyleReference } = {};
  for (const [prefix, property] of [["headerRow", "headerStyle"], ["data", "dataStyle"], ["totalsRow", "totalStyle"]] as const) {
    if (value(node, prefix + "CellStyle") !== undefined) unsupported(node, "Named cell styles on native table regions are not mapped yet.");
    if (value(node, prefix + "DxfId") !== undefined) result[property] = styles.differential(integer(node, prefix + "DxfId"), node);
  }
  return result;
}
function tableFormula(node: XmlNode | undefined, row: number, column: number): string | undefined {
  if (!node) return undefined;
  attributes(node, ["array"]); if (boolean(node, "array")) unsupported(node, "Array formulas in native tables are not supported.");
  if (node.content.some(item => typeof item !== "string")) fail(node, "A table formula has non-text content.");
  const expression = "=" + text(node); parseWorkbookFormula(expression, { row, column }); return expression;
}
export function readTables(root: XmlNode, opc: OpcPackage, styles: XlsxStyles, tableStyles: NativeTableStyles): WorkbookTable[] {
  const parent = one(root, "tableParts"); if (!parent) return [];
  attributes(parent, ["count"]); check(parent, ["tablePart"], S);
  const tables: WorkbookTable[] = [], seen = new Set<string>();
  for (const entry of elements(parent, "tablePart", S)) {
    attributes(entry, ["id"], R); check(entry, [], S);
    const relation = opc.relationships(root.part).get(attr(entry, "id", R) ?? "");
    if (!relation || relation.type !== R + "/table" || relation.external || seen.has(relation.target)) fail(entry, "A table relationship is invalid or repeated.");
    seen.add(relation.target); const node = opc.read(relation.target);
    if (node.uri !== S || node.name !== "table" || opc.contentType(node.part) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml") fail(node, "The native table part is invalid.");
    attributes(node, ["id", "name", "displayName", "comment", "ref", "tableType", "headerRowCount", "insertRow", "insertRowShift", "totalsRowCount", "totalsRowShown", "published", "headerRowDxfId", "dataDxfId", "totalsRowDxfId", "headerRowBorderDxfId", "tableBorderDxfId", "totalsRowBorderDxfId", "headerRowCellStyle", "dataCellStyle", "totalsRowCellStyle", "connectionId"]);
    check(node, ["autoFilter", "sortState", "tableColumns", "tableStyleInfo"], S);
    if ((value(node, "tableType") ?? "worksheet") !== "worksheet" || integer(node, "connectionId", 0, 2147483647, 0)) unsupported(node, "Query and XML-mapped tables are not supported.");
    if (boolean(node, "insertRow") || boolean(node, "insertRowShift")) unsupported(node, "Native table insertion placeholder rows are not implemented.");
    if (value(node, "published") !== undefined) boolean(node, "published");
    const nativeId = integer(node, "id", 1, 4294967294), name = spreadsheetText(value(node, "displayName") ?? "", node, 255), objectName = spreadsheetText(value(node, "name") || name, node, 255);
    const area = range(value(node, "ref") ?? "", node), header = integer(node, "headerRowCount", 0, 1, 1) === 1, totals = integer(node, "totalsRowCount", 0, 1, 0) === 1;
    const columnsRoot = one(node, "tableColumns"); if (!columnsRoot) fail(node, "A native table has no columns.");
    attributes(columnsRoot, ["count"]); check(columnsRoot, ["tableColumn"], S);
    const columns = elements(columnsRoot, "tableColumn", S).map((column, index) => {
      attributes(column, ["id", "name", "totalsRowFunction", "totalsRowLabel", "uniqueName", "queryTableFieldId", "headerRowDxfId", "dataDxfId", "totalsRowDxfId", "headerRowCellStyle", "dataCellStyle", "totalsRowCellStyle"]);
      check(column, ["calculatedColumnFormula", "totalsRowFormula"], S);
      if (value(column, "uniqueName") !== undefined || value(column, "queryTableFieldId") !== undefined) unsupported(column, "External table column identities are not implemented.");
      const id = integer(column, "id", 1, 4294967294), columnName = spreadsheetText(value(column, "name") ?? "", column, 255), totalFunction = value(column, "totalsRowFunction") ?? "none";
      if (!["none", "sum", "min", "max", "average", "count", "countNums", "stdDev", "var", "custom"].includes(totalFunction)) fail(column, "An unknown table totals function was declared.");
      const calculatedFormula = tableFormula(one(column, "calculatedColumnFormula"), area.from.row + Number(header), area.from.column + index);
      const totalFormula = tableFormula(one(column, "totalsRowFormula"), area.to.row, area.from.column + index);
      return { columnId: `column-${id}`, name: columnName, totalFunction: totalFunction as WorkbookTable["columns"][number]["totalFunction"], ...regionStyles(column, styles),
        ...(calculatedFormula ? { calculatedFormula } : {}), ...(totalFormula ? { totalFormula } : {}),
        ...(value(column, "totalsRowLabel") === undefined ? {} : { totalLabel: spreadsheetText(value(column, "totalsRowLabel")!, column) }) };
    });
    if (!columns.length || value(columnsRoot, "count") !== undefined && integer(columnsRoot, "count") !== columns.length) fail(columnsRoot, "The native table column count is inconsistent.");
    const info = one(node, "tableStyleInfo"); if (info) { attributes(info, ["name", "showFirstColumn", "showLastColumn", "showRowStripes", "showColumnStripes"]); check(info, [], S); }
    const styleName = value(info, "name"), autoFilter = readAutoFilter(one(node, "autoFilter")), sort = readSort(one(node, "sortState"));
    const table: WorkbookTable = { tableId: `table-${nativeId}`, name, ...(objectName === name ? {} : { objectName }), range: area, header, totals, totalsPreviouslyShown: boolean(node, "totalsRowShown", true), columns,
      ...regionStyles(node, styles), ...(autoFilter ? { autoFilter } : {}), ...(sort ? { sort } : {}),
      ...(styleName ? { style: tableStyles.resolve(spreadsheetText(styleName, info!, 255), info!) } : {}),
      showFirstColumn: boolean(info, "showFirstColumn"), showLastColumn: boolean(info, "showLastColumn"), rowStripes: boolean(info, "showRowStripes"), columnStripes: boolean(info, "showColumnStripes"),
      ...(value(node, "comment") === undefined ? {} : { description: spreadsheetText(value(node, "comment")!, node, 255) }),
    };
    for (const [key, property] of [["tableBorderDxfId", "borderStyle"], ["headerRowBorderDxfId", "headerBorderStyle"], ["totalsRowBorderDxfId", "totalBorderStyle"]] as const)
      if (value(node, key) !== undefined) table[property] = styles.differential(integer(node, key), node);
    tables.push(table);
  }
  if (value(parent, "count") !== undefined && integer(parent, "count") !== tables.length) fail(parent, "The native table part count is inconsistent.");
  return tables;
}
