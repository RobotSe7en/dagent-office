import { DocumentError } from "@doffi/core";
import type { WorkbookCell, WorkbookRuleReadResponse, WorkbookSnapshot } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";
import { workbookCellFormat } from "./cell-format.js";
import type { WorkbookCellFormat } from "./cell-format.js";
import type { WorkbookWindowFrame } from "./window-loader.js";
import { WorkbookTableAppearance } from "./table-appearance.js";
import { WorkbookConditionalAppearance } from "./conditional-appearance.js";
import { checkedPlainWorkbookAttributes, checkWorkbookCellSource, workbookCellSourceAttributes } from "./plain-cell-source.js";

export interface WorkbookSourceCell { readonly element: HTMLTableCellElement; readonly format: WorkbookCellFormat }
const runAttributes = new Set(["class"]), noAttributes = new Set<string>();

/** A window-only source projection for browser CSS/font measurement. Keep the
 * original chunk/table/row ancestry and source td/span/br attributes so CSS
 * selectors do not accidentally match the virtual grid's clip divs instead.
 * Only validated model cells belong here; the inert parse and shape check also
 * prevent a broken provider from activating executable or resource content. */
export class WorkbookCellSources {
  private readonly root: HTMLElement;
  private readonly cells = new Map<number, WorkbookSourceCell>();
  private disposed = false;
  constructor(scope: HTMLElement, snapshot: WorkbookSnapshot, frame: Pick<WorkbookWindowFrame, "documentId" | "revision" | "sheetId" | "sourceCoordinates" | "getCell">, maxNodes = 50_000, rules?: WorkbookRuleReadResponse, scalarPrototypes = false) {
    const document = scope.ownerDocument, sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === frame.sheetId);
    if (!sheet || snapshot.manifest.documentId !== frame.documentId || snapshot.manifest.revision !== frame.revision) throw new DocumentError("RevisionConflict", "Cell source styles require the same workbook revision as their window.");
    if (!scope.isConnected || !document.defaultView || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "Cell source measurement requires a connected surface and a positive node budget.");
    this.root = markViewNode(document.createElement("office-workbook-source"));
    this.root.dataset.officeView = "workbook-source-measurement"; this.root.setAttribute("aria-hidden", "true");
    // This tree supplies computed source styles and text identities only. Its
    // table geometry is never measured; display:none avoids solving hundreds
    // of irrelevant source-column widths when the real text probes lay out.
    this.root.style.cssText = "display:none!important;position:absolute;visibility:hidden;pointer-events:none;contain:layout;overflow:hidden;left:0;top:0;width:0;height:0";
    if (snapshot.metadata.workbook.themeCssClass) this.root.className = snapshot.metadata.workbook.themeCssClass;
    let count = 1;
    const account = (amount: number): void => { count += amount; if (count > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook source projection exceeds its node budget."); };
    const tables = new Map<number, { table: HTMLTableElement; body: HTMLTableSectionElement; rows: Map<number, HTMLTableRowElement> }>();
    const reusable = scalarPrototypes && !rules && !sheet.tables?.length && !sheet.conditionalRules?.length
      && snapshot.styles.every(style => /^(?:\s*\.[a-zA-Z_][-\w]*\s*\{[^{}]*\}\s*)*$/.test(style.css));
    const prototypes = new Map<string, HTMLTableCellElement>();
    const coordinates = frame.sourceCoordinates;
    for (const coordinate of coordinates) {
      const { row, column } = coordinate, blockRow = Math.floor(row / 256) * 256, blockColumn = Math.floor(column / 32) * 32, blockKey = blockRow / 256 * 512 + blockColumn / 32;
      let chunk = tables.get(blockKey);
      if (!chunk) {
        account(2); const table = document.createElement("table"), body = document.createElement("tbody");
        table.dataset.sheetId = sheet.sheetId; table.dataset.row = String(blockRow); table.dataset.column = String(blockColumn); table.append(body);
        chunk = { table, body, rows: new Map() }; tables.set(blockKey, chunk);
      }
      let rowElement = chunk.rows.get(row);
      if (!rowElement) { account(1); rowElement = document.createElement("tr"); rowElement.dataset.row = String(row); chunk.rows.set(row, rowElement); chunk.body.append(rowElement); }
      const source = frame.getCell(row, column), format = workbookCellFormat(sheet, coordinate, source);
      // Non-text displays are derived from typed values and have no source
      // character offsets. Class-only scalar typography can share a private
      // connected TD, after every scalar's complete source check. Text/rich
      // cells retain their actual source tree and UTF-16 mapping.
      const attributes = reusable && source && source.value.type !== "text" ? checkedPlainWorkbookAttributes(source) : undefined;
      const prototype = attributes ? prototypes.get(format.cssClass) : undefined;
      if (prototype) { this.cells.set(row * 16_384 + column, Object.freeze({ element: prototype, format })); continue; }
      const element = source ? sourceCell(document, source, account) : document.createElement("td");
      if (!source) { account(1); element.dataset.column = String(column); element.dataset.valueType = "blank"; }
      element.className = format.cssClass; rowElement.append(element);
      if (attributes) prototypes.set(format.cssClass, element);
      this.cells.set(row * 16_384 + column, Object.freeze({ element, format }));
    }
    for (const [, chunk] of [...tables].sort((a, b) => a[0] - b[0])) this.root.append(chunk.table);
    // Nothing is connected until all cells have passed the inert shape/budget
    // checks. The caller disposes this temporary projection after measurement.
    scope.append(this.root);
    try {
      if (sheet.tables?.length) {
        const appearance = new WorkbookTableAppearance(snapshot, sheet, document, account);
        for (const [key, source] of this.cells) this.cells.set(key, appearance.apply({ row: Math.floor(key / 16_384), column: key % 16_384 }, source));
      }
      if (rules) {
        const appearance = new WorkbookConditionalAppearance(document, sheet, rules, account);
        for (const [key, source] of this.cells) this.cells.set(key, appearance.apply(key, source));
      }
    } catch (error) { this.root.remove(); this.root.replaceChildren(); this.cells.clear(); throw error; }
  }
  get(row: number, column: number): WorkbookSourceCell | undefined {
    if (this.disposed) throw new DocumentError("InvalidOperation", "This workbook source projection has been disposed.");
    if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || row >= 1_048_576 || column < 0 || column >= 16_384) return undefined;
    return this.cells.get(row * 16_384 + column);
  }
  dispose(): void { this.root.remove(); this.root.replaceChildren(); this.cells.clear(); this.disposed = true; }
}

function sourceCell(document: Document, source: WorkbookCell, account: (amount: number) => void): HTMLTableCellElement {
  const attributes = checkedPlainWorkbookAttributes(source);
  if (attributes) {
    account(1 + Number(source.text.length > 0)); const cell = document.createElement("td");
    for (const [name, value] of attributes) cell.setAttribute(name, value);
    cell.textContent = source.text; return cell;
  }
  if (typeof source.html !== "string" || source.html.length > 8 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "Workbook cell HTML exceeds its source budget.");
  const template = document.createElement("template"); template.innerHTML = "<table><tbody><tr>" + source.html + "</tr></tbody></table>";
  const single = (parent: ParentNode, name: string): HTMLElement => {
    const nodes = [...parent.childNodes].filter(node => node.nodeType !== 3 || node.textContent?.trim());
    const element = nodes[0]; if (nodes.length !== 1 || element?.nodeType !== 1 || (element as Element).localName !== name) throw new DocumentError("InvalidDocument", "A workbook cell projection has invalid HTML structure.");
    return element as HTMLElement;
  };
  const cell = single(single(single(single(template.content, "table"), "tbody"), "tr"), "td") as HTMLTableCellElement;
  const stack: { node: Node; depth: number }[] = [{ node: cell, depth: 0 }], text: string[] = [];
  while (stack.length) {
    const { node, depth } = stack.pop()!; account(1);
    if (depth > 16) throw new DocumentError("ResourceLimitExceeded", "Workbook cell HTML exceeds its nesting budget.");
    if (node.nodeType === 3) { text.push(node.textContent ?? ""); continue; }
    if (node.nodeType !== 1) throw new DocumentError("InvalidDocument", "A workbook cell projection contains unsupported content.");
    const element = node as Element, allowed = element === cell ? workbookCellSourceAttributes : element.localName === "span" ? runAttributes : element.localName === "br" ? noAttributes : undefined;
    if (!allowed || element.namespaceURI !== "http://www.w3.org/1999/xhtml" || [...element.attributes].some(attribute => attribute.namespaceURI || !allowed.has(attribute.name))) throw new DocumentError("InvalidDocument", "A workbook cell projection contains unsupported elements or attributes.");
    if (element.localName === "br") text.push("\n");
    for (const child of [...element.childNodes].reverse()) stack.push({ node: child, depth: depth + 1 });
  }
  checkWorkbookCellSource(source, name => cell.getAttribute(name), text.join(""));
  return cell;
}
