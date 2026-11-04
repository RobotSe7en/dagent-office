import type { WorkbookMetadata, WorkbookRange, WorkbookPrint } from "../contracts/workbook.js";
import { checkLimit, invalid } from "../document/validation.js";
import { MAX_SHEET_COLUMNS } from "./chunk-index.js";
import { parseFormula } from "./formula.js";

function checkRange(range: WorkbookRange): void {
  if (range.from.row > range.to.row || range.from.column > range.to.column) invalid("Workbook ranges must be ordered.");
}
/** Sweep rows with a column segment tree; no quadratic pairwise merge scan. */
function checkMerges(ranges: readonly WorkbookRange[]): void {
  const events = ranges.flatMap(range => { checkRange(range); return [{ row: range.from.row, delta: 1, range }, { row: range.to.row + 1, delta: -1, range }]; }).sort((a, b) => a.row - b.row || a.delta - b.delta);
  const maximum = new Int32Array(MAX_SHEET_COLUMNS * 4), lazy = new Int32Array(MAX_SHEET_COLUMNS * 4);
  function update(node: number, left: number, right: number, from: number, to: number, delta: number): void {
    if (from <= left && right <= to) { maximum[node]! += delta; lazy[node]! += delta; return; }
    const middle = (left + right) >>> 1;
    if (from <= middle) update(node * 2, left, middle, from, to, delta);
    if (to > middle) update(node * 2 + 1, middle + 1, right, from, to, delta);
    maximum[node] = lazy[node]! + Math.max(maximum[node * 2]!, maximum[node * 2 + 1]!);
  }
  for (const event of events) {
    update(1, 0, MAX_SHEET_COLUMNS - 1, event.range.from.column, event.range.to.column, event.delta);
    if (maximum[1]! > 1) invalid("Workbook merges must not overlap.");
  }
}
export function validateWorkbookMetadata(workbook: WorkbookMetadata, classes: ReadonlySet<string>, maxItems: number): Set<string> {
  const ids = new Set<string>(), names = new Set<string>(); let visible = 0, items = 0;
  const classExists = (name: string): void => { if (/[\t\n\f\r ]/.test(name) || !classes.has(name)) invalid("A workbook dimension style is not defined."); };
  if (workbook.themeCssClass !== undefined) classExists(workbook.themeCssClass);
  const formats = new Set<string>();
  for (const format of workbook.numberFormats) { if (formats.has(format.formatId)) invalid("Number format identities must be unique."); formats.add(format.formatId); }
  const formatExists = (id: string | undefined): void => { if (id !== undefined && !formats.has(id)) invalid("A workbook number format is not defined."); };
  for (const sheet of workbook.sheets) {
    if (ids.has(sheet.sheetId) || names.has(sheet.name.toUpperCase()) || sheet.name.startsWith("'") || sheet.name.endsWith("'")) invalid("Worksheet identities and names must be unique and valid.");
    ids.add(sheet.sheetId); names.add(sheet.name.toUpperCase()); if (sheet.visibility === "visible") visible++;
    classExists(sheet.defaultCssClass);
    if (sheet.tabCssClass !== undefined) { classExists(sheet.tabCssClass); if (sheet.tabColor !== undefined) invalid("A worksheet tab color must have one source."); }
    formatExists(sheet.defaultNumberFormatId);
    for (const dimensions of [sheet.rowDimensions, sheet.columnDimensions]) {
      let previous = -1;
      for (const dimension of dimensions) {
        if (dimension.from > dimension.to || dimension.from <= previous) invalid("Workbook dimension spans must be ordered and non-overlapping.");
        previous = dimension.to; if (dimension.cssClass !== undefined) classExists(dimension.cssClass);
        formatExists(dimension.numberFormatId);
      }
    }
    items += sheet.rowDimensions.length + sheet.columnDimensions.length + sheet.merges.length + 1; checkLimit(items, maxItems);
    checkMerges(sheet.merges);
    // Overlapping links would make a cell's navigation target ambiguous.
    checkMerges((sheet.hyperlinks ?? []).map(link => link.range));
    for (const link of sheet.hyperlinks ?? []) {
      if (!link.url && !link.location) invalid("A workbook hyperlink requires a target.");
      if (link.url !== undefined && (!/^(https?:\/\/|mailto:)/i.test(link.url) || /[\u0000-\u0020\u007f]/.test(link.url))) invalid("A workbook hyperlink has an unsupported target.");
      if (link.location !== undefined) {
        const target = parseFormula("=" + link.location, { row: 0, column: 0 });
        if (!["name", "cell", "range"].includes(target.kind)) invalid("A workbook internal hyperlink must reference a name or cell range.");
      }
    }
    items += sheet.hyperlinks?.length ?? 0;
    if (sheet.print) items += validatePrint(sheet.print, classExists);
    checkLimit(items, maxItems);
  }
  if (!visible || !workbook.sheets.some(sheet => sheet.sheetId === workbook.activeSheetId && sheet.visibility === "visible")) invalid("A workbook needs a visible active worksheet.");
  const nameIds = new Set<string>(), scopedNames = new Set<string>();
  for (const name of workbook.names ?? []) {
    if (nameIds.has(name.nameId) || name.sheetId !== undefined && !ids.has(name.sheetId)) invalid("Defined name identities and scopes must be valid.");
    nameIds.add(name.nameId);
    const key = JSON.stringify([name.sheetId ?? null, name.name.toUpperCase()]);
    if (scopedNames.has(key) || /^[RC]$/i.test(name.name) || /^R\d+C\d+$/i.test(name.name) || name.name.toLowerCase().startsWith("_xlnm.")) invalid("Defined names must be unique in their scope and must not use reserved names.");
    scopedNames.add(key);
    const token = parseFormula("=" + name.name, { row: 0, column: 0 });
    if (token.kind !== "name" || token.sheet !== undefined || token.name !== name.name) invalid("A defined name is not a valid formula identifier.");
    parseFormula(name.expression, name.origin);
  }
  items += nameIds.size;
  checkLimit(items + formats.size, maxItems);
  return formats;
}

function validatePrint(print: WorkbookPrint, classExists: (name: string) => void): number {
  let items = print.areas.length + print.rowBreaks.length + print.columnBreaks.length + 1;
  for (const area of print.areas) checkRange(area);
  for (const span of [print.repeatRows, print.repeatColumns]) if (span && span.from > span.to) invalid("Print title ranges must be ordered.");
  if (print.paper && print.marginsPt && print.orientation !== "default") {
    const width = print.orientation === "portrait" ? print.paper.widthPt : print.paper.heightPt;
    const height = print.orientation === "portrait" ? print.paper.heightPt : print.paper.widthPt;
    if (print.marginsPt.left + print.marginsPt.right >= width || print.marginsPt.top + print.marginsPt.bottom >= height
      || print.marginsPt.header >= height || print.marginsPt.footer >= height) invalid("Print margins leave no page content area.");
  }
  for (const breaks of [print.rowBreaks, print.columnBreaks]) {
    let previous = -1;
    for (const pageBreak of breaks) {
      if (pageBreak.before <= previous || pageBreak.from > pageBreak.to) invalid("Print breaks must have ordered positions and valid spans.");
      previous = pageBreak.before;
    }
  }
  for (const key of ["oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"] as const) {
    const band = print.headerFooter[key]; if (!band) continue;
    for (const runs of [band.left, band.center, band.right]) for (const run of runs) {
      items++; if (run.cssClass !== undefined) classExists(run.cssClass);
    }
  }
  return items;
}

