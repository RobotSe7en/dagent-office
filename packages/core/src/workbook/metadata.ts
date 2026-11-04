import type { WorkbookMetadata, WorkbookPrint } from "../contracts/workbook.js";
import { checkLimit, invalid } from "../document/validation.js";
import { checkRange, checkNonOverlappingRanges } from "./ranges.js";
import { validateWorkbookRules } from "./rules.js";
import { validateWorkbookTables } from "./tables.js";
import { parseFormula } from "./formula.js";

export function validateWorkbookMetadata(workbook: WorkbookMetadata, classes: ReadonlySet<string>, maxItems: number): Set<string> {
  const ids = new Set<string>(), names = new Set<string>(); let visible = 0, items = 0;
  const classExists = (name: string): void => { if (/[\t\n\f\r ]/.test(name) || !classes.has(name)) invalid("A workbook dimension style is not defined."); };
  if (workbook.themeCssClass !== undefined) classExists(workbook.themeCssClass);
  const formats = new Set<string>(), ruleIds = new Set<string>();
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
    checkNonOverlappingRanges(sheet.merges);
    // Overlapping links would make a cell's navigation target ambiguous.
    checkNonOverlappingRanges((sheet.hyperlinks ?? []).map(link => link.range));
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
    items += validateWorkbookRules(sheet, ruleIds, classExists, formatExists, maxItems);
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
  items += validateWorkbookTables(workbook, classExists, formatExists, maxItems);
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

