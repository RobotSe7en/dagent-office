import { ident, parse, walk } from "css-tree";
import type { WorkbookManifest } from "../contracts/document.js";
import type { WorkbookMetadata, WorkbookRange } from "../contracts/workbook.js";
import type { DocumentReadLimits, WorkbookReadLimits, WorkbookModel, WorkbookCell, WorkbookWindow, WorkbookSnapshot } from "../runtime.js";
import { parseContract } from "../schema-registry.js";
import { normalizeCss } from "../document/css.js";
import { assetDigest, validateAssets } from "../document/assets.js";
import { readJson, readText, encode } from "../document/package.js";
import { checkLimit, freeze, invalid, unsupported } from "../document/validation.js";
import { createFormulaTemplatePool } from "./formula.js";
import { CHUNK_ROWS, CHUNK_COLUMNS, MAX_SHEET_COLUMNS, MAX_SHEET_ROWS, indexWorkbookChunk, decodeWorkbookCell } from "./chunk-index.js";
import type { ChunkIndex } from "./chunk-index.js";

export const DEFAULT_WORKBOOK_READ_LIMITS: WorkbookReadLimits = Object.freeze({
  maxCells: 250_000, maxFormulas: 25_000, maxWindowCells: 5000, maxFormulaTemplates: 25_000, maxFormulaKeyCodeUnits: 4 * 1024 * 1024,
});
function readLimits(input: Partial<WorkbookReadLimits> = {}): WorkbookReadLimits {
  for (const [key, value] of Object.entries(input)) if (!Object.hasOwn(DEFAULT_WORKBOOK_READ_LIMITS, key) || !Number.isSafeInteger(value) || value <= 0) invalid("Workbook limits must be positive safe integers with supported names.");
  return Object.freeze({ ...DEFAULT_WORKBOOK_READ_LIMITS, ...input });
}
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
function validateMetadata(workbook: WorkbookMetadata, classes: ReadonlySet<string>, maxItems: number): Set<string> {
  const ids = new Set<string>(), names = new Set<string>(); let visible = 0, items = 0;
  const classExists = (name: string): void => { if (/[\t\n\f\r ]/.test(name) || !classes.has(name)) invalid("A workbook dimension style is not defined."); };
  for (const sheet of workbook.sheets) {
    if (ids.has(sheet.sheetId) || names.has(sheet.name.toUpperCase()) || sheet.name.startsWith("'") || sheet.name.endsWith("'")) invalid("Worksheet identities and names must be unique and valid.");
    ids.add(sheet.sheetId); names.add(sheet.name.toUpperCase()); if (sheet.visibility === "visible") visible++;
    classExists(sheet.defaultCssClass);
    for (const dimensions of [sheet.rowDimensions, sheet.columnDimensions]) {
      let previous = -1;
      for (const dimension of dimensions) {
        if (dimension.from > dimension.to || dimension.from <= previous) invalid("Workbook dimension spans must be ordered and non-overlapping.");
        previous = dimension.to; if (dimension.cssClass !== undefined) classExists(dimension.cssClass);
      }
    }
    items += sheet.rowDimensions.length + sheet.columnDimensions.length + sheet.merges.length + 1; checkLimit(items, maxItems);
    checkMerges(sheet.merges);
  }
  if (!visible || !workbook.sheets.some(sheet => sheet.sheetId === workbook.activeSheetId && sheet.visibility === "visible")) invalid("A workbook needs a visible active worksheet.");
  const formats = new Set<string>();
  for (const format of workbook.numberFormats) { if (formats.has(format.formatId)) invalid("Number format identities must be unique."); formats.add(format.formatId); }
  checkLimit(items + formats.size, maxItems);
  return formats;
}

/** Receives an isolated package copy. Indexed chunks take sole ownership of their
 * stored bytes; the remaining map retains only non-chunk package files. */
export function openWorkbookModel(files: Map<string, Uint8Array>, manifest: WorkbookManifest, documentLimits: DocumentReadLimits, inputLimits?: Partial<WorkbookReadLimits>): WorkbookModel {
  const limits = readLimits(inputLimits), metadata = parseContract("document-metadata", readJson(files, manifest.metadata));
  if (!("workbook" in metadata)) invalid("Manifest kind and workbook metadata must match.");
  const annotations = parseContract("annotations", readJson(files, manifest.annotations));
  if (annotations.length) unsupported("Workbook annotations are not implemented yet.");
  validateAssets(manifest.assets, files);
  const cssBudget = { nodes: 0 }, classes = new Set<string>();
  const styles = manifest.styles.map(path => ({ path, css: normalizeCss(readText(files, path), "stylesheet", documentLimits, cssBudget) }));
  for (const style of styles) walk(parse(style.css), node => { if (node.type === "ClassSelector") classes.add(ident.decode(node.name)); });
  const formats = validateMetadata(metadata.workbook, classes, documentLimits.maxNodes);
  const sheets = new Map(metadata.workbook.sheets.map(sheet => [sheet.sheetId, new Map<number, ChunkIndex>()]));
  const chunks: { path: string; index: ChunkIndex }[] = [], formulas = createFormulaTemplatePool(limits.maxFormulaTemplates, limits.maxFormulaKeyCodeUnits);
  let cellCount = 0, formulaCount = 0;
  for (const chunk of manifest.chunks) {
    const sheet = sheets.get(chunk.sheetId), key = chunk.row / CHUNK_ROWS * (MAX_SHEET_COLUMNS / CHUNK_COLUMNS) + chunk.column / CHUNK_COLUMNS;
    if (!sheet || sheet.has(key)) invalid("Workbook chunks must have unique positions on existing worksheets.");
    const bytes = files.get(chunk.path);
    if (!bytes || assetDigest(bytes) !== chunk.sha256) invalid("Workbook chunk bytes do not match their digest.");
    const index = indexWorkbookChunk(bytes, chunk, documentLimits.maxFileBytes, { formulas, cssClasses: classes, numberFormats: formats, maxNodes: documentLimits.maxNodes, maxDepth: documentLimits.maxDepth });
    cellCount += index.cellCount; formulaCount += index.formulaCount;
    checkLimit(cellCount, limits.maxCells); checkLimit(formulaCount, limits.maxFormulas);
    sheet.set(key, index); chunks.push({ path: chunk.path, index }); files.delete(chunk.path);
  }
  const snapshot: WorkbookSnapshot = freeze({ manifest, metadata, annotations, styles });
  files.set("manifest.json", encode(JSON.stringify(manifest)));
  files.set(manifest.metadata, encode(JSON.stringify(metadata))); files.set(manifest.annotations, encode(JSON.stringify(annotations)));
  for (const style of styles) files.set(style.path, encode(style.css));
  let total = chunks.reduce((sum, chunk) => sum + chunk.index.sourceBytes, 0);
  for (const bytes of files.values()) { checkLimit(bytes.byteLength, documentLimits.maxFileBytes); total += bytes.byteLength; }
  checkLimit(total, documentLimits.maxPackageBytes);
  function readWindow(window: WorkbookWindow): readonly WorkbookCell[] {
    if (Object.keys(window).some(key => !["sheetId", "row", "column", "rows", "columns"].includes(key))
      || ![window.row, window.column, window.rows, window.columns].every(Number.isSafeInteger)
      || window.row < 0 || window.column < 0 || window.rows <= 0 || window.columns <= 0
      || window.row + window.rows > MAX_SHEET_ROWS || window.column + window.columns > MAX_SHEET_COLUMNS) invalid("Invalid workbook window.");
    checkLimit(window.rows * window.columns, limits.maxWindowCells);
    const sheet = sheets.get(window.sheetId); if (!sheet) invalid("A workbook query references an unknown worksheet.");
    const cells: WorkbookCell[] = [];
    for (let row = Math.floor(window.row / CHUNK_ROWS); row <= Math.floor((window.row + window.rows - 1) / CHUNK_ROWS); row++) {
      for (let column = Math.floor(window.column / CHUNK_COLUMNS); column <= Math.floor((window.column + window.columns - 1) / CHUNK_COLUMNS); column++) {
        const chunk = sheet.get(row * (MAX_SHEET_COLUMNS / CHUNK_COLUMNS) + column);
        for (const source of chunk?.readWindow(window) ?? []) cells.push({ sheetId: window.sheetId, row: source.row, column: source.column, html: source.html, ...decodeWorkbookCell(source.html) });
      }
    }
    cells.sort((a, b) => a.row - b.row || a.column - b.column);
    return freeze(cells);
  }
  return Object.freeze({
    kind: "workbook", snapshot: () => snapshot, getNode: () => undefined,
    readWindow,
    getCell: (sheetId: string, row: number, column: number) => readWindow({ sheetId, row, column, rows: 1, columns: 1 })[0],
    serialize: () => new Map([...Array.from(files, ([path, bytes]) => [path, new Uint8Array(bytes)] as const), ...chunks.map(chunk => [chunk.path, chunk.index.serialize()] as const)]),
  });
}
