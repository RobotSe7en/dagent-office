import { ident, parse, walk } from "css-tree";
import type { WorkbookManifest } from "../contracts/document.js";
import type { DocumentReadLimits, WorkbookReadLimits, WorkbookModel, WorkbookCell, WorkbookWindow, WorkbookSnapshot } from "../runtime.js";
import { parseContract } from "../schema-registry.js";
import { normalizeCss } from "../document/css.js";
import { assetDigest, validateAssets } from "../document/assets.js";
import { readJson, readText, encode } from "../document/package.js";
import { checkLimit, freeze, invalid, unsupported } from "../document/validation.js";
import { validateWorkbookMetadata } from "./metadata.js";
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
  const formats = validateWorkbookMetadata(metadata.workbook, classes, documentLimits.maxNodes);
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
