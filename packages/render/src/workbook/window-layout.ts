import { DocumentError } from "@dagent-office/core";
import type { WorkbookDisplayFormat, WorkbookNumberFormatter, WorkbookSnapshot } from "@dagent-office/core";
import type { WorkbookFormulaSource, WorkbookNumberDisplay } from "@dagent-office/core";
import type { FontResolution } from "../contracts.js";
import { WorkbookCellSources } from "./cell-source.js";
import { workbookCellBorders, workbookCellDiagonals } from "./cell-borders.js";
import { workbookCellFill } from "./cell-fill.js";
import { workbookCellShape } from "./cell-shape.js";
import { createWorkbookTextView } from "./text-view.js";
import { createWorkbookNumberText } from "./number-text.js";
import { layoutWorkbookTexts } from "./text-layout.js";
import type { WorkbookTextLayout, WorkbookTextLayoutInput } from "./text-layout.js";
import { layoutWorkbookNumbers } from "./number-text-layout.js";
import type { WorkbookNumberLayoutInput } from "./number-text-layout.js";
import { workbookTextRotation } from "./text-alignment.js";
import { projectWorkbookBorders } from "./border-projection.js";
import { projectWorkbookDiagonals } from "./diagonal-projection.js";
import { createWorkbookBorderView } from "./border-view.js";
import type { AxisDimensions } from "./dimensions.js";
import { WorkbookMergeIndex } from "./projection.js";
import type { WorkbookGridCell } from "./projection.js";
import type { WorkbookSurface } from "./surface.js";
import type { WorkbookWindowFrame } from "./window-loader.js";
import { paintWorkbookWindow } from "./window-paint.js";
import { layoutWorkbookFormattedTexts } from "./formatted-text.js";
import type { WorkbookFormattedTextInput } from "./formatted-text.js";

export interface WorkbookWindowLayoutOptions {
  readonly normalSpaceWidthPx: number;
  readonly maxNodes?: number;
  readonly maxSourceNodes?: number;
  readonly maxTextCodeUnits?: number;
}
export interface WorkbookWindowLayout {
  readonly element: HTMLElement;
  readonly fonts: readonly FontResolution[];
  readonly usedFallbackLineMetrics: boolean;
  readonly cells: readonly {
    readonly projection: WorkbookGridCell;
    readonly text?: WorkbookTextLayout;
    readonly cacheState?: WorkbookFormulaSource["cacheState"];
    readonly numberOverflow?: Extract<WorkbookNumberDisplay, { kind: "overflow" }>["reason"];
  }[];
}

/** Internal cell compositor. Source CSS is measured in the source surface;
 * derived text is measured in the isolated display surface. This prepares a
 * whole detached grid and never replaces the caller's published window.
 * Objects/rules/automatic row heights belong to the enclosing preview stage. */
export function layoutWorkbookWindow(surface: Pick<WorkbookSurface, "source" | "view">, snapshot: WorkbookSnapshot, frame: WorkbookWindowFrame,
  dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, formatter: WorkbookNumberFormatter,
  options: WorkbookWindowLayoutOptions): WorkbookWindowLayout {
  const maxNodes = options.maxNodes ?? 50_000, maxSourceNodes = options.maxSourceNodes ?? 50_000, maxTextUnits = options.maxTextCodeUnits ?? 8 * 1024 * 1024;
  if (![maxNodes, maxSourceNodes, maxTextUnits].every(n => Number.isSafeInteger(n) && n > 0) || !Number.isFinite(options.normalSpaceWidthPx) || options.normalSpaceWidthPx < 0
    || surface.source.ownerDocument !== surface.view.ownerDocument || !surface.view.isConnected) throw new DocumentError("InvalidOperation", "Workbook window layout requires one connected surface and positive budgets.");
  const document = surface.source.ownerDocument, sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === frame.sheetId);
  if (!sheet) throw new DocumentError("InvalidOperation", "The workbook window has no source worksheet.");
  const sources = new WorkbookCellSources(surface.source, snapshot, frame, maxSourceNodes);
  try {
    const styles = new Map<number, CSSStyleDeclaration>();
    const style = (row: number, column: number): CSSStyleDeclaration => {
      const key = row * 16_384 + column, existing = styles.get(key); if (existing) return existing;
      const source = sources.get(row, column);
      if (!source) throw new DocumentError("InvalidOperation", "A workbook appearance source is missing from its window.");
      const value = document.defaultView!.getComputedStyle(source.element); styles.set(key, value); return value;
    };
    const bordersAt = (row: number, column: number) => workbookCellBorders(document, style(row, column), sheet.rightToLeft);
    const bordered = (cell: WorkbookGridCell): boolean => {
      const { from, to } = cell.range, anchor = bordersAt(from.row, from.column);
      return Boolean(anchor.top || anchor.before || bordersAt(from.row, to.column).after || bordersAt(to.row, from.column).bottom);
    };
    const formats = new Map(snapshot.metadata.workbook.numberFormats.map(format => [format.formatId, format]));
    const texts: WorkbookTextLayoutInput[] = [], numbers: WorkbookNumberLayoutInput[] = [], formatted: WorkbookFormattedTextInput[] = [];
    const locations = new Map<WorkbookGridCell, { kind: "text" | "number" | "formatted"; index: number }>();
    let textNodes = 0, units = 0;
    const remaining = (): number => {
      if (textNodes >= maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook window text exceeds its cumulative node budget.");
      return maxNodes - textNodes;
    };
    for (const projection of frame.projections) {
      const cell = frame.getCell(projection.row, projection.column), source = sources.get(projection.row, projection.column)!;
      if (!cell || cell.value.type === "blank") continue;
      const format: WorkbookDisplayFormat | undefined = source.format.numberFormatId ? formats.get(source.format.numberFormatId) : undefined;
      const size = { widthPt: projection.rect.width, heightPt: projection.rect.height, bordered: bordered(projection) };
      if (cell.value.type === "number") {
        locations.set(projection, { kind: "number", index: numbers.length }); numbers.push({ source, value: cell.value.value, ...(format ? { format } : {}), ...size });
        continue;
      }
      if (cell.value.type === "text" && format) {
        const display = formatter.format(cell.value.value, format);
        if (display.kind !== "parts" || display.color || display.parts.some(part => part.kind !== "text")
          || display.parts.length !== 1 || display.parts[0]?.kind !== "text" || !display.parts[0].source) {
          // Rich-text/native-format precedence is a pending product choice.
          // Plain text already uses precise source fields and derived affixes.
          if (source.element.querySelector("span")) throw new DocumentError("UnsupportedFeature", "Rich text combined with a text number format awaits its display policy.");
          locations.set(projection, { kind: "formatted", index: formatted.length }); formatted.push({ source, value: cell.value.value, format, ...size }); continue;
        }
      }
      const text = cell.value.type === "text" ? createWorkbookTextView(source, remaining())
        : createWorkbookNumberText(source, [{ kind: "text", text: cell.value.type === "boolean" ? cell.value.value ? "TRUE" : "FALSE" : cell.value.value }], undefined, remaining());
      textNodes += text.nodes + 1; units += text.element.textContent?.length ?? 0;
      if (units > maxTextUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook window text exceeds its cumulative source-text budget.");
      locations.set(projection, { kind: "text", index: texts.length }); texts.push({ source, text, valueType: cell.value.type, ...size });
    }
    const rich = layoutWorkbookTexts(surface.view, texts, options.normalSpaceWidthPx, maxNodes);
    textNodes = rich.cells.reduce((sum, cell) => sum + cell.text.nodes + 1, 0);
    const formattedText = layoutWorkbookFormattedTexts(surface.view, formatted, formatter, options.normalSpaceWidthPx, remaining(), Math.max(1, maxTextUnits - units));
    textNodes += formattedText.cells.reduce((sum, cell) => sum + cell.text.nodes + 1, 0);
    units += formattedText.cells.reduce((sum, cell) => sum + (cell.text.element.textContent?.length ?? 0), 0);
    const numeric = layoutWorkbookNumbers(surface.view, numbers, formatter, options.normalSpaceWidthPx, remaining(), Math.max(1, maxTextUnits - units));
    textNodes += numeric.cells.reduce((sum, cell) => sum + cell.text.nodes + 1, 0);
    remaining();
    const prepared = frame.projections.map(projection => {
      const computed = style(projection.row, projection.column), location = locations.get(projection);
      const text = location ? (location.kind === "number" ? numeric : location.kind === "formatted" ? formattedText : rich).cells[location.index]! : undefined;
      const shape = text?.shape ?? workbookCellShape(projection.rect.width, projection.rect.height,
        workbookTextRotation(computed).rotationDegrees, bordered(projection));
      return { projection, shape, fill: workbookCellFill(document, computed), ...(text ? { text } : {}) };
    });
    const merges = new WorkbookMergeIndex(sheet.merges);
    const borders = projectWorkbookBorders(dimensions, frame.layout, merges, bordersAt, sheet.rightToLeft,
      maxNodes, (row, column) => workbookTextRotation(style(row, column)).rotationDegrees);
    const diagonals = projectWorkbookDiagonals(frame.projections, frame.layout, (row, column) => workbookCellDiagonals(document, style(row, column)), sheet.rightToLeft, maxNodes);
    const materialize = (panes: ReturnType<typeof projectWorkbookBorders>): SVGSVGElement[] => panes.filter(pane => pane.edges.length).map(pane => {
      const element = createWorkbookBorderView(document, pane.bounds, pane.edges, remaining()); textNodes += 1 + element.querySelectorAll("*").length; return element;
    });
    const views = materialize([...borders, ...diagonals]);
    const gridPaint = { kind: "css" as const, style: "solid" as const, widthPt: .75, color: "rgb(220, 224, 230)" };
    const gridLines = sheet.showGridLines ? materialize(projectWorkbookBorders(dimensions, frame.layout, merges,
      () => ({ before: gridPaint, after: gridPaint, top: gridPaint, bottom: gridPaint }), sheet.rightToLeft, maxNodes)) : [];
    for (const grid of gridLines) grid.dataset.officeView = "workbook-gridlines";
    const element = paintWorkbookWindow(document, frame, prepared, sheet.rightToLeft, views, maxNodes, gridLines);
    return Object.freeze({ element, fonts: Object.freeze([...rich.fonts, ...formattedText.fonts, ...numeric.fonts]), usedFallbackLineMetrics: rich.usedFallbackLineMetrics || formattedText.usedFallbackLineMetrics || numeric.usedFallbackLineMetrics,
      cells: Object.freeze(prepared.map(cell => {
        const stored = frame.getCell(cell.projection.row, cell.projection.column), location = locations.get(cell.projection);
        const overflow = location?.kind === "number" ? numeric.overflow[location.index] : undefined;
        return Object.freeze({ projection: cell.projection, ...(cell.text ? { text: cell.text } : {}), ...(stored?.formula ? { cacheState: stored.formula.cacheState } : {}),
          ...(overflow ? { numberOverflow: overflow } : {}) });
      })) });
  } finally { sources.dispose(); }
}
