import { DocumentError } from "@dagent-office/core";
import type { WorkbookNumberFormatter, WorkbookSnapshot } from "@dagent-office/core";
import type { FontResolution } from "../contracts.js";
import type { WorkbookAssetSource } from "./assets.js";
import type { WorkbookChartSource } from "./chart-source.js";
import { measureWorkbookColumnFont } from "./column-font.js";
import type { WorkbookColumnFont } from "./column-font.js";
import { WorkbookContentSources } from "./content-source.js";
import { worksheetDimensions } from "./dimensions.js";
import { layoutWorkbookDrawings } from "./drawing-layout.js";
import type { WorkbookDrawingLayout, WorkbookDrawingLayoutLimits } from "./drawing-layout.js";
import type { WorkbookSurface } from "./surface.js";
import type { WorkbookViewport } from "./viewport.js";
import { layoutWorkbookWindow } from "./window-layout.js";
import type { WorkbookWindowLayout, WorkbookWindowLayoutOptions } from "./window-layout.js";
import { WorkbookWindowLoader } from "./window-loader.js";
import type { WorkbookViewportReadLimits, WorkbookWindowFrame } from "./window-loader.js";
import { checkWindowIdentity } from "./window-source.js";
import type { WorkbookWindowSource } from "./window-source.js";

export interface WorkbookWindowProviders {
  readonly cells: WorkbookWindowSource;
  readonly assets: WorkbookAssetSource;
  readonly charts: WorkbookChartSource;
}
export interface WorkbookWindowControllerOptions {
  /** Combined final cell/object nodes, independently of each phase's budget. */
  readonly maxNodes?: number;
  readonly read?: Partial<WorkbookViewportReadLimits>;
  readonly cells?: Omit<WorkbookWindowLayoutOptions, "normalSpaceWidthPx">;
  readonly drawings?: WorkbookDrawingLayoutLimits;
}
export interface WorkbookRenderedWindow {
  readonly frame: WorkbookWindowFrame;
  readonly element: HTMLElement;
  readonly nodes: number;
  readonly dimensions: ReturnType<typeof worksheetDimensions>;
  readonly cells: WorkbookWindowLayout["cells"];
  readonly objects: WorkbookDrawingLayout["objects"];
  readonly fonts: readonly FontResolution[];
  readonly diagnostics: WorkbookDrawingLayout["diagnostics"];
  readonly usedFallbackLineMetrics: boolean;
}

/** Internal lifecycle for one immutable workbook. The owner supplies a ready,
 * dedicated surface and authorized providers, and owns their disposal. This
 * controller owns only its window candidates and auxiliary source index.
 * A slow, failed or cancelled request never replaces the published window. */
export class WorkbookWindowController {
  private readonly snapshot: WorkbookSnapshot;
  private readonly content: WorkbookContentSources;
  private readonly options: WorkbookWindowControllerOptions;
  private readonly maxNodes: number;
  private readonly sheets = new Map<string, { readonly normal: WorkbookColumnFont; readonly dimensions: ReturnType<typeof worksheetDimensions> }>();
  private pending: { cancel(): void } | undefined;
  private active: { readonly state: WorkbookRenderedWindow; readonly drawings: WorkbookDrawingLayout } | undefined;
  private disposed = false;

  constructor(private readonly surface: WorkbookSurface, private readonly providers: WorkbookWindowProviders,
    private readonly formatter: WorkbookNumberFormatter, input: WorkbookWindowControllerOptions = {}) {
    this.maxNodes = input.maxNodes ?? 50_000;
    if (!Number.isSafeInteger(this.maxNodes) || this.maxNodes < 1 || Object.keys(input).some(key => !["maxNodes", "read", "cells", "drawings"].includes(key))
      || !surface.source.isConnected || !surface.view.isConnected || surface.source.ownerDocument !== surface.document || surface.view.ownerDocument !== surface.document
      || surface.view.childNodes.length) throw new DocumentError("InvalidOperation", "A workbook controller requires its own empty connected surface and supported node budget.");
    this.snapshot = providers.cells.snapshot;
    checkWindowIdentity(providers.assets.snapshot.manifest, this.snapshot); checkWindowIdentity(providers.charts.snapshot.manifest, this.snapshot);
    this.options = Object.freeze({ ...input, read: Object.freeze({ ...input.read }), cells: Object.freeze({ ...input.cells }),
      drawings: Object.freeze({ ...input.drawings, ...(input.drawings?.assets ? { assets: Object.freeze({ ...input.drawings.assets }) } : {}) }) });
    // Validate read limits before retaining auxiliary content or starting work.
    const probe = new WorkbookWindowLoader(providers.cells, this.options.read); probe.dispose();
    this.content = new WorkbookContentSources(surface.document, this.snapshot);
  }

  get current(): WorkbookRenderedWindow | undefined { return this.active?.state; }

  /** False means superseded/disposed; genuine failures reject. A newer request
   * also settles an older request whose provider ignores its AbortSignal. */
  async render(sheetId: string, input: WorkbookViewport): Promise<boolean> {
    if (this.disposed) throw new DocumentError("InvalidOperation", "This workbook window controller has been disposed.");
    const viewport = Object.freeze({ widthPt: input.widthPt, heightPt: input.heightPt, scrollLeftPt: input.scrollLeftPt, scrollTopPt: input.scrollTopPt });
    if (![viewport.widthPt, viewport.heightPt, viewport.scrollLeftPt, viewport.scrollTopPt].every(Number.isFinite)
      || Math.min(viewport.widthPt, viewport.heightPt) <= 0 || Math.max(viewport.widthPt, viewport.heightPt) > 1_000_000)
      throw new DocumentError("InvalidOperation", "A workbook window requires finite positive viewport extents.");
    const sheet = this.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId);
    if (!sheet) throw new DocumentError("InvalidOperation", "The requested worksheet is not part of this workbook.");
    this.pending?.cancel();
    const abort = new AbortController(), loader = new WorkbookWindowLoader(this.providers.cells, this.options.read);
    let cancel!: () => void;
    const cancelled = new Promise<false>(resolve => { cancel = () => { abort.abort(); loader.dispose(); resolve(false); }; });
    const request = { cancel }; this.pending = request;
    const current = (): boolean => !this.disposed && this.pending === request && !abort.signal.aborted;
    let drawings: WorkbookDrawingLayout | undefined;
    try {
      let metrics = this.sheets.get(sheetId);
      if (!metrics) {
        const normal = measureWorkbookColumnFont(this.surface.source, this.snapshot.metadata.workbook, sheetId);
        metrics = { normal, dimensions: worksheetDimensions(sheet, normal.maximumDigitWidthPx) }; this.sheets.set(sheetId, metrics);
      }
      const { normal, dimensions } = metrics;
      if (!await Promise.race([loader.load(sheetId, dimensions, viewport), cancelled]) || !current()) return false;
      const frame = loader.frame!;
      const cells = layoutWorkbookWindow(this.surface, this.snapshot, frame, dimensions, this.formatter,
        { ...this.options.cells, normalSpaceWidthPx: normal.spaceWidthPx, maxNodes: Math.min(this.maxNodes, this.options.cells?.maxNodes ?? this.maxNodes) });
      const walker = this.surface.document.createTreeWalker(cells.element, 0xffff_ffff); let nodes = 1;
      while (walker.nextNode()) if (++nodes >= this.maxNodes) throw new DocumentError("ResourceLimitExceeded", "The combined workbook window exceeds its node budget.");
      if (nodes >= this.maxNodes) throw new DocumentError("ResourceLimitExceeded", "The combined workbook window has no remaining object-layer budget.");
      drawings = await layoutWorkbookDrawings(this.surface, this.snapshot, frame, dimensions, this.content, this.providers.assets, this.providers.charts,
        this.formatter, abort.signal, { ...this.options.drawings, maxNodes: Math.min(this.maxNodes - nodes, this.options.drawings?.maxNodes ?? this.maxNodes) });
      if (!current()) return false;
      nodes += drawings.nodes; cells.element.append(drawings.element);
      const state: WorkbookRenderedWindow = Object.freeze({ frame, element: cells.element, nodes, dimensions, cells: cells.cells, objects: drawings.objects,
        fonts: Object.freeze([...normal.fonts, ...cells.fonts, ...drawings.fonts]), diagnostics: drawings.diagnostics,
        usedFallbackLineMetrics: cells.usedFallbackLineMetrics || drawings.usedFallbackLineMetrics });
      const previous = this.active;
      this.surface.view.replaceChildren(cells.element);
      this.surface.frame.style.setProperty("width", viewport.widthPt + "pt", "important");
      this.surface.frame.style.setProperty("height", viewport.heightPt + "pt", "important");
      this.surface.frame.style.setProperty("visibility", "visible", "important");
      this.active = { state, drawings }; drawings = undefined;
      previous?.drawings.dispose(); previous?.state.element.replaceChildren();
      return true;
    } catch (error) {
      if (!current()) return false;
      throw error;
    } finally {
      drawings?.dispose(); loader.dispose();
      if (this.pending === request) this.pending = undefined;
    }
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true; this.pending?.cancel(); this.pending = undefined;
    this.active?.drawings.dispose(); this.active?.state.element.remove(); this.active?.state.element.replaceChildren(); this.active = undefined;
    this.content.dispose(); this.sheets.clear();
  }
}
