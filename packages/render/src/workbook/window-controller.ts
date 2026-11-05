import { DEFAULT_WORKBOOK_RULE_LIMITS, DocumentError } from "@dagent-office/core";
import type { WorkbookNumberFormatter, WorkbookRuleReadLimits, WorkbookRuleReadResponse, WorkbookSnapshot } from "@dagent-office/core";
import type { FontResolution } from "../contracts.js";
import type { WorkbookAssetSource } from "./assets.js";
import { projectWorkbookAnnotations } from "./annotation-projection.js";
import type { WorkbookAnnotationProjectionLimits, WorkbookAnnotationWindow } from "./annotation-projection.js";
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
import { checkedWorkbookRules } from "./rule-response.js";
import { workbookRuleDiagnostics } from "./rule-source.js";
import type { WorkbookRuleSource, WorkbookRuleDiagnostic } from "./rule-source.js";
import { maximumWorkbookSkew } from "./skew-sources.js";
import { measureWorkbookAutoRows } from "./auto-rows.js";
import type { WorkbookAutoRowLimits, WorkbookAutoRows } from "./auto-rows.js";

export interface WorkbookWindowProviders {
  readonly cells: WorkbookWindowSource;
  readonly assets: WorkbookAssetSource;
  readonly charts: WorkbookChartSource;
  readonly rules?: WorkbookRuleSource;
}
export interface WorkbookWindowControllerOptions {
  /** Combined final cell/object nodes, independently of each phase's budget. */
  readonly maxNodes?: number;
  readonly read?: Partial<WorkbookViewportReadLimits>;
  readonly cells?: Omit<WorkbookWindowLayoutOptions, "normalSpaceWidthPx">;
  readonly drawings?: WorkbookDrawingLayoutLimits;
  readonly annotations?: WorkbookAnnotationProjectionLimits;
  readonly rules?: Partial<WorkbookRuleReadLimits>;
  readonly autoRows?: Partial<WorkbookAutoRowLimits>;
}
export interface WorkbookRenderedWindow {
  readonly frame: WorkbookWindowFrame;
  readonly element: HTMLElement;
  readonly nodes: number;
  readonly dimensions: ReturnType<typeof worksheetDimensions>;
  readonly cells: WorkbookWindowLayout["cells"];
  readonly objects: WorkbookDrawingLayout["objects"];
  readonly annotations: WorkbookAnnotationWindow;
  readonly fonts: readonly FontResolution[];
  readonly diagnostics: readonly (WorkbookDrawingLayout["diagnostics"][number] | WorkbookRuleDiagnostic)[];
  readonly rules?: WorkbookRuleReadResponse;
  readonly usedFallbackLineMetrics: boolean;
}
export interface WorkbookWindowDecoration {
  /** May change as an owned panel/selection changes, within its assigned cap. */
  readonly nodes: number;
  readonly fonts?: readonly FontResolution[];
  dispose(): void;
}
export type WorkbookWindowDecorator = (candidate: WorkbookRenderedWindow, content: WorkbookContentSources, maxNodes: number) => WorkbookWindowDecoration;

/** Internal lifecycle for one immutable workbook. The owner supplies a ready,
 * dedicated surface and authorized providers, and owns their disposal. This
 * controller owns only its window candidates and auxiliary source index.
 * A slow, failed or cancelled request never replaces the published window. */
export class WorkbookWindowController {
  private readonly snapshot: WorkbookSnapshot;
  private readonly content: WorkbookContentSources;
  private readonly options: WorkbookWindowControllerOptions;
  private readonly maxNodes: number;
  private readonly ruleLimits: WorkbookRuleReadLimits;
  private readonly skewCoefficient: number;
  private readonly sheets = new Map<string, WorkbookAutoRows & { readonly normal: WorkbookColumnFont }>();
  private pending: { cancel(): void } | undefined;
  private active: { readonly state: WorkbookRenderedWindow; readonly drawings: WorkbookDrawingLayout; readonly decoration?: WorkbookWindowDecoration } | undefined;
  private disposed = false;

  constructor(private readonly surface: WorkbookSurface, private readonly providers: WorkbookWindowProviders,
    private readonly formatter: WorkbookNumberFormatter, input: WorkbookWindowControllerOptions = {}, private readonly decorate: WorkbookWindowDecorator | undefined = undefined) {
    this.maxNodes = input.maxNodes ?? 50_000;
    this.ruleLimits = Object.freeze({ ...DEFAULT_WORKBOOK_RULE_LIMITS, ...input.rules });
    if (Object.keys(this.ruleLimits).some(key => !Object.hasOwn(DEFAULT_WORKBOOK_RULE_LIMITS, key)) || !Object.values(this.ruleLimits).every(value => Number.isSafeInteger(value) && value > 0))
      throw new DocumentError("InvalidOperation", "Workbook rules require supported positive safe-integer budgets.");
    if (!Number.isSafeInteger(this.maxNodes) || this.maxNodes < 1 || Object.keys(input).some(key => !["maxNodes", "read", "cells", "drawings", "annotations", "rules", "autoRows"].includes(key))
      || !surface.source.isConnected || !surface.view.isConnected || surface.source.ownerDocument !== surface.document || surface.view.ownerDocument !== surface.document
      || surface.view.childNodes.length) throw new DocumentError("InvalidOperation", "A workbook controller requires its own empty connected surface and supported node budget.");
    this.snapshot = providers.cells.snapshot;
    this.skewCoefficient = maximumWorkbookSkew(surface.document);
    checkWindowIdentity(providers.assets.snapshot.manifest, this.snapshot); checkWindowIdentity(providers.charts.snapshot.manifest, this.snapshot);
    if (providers.rules) checkWindowIdentity(providers.rules.snapshot.manifest, this.snapshot);
    this.options = Object.freeze({ ...input, read: Object.freeze({ ...input.read }), cells: Object.freeze({ ...input.cells }), annotations: Object.freeze({ ...input.annotations }), autoRows: Object.freeze({ ...input.autoRows }),
      drawings: Object.freeze({ ...input.drawings, ...(input.drawings?.assets ? { assets: Object.freeze({ ...input.drawings.assets }) } : {}) }) });
    // Validate read limits before retaining auxiliary content or starting work.
    const probe = new WorkbookWindowLoader(providers.cells, this.options.read, true); probe.dispose();
    this.content = new WorkbookContentSources(surface.document, this.snapshot);
  }

  get current(): WorkbookRenderedWindow | undefined { return this.active?.state; }

  /** Cancel only unfinished work while retaining the accepted window. */
  cancelPending(): void { this.pending?.cancel(); this.pending = undefined; }

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
    const abort = new AbortController(), loader = new WorkbookWindowLoader(this.providers.cells, this.options.read, true, this.skewCoefficient);
    let cancel!: () => void;
    const cancelled = new Promise<false>(resolve => { cancel = () => { abort.abort(); loader.dispose(); resolve(false); }; });
    const request = { cancel }; this.pending = request;
    const current = (): boolean => !this.disposed && this.pending === request && !abort.signal.aborted;
    let drawings: WorkbookDrawingLayout | undefined;
    let decoration: WorkbookWindowDecoration | undefined;
    try {
      let metrics = this.sheets.get(sheetId);
      if (!metrics) {
        const normal = measureWorkbookColumnFont(this.surface.source, this.snapshot.metadata.workbook, sheetId);
        const measured = await Promise.race([measureWorkbookAutoRows(this.surface, this.providers.cells, sheetId, normal, this.formatter, abort.signal, this.options.autoRows,
          this.providers.rules ? { source: this.providers.rules, limits: this.ruleLimits } : undefined), cancelled]);
        if (measured === false || !current()) return false;
        metrics = { normal, ...measured }; this.sheets.set(sheetId, metrics);
      }
      const { normal, dimensions } = metrics;
      if (!await Promise.race([loader.load(sheetId, dimensions, viewport), cancelled]) || !current()) return false;
      const frame = loader.frame!;
      let rules: WorkbookRuleReadResponse | undefined;
      if (sheet.conditionalRules?.length || sheet.dataValidations?.length) {
        if (!this.providers.rules) throw new DocumentError("InvalidOperation", "This worksheet requires an authorized current-value rule provider.");
        const request = Object.freeze({ documentId: frame.documentId, revision: frame.revision, sheetId, ...this.ruleLimits,
          coordinates: Object.freeze(frame.sourceCoordinates.map(point => Object.freeze({ row: point.row, column: point.column }))) });
        const response = await Promise.race([this.providers.rules.read(request, abort.signal), cancelled]);
        if (response === false || !current()) return false;
        rules = checkedWorkbookRules(response, request, this.snapshot);
      }
      const annotations = projectWorkbookAnnotations(this.snapshot, frame, dimensions, this.options.annotations);
      const cells = layoutWorkbookWindow(this.surface, this.snapshot, frame, dimensions, this.formatter,
        { ...this.options.cells, normalSpaceWidthPx: normal.spaceWidthPx, maxNodes: Math.min(this.maxNodes, this.options.cells?.maxNodes ?? this.maxNodes),
          maxCellFragments: Math.min(this.options.read?.maxCells ?? 5000, this.options.cells?.maxCellFragments ?? this.options.read?.maxCells ?? 5000) }, rules);
      const walker = this.surface.document.createTreeWalker(cells.element, 0xffff_ffff); let nodes = 1;
      while (walker.nextNode()) if (++nodes >= this.maxNodes) throw new DocumentError("ResourceLimitExceeded", "The combined workbook window exceeds its node budget.");
      if (nodes >= this.maxNodes) throw new DocumentError("ResourceLimitExceeded", "The combined workbook window has no remaining object-layer budget.");
      drawings = await layoutWorkbookDrawings(this.surface, this.snapshot, frame, dimensions, this.content, this.providers.assets, this.providers.charts,
        this.formatter, abort.signal, { ...this.options.drawings, maxNodes: Math.min(this.maxNodes - nodes, this.options.drawings?.maxNodes ?? this.maxNodes) });
      if (!current()) return false;
      nodes += drawings.nodes; cells.element.append(drawings.element);
      const base: WorkbookRenderedWindow = Object.freeze({ frame, element: cells.element, nodes, dimensions, cells: cells.cells, objects: drawings.objects, annotations,
        fonts: Object.freeze([...normal.fonts, ...metrics.fonts, ...cells.fonts, ...drawings.fonts]), diagnostics: Object.freeze([...drawings.diagnostics, ...rules ? workbookRuleDiagnostics(rules) : []]), ...(rules ? { rules } : {}),
        usedFallbackLineMetrics: metrics.usedFallbackLineMetrics || cells.usedFallbackLineMetrics || drawings.usedFallbackLineMetrics });
      decoration = this.decorate?.(base, this.content, this.maxNodes - nodes);
      if (decoration && (!Number.isSafeInteger(decoration.nodes) || decoration.nodes < 0 || nodes + decoration.nodes > this.maxNodes))
        throw new DocumentError("ResourceLimitExceeded", "Workbook controls exceed their assigned window-node budget.");
      if (!current()) return false;
      const extra = decoration;
      const state: WorkbookRenderedWindow = extra ? Object.freeze({ ...base, get nodes() { return base.nodes + extra.nodes; },
        get fonts() { return extra.fonts?.length ? Object.freeze([...base.fonts, ...extra.fonts]) : base.fonts; } }) : base;
      const previous = this.active;
      this.surface.view.replaceChildren(cells.element);
      this.surface.frame.style.setProperty("width", viewport.widthPt + "pt", "important");
      this.surface.frame.style.setProperty("height", viewport.heightPt + "pt", "important");
      this.surface.frame.style.setProperty("visibility", "visible", "important");
      this.active = { state, drawings, ...(decoration ? { decoration } : {}) }; drawings = undefined; decoration = undefined;
      previous?.decoration?.dispose(); previous?.drawings.dispose(); previous?.state.element.replaceChildren();
      return true;
    } catch (error) {
      if (!current()) return false;
      throw error;
    } finally {
      decoration?.dispose(); drawings?.dispose(); loader.dispose();
      if (this.pending === request) this.pending = undefined;
    }
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true; this.pending?.cancel(); this.pending = undefined;
    this.active?.decoration?.dispose(); this.active?.drawings.dispose(); this.active?.state.element.remove(); this.active?.state.element.replaceChildren(); this.active = undefined;
    this.content.dispose(); this.sheets.clear();
  }
}
