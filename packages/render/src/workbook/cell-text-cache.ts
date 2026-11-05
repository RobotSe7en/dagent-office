import { DocumentError } from "@dagent-office/core";
import type { WorkbookNumberFormatter, WorkbookSnapshot } from "@dagent-office/core";
import type { FontResolution } from "../contracts.js";
import { markViewNode, viewNodeRole } from "../view-nodes.js";
import { layoutWorkbookCellTexts } from "./cell-text-layout.js";
import type { WorkbookCellTextInput, WorkbookCellTextLayouts } from "./cell-text-layout.js";
import type { WorkbookTextLayout } from "./text-layout.js";

type Cell = WorkbookCellTextLayouts["cells"][number];
interface Entry { readonly cell: Cell; readonly fonts: readonly FontResolution[]; readonly fallback: boolean; readonly nodes: number; readonly units: number }
function cloneText(text: WorkbookTextLayout): WorkbookTextLayout {
  const document = text.element.ownerDocument, element = text.element.cloneNode(true) as HTMLElement;
  const originals = document.createTreeWalker(text.element, 0xffff_ffff), copies = document.createTreeWalker(element, 0xffff_ffff), nodes = new Map<Node, Node>();
  let original: Node | null = text.element, copy: Node | null = element;
  while (original && copy) {
    nodes.set(original, copy); const role = viewNodeRole(original); if (role) markViewNode(copy, role);
    original = originals.nextNode(); copy = copies.nextNode();
  }
  const root = nodes.get(text.text.element) as HTMLElement | undefined;
  if (!root || original || copy) throw new DocumentError("InvalidOperation", "A cached workbook text tree cannot be projected.");
  return Object.freeze({ ...text, element, text: Object.freeze({ ...text.text, element: root, projections: Object.freeze(text.text.projections.map(part => {
    const node = nodes.get(part.node); if (!node || node.nodeType !== 3) throw new DocumentError("InvalidOperation", "Cached worksheet source offsets have no text node.");
    return Object.freeze({ node: node as Text, from: part.from, to: part.to });
  })) }) });
}

/** One immutable controller owns a bounded detached-text LRU. Store clones so
 * a cached node never retains an old pane through its DOM parent pointers. */
export class WorkbookCellTextCache {
  private readonly entries = new Map<string, Entry>();
  private readonly fontIds = new WeakMap<FontResolution, string>();
  private readonly faces = new WeakMap<FontFace, number>();
  private faceId = 0;
  private fontState: string | undefined;
  private nodes = 0; private units = 0;
  constructor(private readonly maxEntries: number, private readonly maxNodes: number, private readonly maxUnits: number) {
    if (![maxEntries, maxNodes, maxUnits].every(value => Number.isSafeInteger(value) && value > 0))
      throw new DocumentError("InvalidOperation", "Workbook text caching requires positive finite budgets.");
  }
  settled(fonts: FontFaceSet): boolean {
    const state = JSON.stringify([...fonts].map(face => {
      let id = this.faces.get(face); if (id === undefined) { id = ++this.faceId; this.faces.set(face, id); }
      return [id, face.family, face.style, face.weight, face.stretch, face.unicodeRange, face.featureSettings, face.variationSettings,
        face.ascentOverride, face.descentOverride, face.lineGapOverride, face.display, face.status];
    }));
    // Reading fonts.ready can itself force layout for every newly projected
    // source table. Face identity/descriptors/status also detect completed
    // loads between frames, without that synchronous layout checkpoint.
    if (this.fontState !== state || fonts.status === "loading") this.clear();
    this.fontState = state;
    return fonts.status !== "loading";
  }
  fontId(font: FontResolution): string { let key = this.fontIds.get(font); if (!key) { key = JSON.stringify(font); this.fontIds.set(font, key); } return key; }
  get(key: string): Entry | undefined {
    const entry = this.entries.get(key); if (!entry?.cell.text) return undefined;
    this.entries.delete(key); this.entries.set(key, entry);
    return { ...entry, cell: Object.freeze({ ...entry.cell, text: cloneText(entry.cell.text) }) };
  }
  put(key: string, cell: Cell, fonts: readonly FontResolution[], fallback: boolean): void {
    if (!cell.text) return;
    const nodes = cell.text.text.nodes + 1, units = key.length + cell.text.element.outerHTML.length + fonts.reduce((sum, font) => sum + this.fontId(font).length, 0);
    if (nodes > this.maxNodes || units > this.maxUnits) return;
    const previous = this.entries.get(key); if (previous) { this.nodes -= previous.nodes; this.units -= previous.units; this.entries.delete(key); }
    while (this.entries.size && (this.entries.size >= this.maxEntries || this.nodes + nodes > this.maxNodes || this.units + units > this.maxUnits)) {
      const first = this.entries.keys().next().value!, old = this.entries.get(first)!; this.entries.delete(first); this.nodes -= old.nodes; this.units -= old.units;
    }
    this.entries.set(key, { cell: Object.freeze({ ...cell, text: cloneText(cell.text) }), fonts, fallback, nodes, units }); this.nodes += nodes; this.units += units;
  }
  clear(): void { this.entries.clear(); this.nodes = 0; this.units = 0; }
}

export function layoutCachedWorkbookCellTexts(scope: HTMLElement, snapshot: WorkbookSnapshot, inputs: readonly WorkbookCellTextInput[], keys: readonly string[],
  formatter: WorkbookNumberFormatter, options: { readonly normalSpaceWidthPx: number }, maxNodes: number, maxUnits: number, cache: WorkbookCellTextCache): WorkbookCellTextLayouts {
  const entries = cache.settled(scope.ownerDocument.fonts) ? keys.map(key => cache.get(key)) : inputs.map(() => undefined);
  let nodes = 0, units = 0, fallback = false;
  const fonts = new Map<string, FontResolution>(), missing: WorkbookCellTextInput[] = [], indices = new Map<number, number>();
  for (const [index, entry] of entries.entries()) {
    if (!entry?.cell.text) { indices.set(index, missing.length); missing.push(inputs[index]!); continue; }
    nodes += entry.cell.text.text.nodes + 1; units += entry.cell.text.text.element.textContent?.length ?? 0; fallback ||= entry.fallback;
    for (const font of entry.fonts) fonts.set(cache.fontId(font), font);
  }
  if (nodes >= maxNodes || units > maxUnits) throw new DocumentError("ResourceLimitExceeded", "Cached workbook text exceeds the complete window budget.");
  const fresh = missing.length ? layoutWorkbookCellTexts(scope, snapshot, missing, formatter, options, maxNodes - nodes, Math.max(1, maxUnits - units))
    : { cells: [], nodes: 0, units: 0, fonts: [], usedFallbackLineMetrics: false };
  nodes += fresh.nodes; units += fresh.units; fallback ||= fresh.usedFallbackLineMetrics;
  if (nodes >= maxNodes || units > maxUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook text exceeds its cumulative cache and layout budget.");
  for (const font of fresh.fonts) fonts.set(cache.fontId(font), font);
  const cells = inputs.map((_, index) => {
    const entry = entries[index]; if (entry) return entry.cell;
    const cell = fresh.cells[indices.get(index)!]!; cache.put(keys[index]!, cell, fresh.fonts, fresh.usedFallbackLineMetrics); return cell;
  });
  return Object.freeze({ nodes, units, cells: Object.freeze(cells), fonts: Object.freeze([...fonts.values()]), usedFallbackLineMetrics: fallback });
}
