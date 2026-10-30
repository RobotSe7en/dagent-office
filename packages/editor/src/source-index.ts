import { DocumentError } from "@dagent-office/core";
import type { Anchor, DeepReadonly, DocumentModel, Rectangle, TextRange } from "@dagent-office/core";

interface SourceElement {
  readonly element: HTMLElement;
  readonly from: number;
  readonly to: number;
}
interface Token {
  readonly node: Node;
  readonly text: string;
  readonly from: number;
  readonly to: number;
  readonly owner: HTMLElement;
}
const blocks = "p,h1,h2,h3,h4,h5,h6,li,td,th";
const hitTargets = blocks + ",img,table";

function boundary(text: string, offset: number): boolean {
  const before = text.charCodeAt(offset - 1), after = text.charCodeAt(offset), code = text.codePointAt(offset);
  return !(before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff)
    && !(offset > 0 && code !== undefined && (code >= 0xfe00 && code <= 0xfe0f || code >= 0xe0100 && code <= 0xe01ef));
}
function intersects(a: Rectangle, b: Rectangle): boolean {
  return a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;
}

/** Capture source positions before the renderer inserts labels or replaces tab text. */
export class PreviewSource {
  readonly elements: Map<string, SourceElement> = new Map();
  readonly tokens: Token[] = [];
  readonly root: HTMLElement;
  private readonly document: Document;

  constructor(document: Document, model: DocumentModel) {
    this.document = document;
    const article = document.querySelector<HTMLElement>("article[data-node-id]")!;
    let offset = 0;
    const visit = (node: Node): void => {
      const from = offset;
      const element = node.nodeType === 1 ? node as HTMLElement : undefined;
      const atomic = element?.matches("br,img,[data-word-tab]");
      const text = node.nodeType === 3 ? node.nodeValue ?? ""
        : atomic ? element!.tagName === "IMG" ? "\ufffc" : element!.tagName === "BR" ? "\n" : element!.textContent ?? "" : undefined;
      if (text !== undefined) {
        offset += text.length;
        const parent = element ?? node.parentElement!;
        const owner = parent.closest<HTMLElement>(blocks);
        // An un-identified paragraph cannot occur in a validated package.
        const identified = owner?.hasAttribute("data-node-id") ? owner : parent.closest<HTMLElement>("[data-node-id]")!;
        this.tokens.push({ node, text, from, to: offset, owner: identified });
      } else {
        for (const child of node.childNodes) visit(child);
      }
      if (element?.dataset.nodeId) this.elements.set(element.dataset.nodeId, { element, from, to: offset });
    };
    visit(article);
    const main = model.snapshot().metadata.word?.stories.find((story) => story.kind === "main");
    this.root = main ? this.elements.get(main.rootNodeId)!.element : article;
  }

  /** Only the main content story participates in document selection. */
  contains(node: Node): boolean { return this.root === node || this.root.contains(node); }

  private tokenRange(token: Token, from = 0, to = token.text.length): Range {
    const range = this.document.createRange();
    if (token.node.nodeType === 3) {
      range.setStart(token.node, from);
      range.setEnd(token.node, to);
    } else range.selectNode(token.node);
    return range;
  }
  private visible(token: Token): boolean {
    const element = token.node.nodeType === 1 ? token.node as HTMLElement : token.node.parentElement!;
    const style = this.document.defaultView!.getComputedStyle(element);
    return token.node.isConnected && style.visibility === "visible" && this.tokenRange(token).getClientRects().length > 0;
  }

  textAnchor(range: Range, model: DocumentModel): Anchor | undefined {
    if (range.collapsed) return undefined;
    if (!this.contains(range.startContainer) || !this.contains(range.endContainer))
      throw new DocumentError("InvalidOperation", "文字选择必须位于同一正文内容域。");
    const ranges: TextRange[] = [];
    for (const token of this.tokens) {
      if (!this.contains(token.node) || !this.visible(token) || !range.intersectsNode(token.node)) continue;
      const from = token.node.nodeType === 3 && range.startContainer === token.node ? range.startOffset : 0;
      const to = token.node.nodeType === 3 && range.endContainer === token.node ? range.endOffset : token.text.length;
      if (from === to) continue;
      const nodeId = token.owner.dataset.nodeId!;
      const source = this.elements.get(nodeId)!;
      const start = token.from - source.from + from, end = token.from - source.from + to;
      const text = model.getNode(nodeId)!.text;
      if (!boundary(text, start) || !boundary(text, end))
        throw new DocumentError("InvalidOperation", "选择不能拆开 UTF-16 代理对或字符的变体选择符。");
      const previous = ranges.at(-1);
      if (previous?.nodeId === nodeId && previous.to === start) {
        previous.to = end;
        previous.quote = text.slice(previous.from, end);
      } else ranges.push({ nodeId, from: start, to: end, quote: text.slice(start, end) });
    }
    return ranges.length === 1 ? { kind: "text", ...ranges[0]! } : ranges.length ? { kind: "text-set", ranges } : undefined;
  }

  orderedIds(ids: readonly string[]): string[] {
    const wanted = new Set(ids);
    return [this.root, ...this.root.querySelectorAll<HTMLElement>("[data-node-id]")]
      .map((element) => element.dataset.nodeId!).filter((id) => wanted.has(id));
  }
  scale(): number {
    return this.root.getBoundingClientRect().width / this.root.offsetWidth || 1;
  }
  point(clientX: number, clientY: number): { x: number; y: number } {
    const rect = this.root.getBoundingClientRect(), scale = this.scale();
    return {
      x: Math.max(0, Math.min(rect.width, clientX - rect.left)) * 72 / 96 / scale,
      y: Math.max(0, Math.min(rect.height, clientY - rect.top)) * 72 / 96 / scale,
    };
  }
  region(rect: Rectangle, layoutKey: string): Anchor {
    const surface = this.root.getBoundingClientRect(), scale = this.scale();
    const hits = [...this.root.querySelectorAll<HTMLElement>("[data-node-id]")].filter((element) => {
      if (!element.matches(hitTargets)) return false;
      return [...element.getClientRects()].some((box) => intersects(rect, {
        x: (box.left - surface.left) * 72 / 96 / scale,
        y: (box.top - surface.top) * 72 / 96 / scale,
        width: box.width * 72 / 96 / scale, height: box.height * 72 / 96 / scale,
      }));
    });
    return { kind: "region", surfaceId: this.root.dataset.nodeId!, rect,
      nodeIds: hits.filter((element) => !hits.some((child) => child !== element && element.contains(child))).map((element) => element.dataset.nodeId!), layoutKey };
  }

  rectangles(anchor: DeepReadonly<Anchor>): DOMRect[] {
    if (anchor.kind === "region") {
      const surface = this.elements.get(anchor.surfaceId)?.element;
      if (!surface) return [];
      const box = surface.getBoundingClientRect(), scale = surface.getBoundingClientRect().width / surface.offsetWidth || 1;
      return [new DOMRect(box.left + anchor.rect.x * 96 / 72 * scale, box.top + anchor.rect.y * 96 / 72 * scale,
        anchor.rect.width * 96 / 72 * scale, anchor.rect.height * 96 / 72 * scale)];
    }
    if (anchor.kind === "node" || anchor.kind === "node-set") {
      const ids = anchor.kind === "node" ? [anchor.nodeId] : anchor.nodeIds;
      return ids.flatMap((id) => [...this.elements.get(id)?.element.getClientRects() ?? []]);
    }
    const ranges = anchor.kind === "text" ? [anchor] : anchor.ranges;
    return ranges.flatMap((range) => {
      const source = this.elements.get(range.nodeId);
      if (!source) return [];
      const start = source.from + range.from, end = source.from + range.to;
      return this.tokens.flatMap((token) => {
        if (token.to <= start || token.from >= end || !this.visible(token)) return [];
        return [...this.tokenRange(token, Math.max(0, start - token.from), Math.min(token.text.length, end - token.from)).getClientRects()];
      });
    });
  }

  /** Source-relative geometry and resolved faces; scrolling and SDK zoom are excluded. */
  layoutKey(): string {
    const view = this.document.defaultView!;
    const surface = this.root.getBoundingClientRect();
    let hash = 2166136261;
    const add = (value: string): void => {
      for (let index = 0; index < value.length; index++) hash = Math.imul(hash ^ value.charCodeAt(index), 16777619);
    };
    add("continuous-v1");
    for (const [id, { element }] of this.elements) {
      if (!this.contains(element)) continue;
      const style = view.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      add(JSON.stringify([id, rect.left - surface.left, rect.top - surface.top, rect.width, rect.height,
        style.fontFamily, style.fontSize, style.fontWeight, style.fontStyle, style.lineHeight, style.letterSpacing]));
    }
    for (const token of this.tokens) {
      if (!this.contains(token.node) || !this.visible(token)) continue;
      for (const rect of this.tokenRange(token).getClientRects())
        add(JSON.stringify([rect.left - surface.left, rect.top - surface.top, rect.width, rect.height]));
    }
    return "continuous-v1-" + (hash >>> 0).toString(16);
  }
}
