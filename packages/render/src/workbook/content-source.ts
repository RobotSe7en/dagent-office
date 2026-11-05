import { DocumentError } from "@dagent-office/core";
import type { WorkbookSnapshot } from "@dagent-office/core";
import { packagePath } from "../preview-package.js";
import { markViewNode } from "../view-nodes.js";

export interface WorkbookContentProjection {
  readonly nodes: number;
  readonly assets: readonly { readonly nodeId: string; readonly path: string }[];
  get(nodeId: string): HTMLElement | undefined;
  dispose(): void;
}
const tags = new Set(["article", "section", "div", "p", "span", "strong", "em", "u", "s", "sub", "sup", "br", "img"]);
const globals = new Set(["id", "class", "title", "lang", "dir", "style"]);
const imageAttributes = new Set(["src", "alt", "width", "height"]);
function invalid(): never { throw new DocumentError("InvalidDocument", "Workbook auxiliary HTML does not match its validated source structure."); }

/** An inert index of already core-validated auxiliary HTML, never cell chunks.
 * Strip image URLs while nodes still belong to the template document, before
 * cloning/adoption can start a fetch. Source CSS is core's resource-free CSS;
 * the owning preview surface supplies that validated stylesheet separately.
 * Only requested object/story subtrees are connected for CSS measurement. */
export class WorkbookContentSources {
  readonly identity: Readonly<{ documentId: string; revision: number }>;
  private root: HTMLElement | undefined;
  private readonly roots = new Map<string, HTMLElement>();
  private readonly images = new Map<string, string>();
  private readonly active = new Set<WorkbookContentProjection>();
  private readonly themeCssClass: string;
  private disposed = false;
  constructor(private readonly document: Document, snapshot: WorkbookSnapshot, maxNodes: number = 100_000, maxCodeUnits: number = 8 * 1024 * 1024) {
    this.identity = Object.freeze({ documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision });
    this.themeCssClass = snapshot.metadata.workbook.themeCssClass ?? "";
    if (![maxNodes, maxCodeUnits].every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Auxiliary source budgets must be positive safe integers.");
    const content = snapshot.content;
    if (!content) return;
    if (content.path !== snapshot.manifest.content) invalid();
    if (content.html.length > maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook auxiliary HTML exceeds its source budget.");
    const template = document.createElement("template"); template.innerHTML = content.html;
    const candidates = template.content.querySelectorAll<HTMLElement>('[data-office-kind="workbook-content"]');
    if (candidates.length !== 1) invalid();
    const root = candidates[0]!;
    if (!root.dataset.nodeId) invalid();
    const stack = [{ node: root as Node, depth: 0 }], ids = new Set<string>(), assets = new Set(snapshot.manifest.assets.map(asset => asset.path));
    let count = 0;
    while (stack.length) {
      const { node, depth } = stack.pop()!;
      if (++count > maxNodes || depth > 64) throw new DocumentError("ResourceLimitExceeded", "Workbook auxiliary HTML exceeds its node or depth budget.");
      if (node.nodeType === 3) continue;
      if (node.nodeType !== 1) invalid();
      const element = node as HTMLElement, id = element.dataset.nodeId;
      if (!id || ids.has(id) || element.namespaceURI !== "http://www.w3.org/1999/xhtml" || !tags.has(element.localName)) invalid();
      ids.add(id!);
      for (const attribute of element.attributes) if (attribute.namespaceURI || !globals.has(attribute.name) && !/^(?:data-[a-z0-9-]+|aria-[a-z-]+)$/.test(attribute.name) && !(element.localName === "img" && imageAttributes.has(attribute.name))) invalid();
      if (element.localName === "img") {
        const src = element.getAttribute("src"); if (!src) invalid();
        let path: string;
        try { path = packagePath(src!, content.path); } catch { return invalid(); }
        if (!assets.has(path)) invalid();
        this.images.set(id!, path); element.removeAttribute("src");
      }
      for (const child of [...node.childNodes].reverse()) stack.push({ node: child, depth: depth + 1 });
    }
    const declared = new Set([...(snapshot.metadata.workbook.objects ?? []).map(object => object.nodeId), ...(snapshot.metadata.workbook.stories ?? []).map(story => story.rootNodeId)]);
    for (const node of root.childNodes) {
      if (node.nodeType === 3 && !node.textContent?.trim()) continue;
      const element = node as HTMLElement, id = element.dataset?.nodeId;
      if (node.nodeType !== 1 || !id || !declared.delete(id)) invalid();
      this.roots.set(id!, element);
    }
    if (declared.size) invalid();
    this.root = root;
  }
  private check(): void { if (this.disposed) throw new DocumentError("InvalidOperation", "This workbook auxiliary source has been disposed."); }

  project(scope: HTMLElement, rootIds: readonly string[], maxNodes = 50_000): WorkbookContentProjection {
    this.check();
    if (!scope.isConnected || scope.ownerDocument !== this.document || !this.document.defaultView || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "Auxiliary measurement requires a connected source surface and a positive node budget.");
    if (rootIds.length > 5000) throw new DocumentError("ResourceLimitExceeded", "An auxiliary projection exceeds its root budget.");
    const selected = new Set(rootIds);
    if (selected.size !== rootIds.length || rootIds.some(id => !this.roots.has(id))) throw new DocumentError("InvalidOperation", "Auxiliary projection roots must be unique declared objects or stories.");
    const wrapper = markViewNode(this.document.createElement("office-workbook-source")), byId = new Map<string, HTMLElement>(), assets: { nodeId: string; path: string }[] = [];
    wrapper.dataset.officeView = "workbook-content-measurement"; wrapper.setAttribute("aria-hidden", "true");
    wrapper.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;overflow:hidden;left:0;top:0;width:0;height:0";
    wrapper.className = this.themeCssClass;
    const root = this.root?.cloneNode(false) as HTMLElement | undefined;
    let nodes = root ? 2 : 1;
    if (root) {
      byId.set(root.dataset.nodeId!, root); wrapper.append(root);
      // Preserve source order, regardless of selection/request order.
      for (const [id, source] of this.roots) if (selected.has(id)) {
        const stack = [source as Node];
        while (stack.length) { const node = stack.pop()!; if (++nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook auxiliary projection exceeds its node budget."); stack.push(...node.childNodes); }
        const clone = source.cloneNode(true) as HTMLElement;
        for (const element of [clone, ...clone.querySelectorAll<HTMLElement>("[data-node-id]")]) {
          const nodeId = element.dataset.nodeId!, path = this.images.get(nodeId); byId.set(nodeId, element);
          if (path) assets.push(Object.freeze({ nodeId, path }));
        }
        root.append(clone);
      }
    }
    if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook auxiliary projection exceeds its node budget.");
    let disposed = false;
    const projection: WorkbookContentProjection = Object.freeze({ nodes, assets: Object.freeze(assets), get(nodeId: string) {
      if (disposed) throw new DocumentError("InvalidOperation", "This workbook auxiliary projection has been disposed.");
      return byId.get(nodeId);
    }, dispose: () => { if (!disposed) { disposed = true; wrapper.remove(); wrapper.replaceChildren(); byId.clear(); this.active.delete(projection); } } });
    scope.append(wrapper); this.active.add(projection); return projection;
  }
  dispose(): void {
    if (this.disposed) return;
    for (const projection of this.active) projection.dispose();
    this.root?.replaceChildren(); this.root = undefined; this.roots.clear(); this.images.clear(); this.disposed = true;
  }
}
