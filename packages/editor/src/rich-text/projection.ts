/** Reversible projection: layout, font fallback, assets and labels never become source. */
export class SourceProjection {
  private readonly originals = new WeakMap<Element, Element>();
  private readonly rendered = new WeakMap<Element, Element>();
  private readonly derived = new WeakSet<Node>();
  private readonly identities = new Map<string, { original: Element; rendered: Element }>();
  readonly ignored = (node: Node): boolean => this.derived.has(node);

  markDerived(node: Node): void { this.derived.add(node); }
  capture(root: HTMLElement, source: HTMLElement): void {
    const originals = new Map([source, ...source.querySelectorAll<HTMLElement>("[data-node-id]")].map((node) => [node.dataset.nodeId, node]));
    for (const node of [root, ...root.querySelectorAll<HTMLElement>("*")]) {
      const original = originals.get(node.dataset.nodeId);
      if (!original) {
        // HTML parsers may introduce an identity-free tbody. It is source structure,
        // whereas numbering labels and measured grid columns are renderer children.
        if (node.matches(".sdk-number-marker,colgroup,.sdk-story-label")) this.derived.add(node);
        continue;
      }
      const authored = original.cloneNode(original.hasAttribute("data-word-tab")) as Element, view = node.cloneNode(false) as Element;
      this.originals.set(node, authored); this.rendered.set(node, view);
      this.identities.set(node.dataset.nodeId!, { original: authored, rendered: view });
    }
  }
  serialize(root: HTMLElement): HTMLElement {
    const seen = new Set<string>();
    const clone = (node: Node): Node | undefined => {
      if (this.ignored(node)) return;
      if (node.nodeType !== 1) return node.cloneNode(false);
      const element = node as Element, result = element.cloneNode(false) as HTMLElement;
      // Range.extractContents clones partially selected ancestors; their attributes
      // still come from the same source identity and need the same reversal.
      const identity = this.identities.get(element.getAttribute("data-node-id") ?? "");
      const original = this.originals.get(element) ?? identity?.original, rendered = this.rendered.get(element) ?? identity?.rendered;
      if (original && rendered) {
        for (const name of new Set([...element.attributes, ...rendered.attributes, ...original.attributes].map((item) => item.name))) {
          if (name === "style") continue;
          if (element.getAttribute(name) !== rendered.getAttribute(name)) continue;
          const value = original.getAttribute(name);
          if (value === null) result.removeAttribute(name); else result.setAttribute(name, value);
        }
        const before = (original as HTMLElement).style, view = (rendered as HTMLElement).style;
        // CSSOM expands shorthands and exposes empty longhands for var(...).
        // Restoring an unchanged style as text preserves those source declarations.
        if (element.getAttribute("style") === rendered.getAttribute("style")) {
          const value = original.getAttribute("style");
          if (value === null) result.removeAttribute("style"); else result.setAttribute("style", value);
        } else {
          for (const property of new Set([...view, ...before])) {
            if (result.style.getPropertyValue(property) !== view.getPropertyValue(property)) continue;
            const value = before.getPropertyValue(property);
            if (value) result.style.setProperty(property, value, before.getPropertyPriority(property)); else result.style.removeProperty(property);
          }
          if (!result.style.length) {
            if (original.hasAttribute("style")) result.setAttribute("style", ""); else result.removeAttribute("style");
          }
        }
      }
      const id = result.getAttribute("data-node-id");
      if (id) { if (seen.has(id)) { result.removeAttribute("data-node-id"); result.removeAttribute("id"); } else seen.add(id); }
      if (original && element.hasAttribute("data-word-tab")) {
        result.append(...[...original.childNodes].map((child) => child.cloneNode(true)));
      } else for (const child of node.childNodes) { const value = clone(child); if (value) result.append(value); }
      return result;
    };
    return clone(root) as HTMLElement;
  }
}
