// Only ownership roles are shared. Weak keys retain neither documents nor source
// content, and copied HTML cannot claim ownership by supplying an attribute.
const roles = new WeakMap<Node, "remove" | "unwrap">();
const selectionHidden = new WeakSet<Node>();
const measured = new WeakSet<Node>();
export function measureViewContent(node: Node): void { measured.add(node); }
export function isMeasuredViewContent(node: Node): boolean { return measured.has(node); }
export function hideFromViewSelection(node: Node): void { selectionHidden.add(node); }
export function hiddenFromViewSelection(node: Node): boolean { return selectionHidden.has(node); }
const attributes = new WeakMap<Element, Set<string>>();
const styles = new WeakMap<Element, Set<string>>();
const styleSources = new WeakMap<Element, Element>();
export function ownViewAttribute(node: Element, name: string): void {
  let names = attributes.get(node); if (!names) { names = new Set(); attributes.set(node, names); } names.add(name);
}
export function ownViewStyle(node: Element, ...names: string[]): void {
  if (!styleSources.has(node)) styleSources.set(node, node.cloneNode(false) as Element);
  let properties = styles.get(node); if (!properties) { properties = new Set(); styles.set(node, properties); }
  for (const name of names) properties.add(name);
}
export function ownedViewAttribute(node: Element, name: string): boolean { return attributes.get(node)?.has(name) ?? false; }
export function ownedViewStyle(node: Element, name: string): boolean { return styles.get(node)?.has(name) ?? false; }
export function viewStyleSource(node: Element): Element | undefined { return styleSources.get(node); }
export function markViewNode<ElementType extends Node>(node: ElementType, role: "remove" | "unwrap" = "remove"): ElementType {
  roles.set(node, role); return node;
}
export function viewNodeRole(node: Node): "remove" | "unwrap" | undefined { return roles.get(node); }
