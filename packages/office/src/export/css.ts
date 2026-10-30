import { generate, ident, lexer, List, parse, walk } from "css-tree";
import type { CssNode, Declaration, Selector } from "css-tree";
import { DocumentError } from "@dagent-office/core";
import type { DocumentReadLimits } from "@dagent-office/core";
import type { SourceElement } from "./source.js";
import { length } from "./values.js";

export type Css = Readonly<Record<string, string>>;
interface Compound { tag?: string; classes: string[]; attrs: { name: string; value?: string }[]; root: boolean }
type Specificity = readonly [number, number, number];
interface Rule { compounds: Compound[]; declarations: Declaration[]; specificity: Specificity; order: number; key: string; revisionDisplay: boolean }
interface Entry { value: string; important: boolean; specificity: Specificity; order: number }
const inherited = new Set(["color", "font-family", "font-size", "font-weight", "font-style", "font-variant", "text-transform", "text-align", "text-indent", "line-height", "letter-spacing", "word-spacing", "white-space", "widows", "orphans", "tab-size", "overflow-wrap", "list-style-type", "list-style-position"]);
const inline = new Set(["span", "a", "strong", "b", "em", "i", "u", "s", "sub", "sup", "ins", "del", "code"]);
const initial: Css = { "margin-top": "0pt", "margin-right": "0pt", "margin-bottom": "0pt", "margin-left": "0pt", color: "#000000", "font-family": '"Noto Sans CJK SC",serif', "font-size": "12pt", "font-weight": "normal", "font-style": "normal", "font-variant": "normal", "text-transform": "none", "text-align": "left", "text-indent": "0pt", "line-height": "normal", "letter-spacing": "normal", "word-spacing": "normal", "white-space": "normal", "vertical-align": "baseline", "background-color": "transparent", "text-decoration-line": "none", "text-decoration-style": "solid", "text-decoration-color": "currentColor", widows: "2", orphans: "2", "break-before": "auto", "break-after": "auto", "break-inside": "auto" };
function valueString(nodes: CssNode[]): string { return generate({ type: "Value", children: new List<CssNode>().fromArray(nodes) }); }
/** Resolve CSS variables as values, never by executing CSS or JavaScript. */
export function variables(value: string, get: (name: string) => string | undefined, depth = 0, maximum: number = 8 * 1024 * 1024): string | undefined {
  if (depth > 128) throw new DocumentError("ResourceLimitExceeded", "CSS variable nesting exceeds the DOCX budget.");
  let size = value.length;
  if (size > maximum) throw new DocumentError("ResourceLimitExceeded", "Expanded CSS exceeds the DOCX value budget.");
  const ast = parse(value, { context: "value", parseCustomProperty: true }); let valid = true;
  walk(ast, function (node, item, list) {
    if (node.type !== "Function" || ident.decode(node.name).toLowerCase() !== "var") return;
    const nodes = node.children.toArray(), name = nodes[0];
    if (name?.type !== "Identifier") { valid = false; return this.skip; }
    let replacement = get(ident.decode(name.name));
    if (replacement === undefined) {
      const comma = nodes.findIndex((child) => child.type === "Operator" && child.value === ",");
      if (comma >= 0) replacement = variables(valueString(nodes.slice(comma + 1)), get, depth + 1, maximum);
    }
    if (replacement === undefined) valid = false;
    else {
      size += replacement.length;
      if (size > maximum) throw new DocumentError("ResourceLimitExceeded", "Expanded CSS exceeds the DOCX value budget.");
      const parsed = parse(replacement, { context: "value", parseCustomProperty: true });
      if (parsed.type !== "Value" || !item || !list) valid = false; else list.replace(item, parsed.children);
    }
    return this.skip;
  });
  return valid ? generate(ast) : undefined;
}
function selector(source: Selector): { compounds: Compound[]; specificity: Specificity; key: string } {
  const compounds: Compound[] = [{ classes: [], attrs: [], root: false }]; const specificity: [number, number, number] = [0, 0, 0];
  for (const part of source.children) {
    let compound = compounds.at(-1)!;
    if (part.type === "Combinator") { compounds.push({ classes: [], attrs: [], root: false }); continue; }
    if (part.type === "TypeSelector") { compound.tag = ident.decode(part.name).toLowerCase(); specificity[2]++; }
    else if (part.type === "ClassSelector") { compound.classes.push(ident.decode(part.name)); specificity[1]++; }
    else if (part.type === "AttributeSelector") {
      compound.attrs.push({ name: ident.decode(part.name.name), ...(part.value ? { value: part.value.type === "String" ? part.value.value : ident.decode(part.value.name) } : {}) }); specificity[1]++;
    } else if (part.type === "PseudoClassSelector") { compound.root = true; specificity[1]++; }
    else throw new DocumentError("UnsupportedFeature", "DOCX export encountered an unsupported selector.");
  }
  const last = compounds.at(-1)!, key = last.classes[0] ? "." + last.classes[0] : last.attrs[0] ? "[" + last.attrs[0].name : last.tag ?? (last.root ? ":root" : "*");
  return { compounds, specificity, key };
}
function tokens(value: string): string[] {
  const ast = parse(value, { context: "value", parseCustomProperty: true }); return ast.type === "Value" ? ast.children.toArray().map((node) => generate(node)) : [];
}
function expand(property: string, value: string): [string, string][] {
  if (["inherit", "initial", "unset", "revert"].includes(value)
    && (/^(margin|padding|border(?:-(?:width|style|color|top|right|bottom|left))?)$/.test(property) || property === "text-decoration"))
    return expand(property, property === "margin" || property === "padding" || /^border-(width|style|color)$/.test(property) ? "0" : "none").map(([name]) => [name, value]);
  const parts = tokens(value), box = /^(margin|padding|border-(?:width|style|color))$/.exec(property);
  if (box && parts.length >= 1 && parts.length <= 4) {
    const [a, b = a, c = a, d = b] = parts;
    return ["top", "right", "bottom", "left"].map((side, index) => [property.startsWith("border-") ? "border-" + side + property.slice(6) : property + "-" + side, [a, b, c, d][index]!] as [string, string]);
  }
  if (property === "border" || /^border-(top|right|bottom|left)$/.test(property)) {
    let width = "medium", style = "none", color = "currentColor";
    for (const part of parts) {
      if (/^(none|hidden|solid|double|dotted|dashed|groove|ridge|inset|outset)$/.test(part)) style = part;
      else if (/^(thin|medium|thick|[-+.\d])/.test(part)) width = part;
      else color = part;
    }
    const sides = property === "border" ? ["top", "right", "bottom", "left"] : [property.slice(7)];
    return sides.flatMap((side) => [["border-" + side + "-width", width], ["border-" + side + "-style", style], ["border-" + side + "-color", color]] as [string, string][]);
  }
  if (property === "text-decoration") {
    const lines: string[] = []; let style = "solid", color = "currentColor";
    for (const part of parts) {
      if (["none", "underline", "overline", "line-through"].includes(part)) lines.push(part);
      else if (["solid", "double", "dotted", "dashed", "wavy"].includes(part)) style = part;
      else color = part;
    }
    return [["text-decoration-line", lines.join(" ") || "none"], ["text-decoration-style", style], ["text-decoration-color", color]];
  }
  return [[property, value]];
}
function higher(next: Entry, previous: Entry | undefined): boolean {
  if (!previous) return true;
  if (next.important !== previous.important) return next.important;
  for (const index of [0, 1, 2] as const) if (next.specificity[index] !== previous.specificity[index]) return next.specificity[index] > previous.specificity[index];
  return next.order >= previous.order;
}
export class ExportCss {
  private readonly indexed = new Map<string, Rule[]>();
  private readonly cache = new WeakMap<SourceElement, Css>();
  private work = 0;
  constructor(styles: readonly string[], private readonly limits: DocumentReadLimits, private readonly revisionIds: ReadonlySet<string> = new Set()) {
    let order = 0;
    for (const source of ["p,h1,h2,h3,h4,h5,h6{margin:0;font-weight:inherit;line-height:normal}td,th{padding:0;font-weight:inherit;text-align:inherit;vertical-align:top}", ...styles]) {
      const ast = parse(source, { context: "stylesheet", parseCustomProperty: true });
      if (ast.type !== "StyleSheet") continue;
      for (const item of ast.children) {
        if (item.type !== "Rule" || item.prelude?.type !== "SelectorList") throw new DocumentError("UnsupportedFeature", "Only validated source styles can be compiled to DOCX.");
        const declarations = item.block.children.toArray().filter((node): node is Declaration => node.type === "Declaration");
        for (const choice of item.prelude.children) {
          if (choice.type !== "Selector") continue;
          const compiled = selector(choice), description = generate(choice);
          const revisionDisplay = [".word-main ins", ".word-main del"].includes(description) && declarations.every((decl) => {
            const value = generate(decl.value).toLowerCase();
            return decl.property === "color" && ["#006400", "#a00000"].includes(value) || decl.property === "text-decoration" && ["underline", "line-through"].includes(value);
          });
          const rule = { ...compiled, declarations, order: order++, revisionDisplay };
          const list = this.indexed.get(rule.key) ?? []; list.push(rule); this.indexed.set(rule.key, list);
        }
      }
    }
  }
  private step(): void {
    if (++this.work > this.limits.maxCssNodes + this.limits.maxNodes * 128) throw new DocumentError("ResourceLimitExceeded", "DOCX style resolution exceeds its work budget.");
  }
  private matches(node: SourceElement, rule: Rule): boolean {
    const compound = (target: SourceElement, part: Compound): boolean => {
      this.step(); const classes = new Set((target.attrs.class ?? "").split(/\s+/));
      return (!part.tag || part.tag === target.tag) && (!part.root || target.tag === "html") && part.classes.every((name) => classes.has(name))
        && part.attrs.every((attr) => target.attrs[attr.name] !== undefined && (attr.value === undefined || target.attrs[attr.name] === attr.value));
    };
    let current: SourceElement | undefined = node;
    for (let index = rule.compounds.length - 1; index >= 0; index--) {
      if (index === rule.compounds.length - 1) { if (!compound(current!, rule.compounds[index]!)) return false; }
      else { while (current && !compound(current, rule.compounds[index]!)) current = current.parent; if (!current) return false; }
      current = current?.parent;
    }
    return true;
  }
  get(node: SourceElement): Css {
    const known = this.cache.get(node); if (known) return known;
    const parent = node.parent ? this.get(node.parent) : initial;
    const selected = new Set<Rule>();
    for (const key of ["*", node.tag, ...(node.tag === "html" ? [":root"] : []), ...(node.attrs.class ?? "").split(/\s+/).filter(Boolean).map((name) => "." + name), ...Object.keys(node.attrs).map((name) => "[" + name)])
      for (const rule of this.indexed.get(key) ?? []) if (this.matches(node, rule) && !(rule.revisionDisplay && this.revisionIds.has(node.attrs["data-node-id"] ?? ""))) selected.add(rule);
    const entries: { declarations: Declaration[]; specificity: Specificity; order: number }[] = [...selected];
    if (node.attrs.style) {
      const ast = parse(node.attrs.style, { context: "declarationList", parseCustomProperty: true });
      if (ast.type === "DeclarationList") entries.push({ declarations: ast.children.toArray().filter((item): item is Declaration => item.type === "Declaration"), specificity: [1, 0, 0], order: Number.MAX_SAFE_INTEGER });
    }
    const result = this.compute(parent, entries, node);
    this.cache.set(node, result); return result;
  }
  classStyle(name: string, parent: Css = initial): Css {
    const rules = (this.indexed.get("." + name) ?? []).filter((rule) => rule.compounds.length === 1 && rule.compounds[0]!.classes.length > 0 && rule.compounds[0]!.classes.every((value) => value === name)
      && !rule.compounds[0]!.tag && !rule.compounds[0]!.root && rule.compounds[0]!.attrs.length === 0);
    return this.compute(parent, rules);
  }
  classKeys(name: string): readonly string[] {
    const keys = new Set<string>();
    for (const rule of this.indexed.get("." + name) ?? []) if (rule.compounds.length === 1 && rule.compounds[0]!.classes.every((value) => value === name))
      for (const declaration of rule.declarations) for (const [key] of expand(declaration.property, generate(declaration.value))) keys.add(key);
    return [...keys];
  }
  private compute(parent: Css, rules: readonly { declarations: Declaration[]; specificity: Specificity; order: number }[], node?: SourceElement): Css {
    const result: Record<string, string> = { ...initial }, custom = new Map<string, Entry>(), winners = new Map<string, Entry>();
    for (const [name, value] of Object.entries(parent)) if (name.startsWith("--") || inherited.has(name)) result[name] = value;
    if (node?.parent && inline.has(node.parent.tag)) for (const name of ["background-color", "text-decoration-line", "text-decoration-style", "text-decoration-color"]) result[name] = parent[name] ?? initial[name]!;
    if (node) {
      if (["b", "strong"].includes(node.tag)) result["font-weight"] = "bold";
      if (["em", "i"].includes(node.tag)) result["font-style"] = "italic";
      if (node.tag === "u" || node.tag === "ins" && !this.revisionIds.has(node.attrs["data-node-id"] ?? "")) result["text-decoration-line"] = "underline";
      if (node.tag === "s" || node.tag === "del" && !this.revisionIds.has(node.attrs["data-node-id"] ?? "")) result["text-decoration-line"] = "line-through";
      if (/^h[1-6]$/.test(node.tag)) result["font-size"] = [2, 1.5, 1.17, 1, .83, .67][Number(node.tag[1]) - 1] + "em";
      if (["sub", "sup"].includes(node.tag)) result["vertical-align"] = node.tag === "sup" ? "super" : "sub";
      if (["td", "th"].includes(node.tag)) result["vertical-align"] = "top";
    }
    let serial = 0;
    for (const rule of [...rules].sort((a, b) => a.order - b.order)) for (const declaration of rule.declarations) {
      this.step(); const decoded = ident.decode(declaration.property), name = decoded.startsWith("--") ? decoded : decoded.toLowerCase(), entry = { value: generate(declaration.value), important: Boolean(declaration.important), specificity: rule.specificity, order: serial++ };
      if (name.startsWith("--") && higher(entry, custom.get(name))) custom.set(name, entry);
    }
    const cycles = new Set<string>(), visited = new Set<string>(), path: string[] = [];
    const dependencies = (name: string): void => {
      const start = path.indexOf(name); if (start >= 0) { for (const member of path.slice(start)) cycles.add(member); return; }
      if (visited.has(name) || !custom.has(name)) return;
      if (path.length >= this.limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "CSS variable graph exceeds its depth budget.");
      path.push(name);
      walk(parse(custom.get(name)!.value, { context: "value", parseCustomProperty: true }), (item) => {
        this.step(); if (item.type === "Function" && ident.decode(item.name).toLowerCase() === "var" && item.children.first?.type === "Identifier") dependencies(ident.decode(item.children.first.name));
      });
      path.pop(); visited.add(name);
    };
    for (const name of custom.keys()) dependencies(name);
    const memo = new Map<string, string | undefined>(), resolving = new Set<string>();
    const variable = (name: string): string | undefined => {
      if (memo.has(name)) return memo.get(name);
      if (cycles.has(name)) return undefined;
      if (resolving.has(name)) return undefined;
      if (resolving.size >= this.limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "CSS variable resolution exceeds its depth budget.");
      const local = custom.get(name); if (!local) return parent[name];
      resolving.add(name); const value = ["inherit", "unset"].includes(local.value) ? parent[name] : local.value === "initial" ? undefined : variables(local.value, variable, 0, this.limits.maxFileBytes);
      resolving.delete(name); memo.set(name, value); return value;
    };
    for (const name of custom.keys()) { const value = variable(name); if (value === undefined) delete result[name]; else result[name] = value; }
    for (const rule of [...rules].sort((a, b) => a.order - b.order)) for (const declaration of rule.declarations) {
      const decoded = ident.decode(declaration.property), property = decoded.startsWith("--") ? decoded : decoded.toLowerCase(); if (property.startsWith("--")) continue;
      const value = variables(generate(declaration.value), (name) => result[name], 0, this.limits.maxFileBytes);
      // Invalid-at-computed-value declarations still win cascade and compute to unset.
      for (const [name, expanded] of expand(property, value ?? "unset")) {
        const entry = { value: expanded, important: Boolean(declaration.important), specificity: rule.specificity, order: serial++ };
        if (higher(entry, winners.get(name))) winners.set(name, entry);
      }
    }
    for (const [name, entry] of winners) {
      const value = entry.value;
      if (value === "inherit" || value === "unset" && inherited.has(name)) result[name] = parent[name] ?? initial[name] ?? "initial";
      else if (value === "initial" || value === "unset" || value === "revert") result[name] = initial[name] ?? "initial";
      else result[name] = lexer.matchProperty(name, value).error ? (inherited.has(name) ? parent[name] ?? initial[name] ?? "initial" : initial[name] ?? "initial") : value;
    }
    const base = parseFloat(parent["font-size"] ?? "12pt"), size = result["font-size"] ?? "12pt";
    const sizes: Record<string, number> = { "xx-small": 7.2, "x-small": 9, small: 10, medium: 12, large: 13.5, "x-large": 18, "xx-large": 24, "xxx-large": 36, smaller: base / 1.2, larger: base * 1.2 };
    result["font-size"] = String(sizes[size] ?? length(size, { font: base, percent: base })) + "pt";
    if (["bolder", "lighter"].includes(result["font-weight"] ?? "")) {
      const weight = parent["font-weight"] === "bold" ? 700 : Number(parent["font-weight"]) || 400;
      result["font-weight"] = String(result["font-weight"] === "bolder" ? weight < 350 ? 400 : weight < 550 ? 700 : 900 : weight < 550 ? 100 : weight < 750 ? 400 : 700);
    }
    return Object.freeze(result);
  }
}
