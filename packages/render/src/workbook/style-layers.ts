import { DocumentError } from "@doffi/core";

export type WorkbookStyleProperties = ReadonlyMap<string, string>;
const standard = new Set(["color", "background-color", "font-family", "font-size", "font-weight", "font-style", "font-variant", "line-height",
  "letter-spacing", "word-spacing", "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-transform", "vertical-align",
  "text-align", "text-indent", "white-space", "padding-left", "padding-right", "padding-top", "padding-bottom",
  ...["left", "right", "top", "bottom"].flatMap(side => [`border-${side}-style`, `border-${side}-width`, `border-${side}-color`])]);
export function workbookAppearanceProperty(name: string): boolean {
  return standard.has(name) || name.startsWith("--office-cell-") || name.startsWith("--office-font-");
}
function selectors(source: string): string[] {
  const values: string[] = []; let from = 0, quote = "", bracket = 0;
  for (let i = 0; i < source.length; i++) {
    const character = source[i]!;
    if (character === "\\") { i++; continue; }
    if (quote) { if (character === quote) quote = ""; continue; }
    if (character === '"' || character === "'") quote = character;
    else if (character === "[") bracket++;
    else if (character === "]") bracket--;
    else if (character === "," && !bracket) { values.push(source.slice(from, i)); from = i + 1; }
  }
  values.push(source.slice(from)); return values;
}
function longhands(name: string): readonly string[] {
  if (name === "text-decoration") return ["text-decoration-line", "text-decoration-style", "text-decoration-color"];
  if (name === "padding") return ["padding-left", "padding-right", "padding-top", "padding-bottom"];
  if (name === "border" || name === "border-width" || name === "border-style" || name === "border-color") {
    const kinds = name === "border" ? ["width", "style", "color"] : [name.slice(7)];
    return ["left", "right", "top", "bottom"].flatMap(side => kinds.map(kind => `border-${side}-${kind}`));
  }
  if (/^border-(?:left|right|top|bottom)$/.test(name)) return [name + "-width", name + "-style", name + "-color"];
  return [name];
}

/** The browser resolves specificity, !important, variables and inheritance.
 * Only declarations belonging to a differential's class become overrides;
 * inherited/default values never accidentally turn a DXF into a full XF.
 * These probes live beside the source cell and are always released. */
export class WorkbookStyleLayers {
  private readonly rules: readonly { readonly selector: string; readonly names: readonly string[] }[];
  private readonly names: ReadonlySet<string>;
  constructor(private readonly document: Document, private readonly account: (amount: number) => void) {
    const rules: { selector: string; names: readonly string[] }[] = [];
    for (const sheet of document.styleSheets) for (const raw of sheet.cssRules) {
      if (raw.type !== 1) throw new DocumentError("InvalidDocument", "Workbook appearance requires ordinary source style rules.");
      const rule = raw as CSSStyleRule, names = [...new Set([...rule.style].flatMap(longhands).filter(workbookAppearanceProperty))];
      if (names.length) for (const selector of selectors(rule.selectorText)) rules.push({ selector, names });
    }
    this.rules = rules;
    this.names = new Set([...standard, ...rules.flatMap(rule => rule.names)]);
  }
  read(element: HTMLTableCellElement, cssClass: string, differential = false): WorkbookStyleProperties {
    this.account(1); const probe = element.cloneNode(false) as HTMLTableCellElement;
    probe.removeAttribute("style"); probe.className = ""; element.parentElement!.append(probe);
    try {
      const without = differential ? this.rules.map(rule => probe.matches(rule.selector)) : [];
      probe.className = cssClass;
      const computed = this.document.defaultView!.getComputedStyle(probe);
      const names = differential ? new Set(this.rules.flatMap((rule, index) => !without[index] && probe.matches(rule.selector) ? rule.names : []))
        // Old Chromium resolves custom properties by name but does not list
        // them in the computed declaration's iterator. Source declarations
        // provide the same complete property vocabulary in every engine.
        : new Set([...this.names, ...computed].filter(workbookAppearanceProperty));
      const result = new Map([...names].map(name => [name, computed.getPropertyValue(name).trim()]));
      // An ordinary CSS paint is a complete paint at this layer. Flattened
      // native defaults on the target cell must not mask it later.
      if (differential && result.has("color") && !result.has("--office-cell-foreground-color")) {
        result.set("--office-cell-foreground-color", result.get("color")!); result.set("--office-cell-foreground-tint", "0");
      }
      if (differential && result.has("background-color") && !result.has("--office-cell-fill-kind")) result.set("--office-cell-fill-kind", "initial");
      return result;
    } finally { probe.remove(); this.account(-1); }
  }
}

/** Properties that jointly express one native paint must travel together. A
 * changed fill kind or border kind cannot inherit a lower layer's color/tint. */
export function workbookStylePropertyGroup(name: string): string {
  if (name.startsWith("--office-cell-fill-") || name.startsWith("--office-cell-gradient-")) return "fill";
  const border = /^--office-cell-border-(left|right|start|end|top|bottom|diagonal|vertical|horizontal)(?:-|$)/.exec(name);
  if (border) return "border-" + (border[1] === "start" ? "left" : border[1] === "end" ? "right" : border[1]);
  if (name === "color" || name.startsWith("--office-cell-foreground-")) return "foreground";
  if (name === "font-family" || /^--office-font-(?:source-name|ascii|h-ansi|east-asia|complex-script|scheme)$/.test(name)) return "typeface";
  return name;
}

export function workbookStyleDifferences(source: WorkbookStyleProperties, normal: WorkbookStyleProperties): Map<string, string> {
  const changed = new Set([...source].filter(([name, value]) => !name.endsWith("-native-alpha") && value !== (normal.get(name) ?? "")).map(([name]) => workbookStylePropertyGroup(name)));
  return new Map([...source].filter(([name]) => changed.has(workbookStylePropertyGroup(name))));
}
