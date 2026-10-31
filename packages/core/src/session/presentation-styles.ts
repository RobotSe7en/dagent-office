import { clone, generate, ident, parse, walk } from "css-tree";
import type { DocumentReadLimits } from "../runtime.js";
import { normalizeCss, setInlineStyle } from "../document/css.js";
import { encode, readText } from "../document/package.js";
import { invalid } from "../document/validation.js";
import type { PresentationEditContext } from "./presentation-edit.js";
export function allocatePresentationClass(ctx: PresentationEditContext, prefix: string): string {
    const classes = new Set<string>(), html = ctx.parse();
    for (const id of html.nodeIds())
        for (const name of (html.peekNode(id)!.attributes.class ?? "").split(/\s+/))
            classes.add(name);
    for (const sheet of ctx.styles.snapshot())
        walk(parse(sheet.css, { context: "stylesheet" }), node => { if (node.type === "ClassSelector")
            classes.add(ident.decode(node.name)); });
    let value: string;
    do {
        value = prefix + ctx.allocate();
    } while (classes.has(value));
    return value;
}
/** Candidate-local CSS writes use the package's existing, bounded stylesheets. */
export class PresentationStyleEdit {
    constructor(private readonly files: Map<string, Uint8Array>, private readonly paths: readonly string[], private readonly limits: DocumentReadLimits) { }
    snapshot(): {
        path: string;
        css: string;
    }[] { return this.paths.map(path => ({ path, css: normalizeCss(readText(this.files, path), "stylesheet", this.limits, { nodes: 0 }) })); }
    private append(css: string): void { const path = this.paths.at(-1); if (!path)
        invalid("Presentation style changes require a declared stylesheet."); this.files.set(path, encode(normalizeCss(readText(this.files, path) + css, "stylesheet", this.limits, { nodes: 0 }))); }
    setClass(name: string, properties: Readonly<Record<string, string>>): void {
        const current = new Map<string, string>();
        for (const sheet of this.snapshot()) {
            const ast = parse(sheet.css, { context: "stylesheet" });
            if (ast.type !== "StyleSheet")
                continue;
            ast.children.forEach(rule => { if (rule.type !== "Rule" || generate(rule.prelude) !== "." + ident.encode(name))
                return; rule.block.children.forEach(node => { if (node.type === "Declaration")
                current.set(node.property, generate(node.value) + (node.important ? "!important" : "")); }); });
        }
        let declarations = "";
        for (const [property, value] of Object.entries(properties)) {
            const incoming = setInlineStyle("", property, value, this.limits), existing = current.get(property);
            if (existing !== undefined && incoming === setInlineStyle("", property, existing, this.limits))
                continue;
            declarations = setInlineStyle(declarations, property, value, this.limits);
        }
        if (declarations)
            this.append("." + ident.encode(name) + "{" + declarations + "}");
    }
    cloneClass(source: string, target: string): void {
        let css = "";
        for (const sheet of this.snapshot()) {
            const ast = parse(sheet.css, { context: "stylesheet" });
            if (ast.type !== "StyleSheet")
                invalid("Invalid presentation stylesheet.");
            ast.children.forEach(rule => {
                if (rule.type !== "Rule")
                    return;
                let found = false;
                walk(rule.prelude, node => { if (node.type === "ClassSelector" && ident.decode(node.name) === source)
                    found = true; });
                if (!found)
                    return;
                const next = clone(rule);
                if (next.type !== "Rule" || next.prelude.type !== "SelectorList")
                    return;
                next.prelude.children.forEach((selector, item, list) => { let contains = false, ambiguous = false; walk(selector, node => { if (node.type === "ClassSelector" && ident.decode(node.name) === source)
                    contains = true; if (node.type === "PseudoClassSelector")
                    ambiguous = true; }); if (!contains)
                    list.remove(item);
                else if (ambiguous)
                    invalid("Cloning a style class in pseudo-class selectors requires an explicit source mapping."); });
                walk(next.prelude, node => { if (node.type === "ClassSelector" && ident.decode(node.name) === source)
                    node.name = ident.encode(target); });
                css += generate(next);
            });
        }
        this.append(css || "." + ident.encode(target) + "{}");
    }
}
