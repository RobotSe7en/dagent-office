import { generate, ident, parse, walk } from "css-tree";
import type { ContextStyle } from "../contracts/interaction.js";
import type { DeepReadonly } from "../runtime.js";
interface SelectorUse {
    classes: string[];
    ids: string[];
    uncertain: boolean;
}
/** Conservative selector dependencies; ambiguous selectors remain global. */
export function presentationStyleImpact(styles: readonly DeepReadonly<ContextStyle>[]): (classes: ReadonlySet<string>, ids: ReadonlySet<string>) => string {
    const rules: {
        css: string;
        selectors: SelectorUse[];
    }[] = [];
    for (const sheet of styles) {
        const ast = parse(sheet.css, { context: "stylesheet" });
        if (ast.type !== "StyleSheet")
            continue;
        ast.children.forEach(rule => {
            if (rule.type !== "Rule")
                return;
            const selectors: SelectorUse[] = [];
            if (rule.prelude.type === "SelectorList")
                rule.prelude.children.forEach(selector => { const use: SelectorUse = { classes: [], ids: [], uncertain: false }; walk(selector, node => { if (node.type === "ClassSelector")
                    use.classes.push(ident.decode(node.name));
                else if (node.type === "IdSelector")
                    use.ids.push(ident.decode(node.name));
                else if (node.type === "PseudoClassSelector" || node.type === "AttributeSelector")
                    use.uncertain = true; }); selectors.push(use); });
            rules.push({ css: generate(rule), selectors });
        });
    }
    return (classes, ids) => rules.filter(rule => !rule.selectors.length || rule.selectors.some(use => use.uncertain || use.classes.every(name => classes.has(name)) && use.ids.every(id => ids.has(id)))).map(rule => rule.css).join("");
}
