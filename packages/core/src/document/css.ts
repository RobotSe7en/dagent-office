import { generate, ident, lexer, parse, walk } from "css-tree";
import type { CssNode, Declaration } from "css-tree";
import type { DocumentReadLimits } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { checkLimit, invalid, unsupported } from "./validation.js";

const properties = new Set([
  "color", "background-color", "font-family", "font-size", "font-weight", "font-style",
  "text-decoration", "text-decoration-line", "text-decoration-color", "text-decoration-style",
  "text-align", "text-indent", "line-height", "letter-spacing", "word-spacing", "white-space",
  "vertical-align", "width", "height", "min-width", "max-width", "min-height", "max-height",
  "margin", "margin-top", "margin-right", "margin-bottom", "margin-left",
  "padding", "padding-top", "padding-right", "padding-bottom", "padding-left",
  "border", "border-width", "border-style", "border-color", "border-collapse", "border-spacing",
  ...["top", "right", "bottom", "left"].flatMap((side) => [
    "border-" + side, "border-" + side + "-width", "border-" + side + "-style", "border-" + side + "-color",
  ]),
]);
const functions = new Set(["var", "calc", "min", "max", "clamp", "rgb", "rgba", "hsl", "hsla"]);

export interface CssBudget { nodes: number }

function validateDeclaration(node: Declaration): void {
  const property = ident.decode(node.property);
  if (property === "--") invalid("A CSS custom property must have a name.");
  if (!property.startsWith("--") && !properties.has(property.toLowerCase())) {
    unsupported("The CSS property is outside the supported document style set.");
  }
  let hasVariable = false;
  walk(node.value, (child) => {
    if (child.type === "Url" || child.type === "Raw") {
      invalid("CSS resources and unparsed values are not supported.");
    }
    if (child.type === "Function") {
      const name = ident.decode(child.name).toLowerCase();
      if (!functions.has(name)) unsupported("The CSS function is not supported.");
      if (name === "var") {
        const first = child.children.first;
        if (first?.type !== "Identifier" || !ident.decode(first.name).startsWith("--")
          || ident.decode(first.name).length <= 2) {
          invalid("CSS var() must reference a custom property.");
        }
        hasVariable = true;
      }
    }
  });
  if (!property.startsWith("--") && !hasVariable && lexer.matchProperty(property, node.value).error) {
    invalid("A CSS declaration value does not match its property.");
  }
}

/** Parse and validate without resolving the cascade or materializing computed styles. */
export function normalizeCss(source: string, context: "stylesheet" | "declarationList", limits: DocumentReadLimits, budget: CssBudget): string {
  try {
    const ast = parse(source, {
      context,
      parseCustomProperty: true,
      onParseError: () => invalid("The CSS source is malformed."),
    });
    let depth = 0;
    walk(ast, {
      enter(node: CssNode) {
        checkLimit(++budget.nodes, limits.maxCssNodes);
        checkLimit(++depth, limits.maxDepth);
        switch (node.type) {
          case "Atrule":
          case "IdSelector":
          case "PseudoElementSelector":
            unsupported("The CSS rule or selector is outside the supported style set.");
            break;
          case "Raw":
          case "Url":
            invalid("CSS contains unparsed content or a resource reference.");
            break;
          case "TypeSelector":
            if (!/^[a-z][a-z0-9-]*$/i.test(ident.decode(node.name))) unsupported("The CSS type selector is not supported.");
            break;
          case "Combinator":
            if (node.name !== " ") unsupported("Only descendant CSS combinators are supported.");
            break;
          case "PseudoClassSelector":
            if (ident.decode(node.name).toLowerCase() !== "root" || node.children !== null) {
              unsupported("Only the :root CSS pseudo-class is supported.");
            }
            break;
          case "AttributeSelector":
            if (!/^data-[a-z0-9-]+$/.test(ident.decode(node.name.name))
              || (node.matcher !== null && node.matcher !== "=") || node.flags !== null) {
              unsupported("Only document data-attribute presence and equality selectors are supported.");
            }
            break;
          case "Declaration":
            validateDeclaration(node);
            break;
          case "Rule":
            if (node.prelude?.type !== "SelectorList") unsupported("Only ordinary CSS style rules are supported.");
            if (node.block.children.some((child) => child.type !== "Declaration")) {
              unsupported("Nested CSS rules are not supported.");
            }
            break;
          case "StyleSheet":
            if (node.children.some((child) => child.type !== "Rule")) unsupported("Only CSS style rules are supported.");
            break;
          case "DeclarationList":
            if (node.children.some((child) => child.type !== "Declaration")) invalid("An inline style must contain only declarations.");
            break;
        }
      },
      leave() { depth--; },
    });
    return generate(ast);
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    return invalid("The CSS source could not be parsed or validated.");
  }
}
