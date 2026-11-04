import { checkLimit, freeze, invalid, unsupported } from "../document/validation.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "./chunk-index.js";

/** Private first-party parser foundation. Structured references and whole axes
 * are added at the R3-05 gate; unrecognized syntax is rejected, never stripped. */
export type FormulaExpression =
  | { readonly kind: "number"; readonly value: number }
  | { readonly kind: "string"; readonly value: string }
  | { readonly kind: "boolean"; readonly value: boolean }
  | { readonly kind: "error"; readonly value: string }
  | { readonly kind: "name"; readonly name: string; readonly sheet?: string }
  | { readonly kind: "cell"; readonly row: number; readonly column: number; readonly absoluteRow: boolean; readonly absoluteColumn: boolean; readonly sheet?: string }
  | { readonly kind: "range"; readonly from: FormulaCell; readonly to: FormulaCell }
  | { readonly kind: "unary"; readonly operator: "+" | "-" | "%"; readonly operand: FormulaExpression }
  | { readonly kind: "binary"; readonly operator: string; readonly left: FormulaExpression; readonly right: FormulaExpression }
  | { readonly kind: "call"; readonly name: string; readonly arguments: readonly (FormulaExpression | null)[] };
export type FormulaCell = Extract<FormulaExpression, { kind: "cell" }>;
export interface FormulaOrigin { readonly row: number; readonly column: number }
interface Token { readonly kind: "number" | "string" | "quoted" | "word" | "error" | "operator" | "end"; readonly text: string }

const operators: Readonly<Record<string, number>> = { "=": 10, "<>": 10, "<": 10, ">": 10, "<=": 10, ">=": 10, "&": 20, "+": 30, "-": 30, "*": 40, "/": 40, "^": 50, ":": 80 };
const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/;
const word = /^\$?[\p{L}\p{Nl}_\\][\p{L}\p{N}\p{M}_.?\\$]*/u;
const error = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA)/i;

function tokenize(source: string): Token[] {
  const tokens: Token[] = [];
  for (let cursor = 1; cursor < source.length;) {
    const character = source[cursor]!;
    if (character === " " || character === "\r" || character === "\n") { cursor++; continue; }
    if (character === '"' || character === "'") {
      const quote = character; let value = "", closed = false; cursor++;
      while (cursor < source.length) {
        const next = source[cursor++]!;
        if (next === quote) {
          if (source[cursor] === quote) { value += quote; cursor++; } else { closed = true; break; }
        } else value += next;
      }
      if (!closed) invalid("Unclosed formula string or sheet name.");
      tokens.push({ kind: quote === '"' ? "string" : "quoted", text: value });
    } else {
      const rest = source.slice(cursor), numeric = number.exec(rest), identifier = word.exec(rest), failure = error.exec(rest);
      if (numeric) { tokens.push({ kind: "number", text: numeric[0] }); cursor += numeric[0].length; }
      else if (identifier) { tokens.push({ kind: "word", text: identifier[0] }); cursor += identifier[0].length; }
      else if (failure) { tokens.push({ kind: "error", text: failure[0].toUpperCase() }); cursor += failure[0].length; }
      else {
        const pair = source.slice(cursor, cursor + 2);
        if (["<=", ">=", "<>"].includes(pair)) { tokens.push({ kind: "operator", text: pair }); cursor += 2; }
        else if ("+-*/^&=<>%:(),!".includes(character)) { tokens.push({ kind: "operator", text: character }); cursor++; }
        else unsupported("Formula syntax is outside the implemented parser grammar.");
      }
    }
    checkLimit(tokens.length, 4096);
  }
  tokens.push({ kind: "end", text: "" }); return tokens;
}

function reference(text: string, origin: FormulaOrigin, sheet?: string): FormulaExpression {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(text);
  const scope = sheet === undefined ? {} : { sheet };
  if (!match) {
    if (text.includes("$")) invalid("Invalid formula reference.");
    return { kind: "name", name: text, ...scope };
  }
  let column = 0;
  for (const letter of match[2]!.toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
  const row = Number(match[4]);
  if (row > MAX_SHEET_ROWS || column > MAX_SHEET_COLUMNS) {
    // An out-of-grid identifier may be a name, but a $ reference cannot be one.
    if (match[1] || match[3]) invalid("Formula reference exceeds worksheet bounds.");
    return { kind: "name", name: text, ...scope };
  }
  return { kind: "cell", row: row - 1 - (match[3] ? 0 : origin.row), column: column - 1 - (match[1] ? 0 : origin.column),
    absoluteRow: Boolean(match[3]), absoluteColumn: Boolean(match[1]), ...scope };
}

/** Relative references are offsets from origin, making copied formulas share an AST. */
export function parseFormula(source: string, origin: FormulaOrigin): FormulaExpression {
  checkLimit(source.length, 8192);
  if (!source.startsWith("=") || source.length < 2) invalid("Formula source must begin with equals.");
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(source)) invalid("Invalid formula control character.");
  if (!Number.isInteger(origin.row) || !Number.isInteger(origin.column) || origin.row < 0 || origin.column < 0
    || origin.row >= MAX_SHEET_ROWS || origin.column >= MAX_SHEET_COLUMNS) invalid("Invalid formula origin.");
  const tokens = tokenize(source); let index = 0, nodes = 0;
  const peek = (): Token => tokens[index]!;
  const take = (): Token => tokens[index++]!;
  const consume = (text: string): void => { if (take().text !== text) invalid("Unexpected formula token."); };
  function expression(minimum: number, depth: number): FormulaExpression {
    checkLimit(depth, 128); checkLimit(++nodes, 4096);
    const token = take(); let left: FormulaExpression;
    if (token.text === "+" || token.text === "-") left = { kind: "unary", operator: token.text, operand: expression(70, depth + 1) };
    else if (token.text === "(") { left = expression(0, depth + 1); consume(")"); }
    else if (token.kind === "number") {
      const value = Number(token.text); if (!Number.isFinite(value)) invalid("Invalid formula numeric constant.");
      left = { kind: "number", value };
    } else if (token.kind === "string") left = { kind: "string", value: token.text };
    else if (token.kind === "error") left = { kind: "error", value: token.text };
    else if (token.kind === "word" || token.kind === "quoted") {
      if (peek().text === "!") {
        take(); const target = take();
        if (!token.text || /[\[\]\\/:?*\u0000-\u001f]/.test(token.text)) unsupported("Unsupported formula sheet reference.");
        if (target.kind === "error" && target.text === "#REF!") left = { kind: "error", value: "#REF!" };
        else { if (target.kind !== "word") invalid("Expected a formula reference after a sheet name."); left = reference(target.text, origin, token.text); }
      } else if (token.kind === "quoted") invalid("Quoted formula sheet names require a reference.");
      else if (peek().text === "(") {
        if (token.text.includes("$") || token.text.includes("\\")) invalid("Invalid formula function name.");
        take(); const args: (FormulaExpression | null)[] = [];
        if (peek().text !== ")") {
          while (true) {
            args.push(peek().text === "," || peek().text === ")" ? null : expression(0, depth + 1));
            checkLimit(args.length, 254);
            if (peek().text !== ",") break;
            take();
          }
        }
        consume(")"); left = { kind: "call", name: token.text.toUpperCase(), arguments: args };
      } else if (/^(TRUE|FALSE)$/i.test(token.text)) left = { kind: "boolean", value: token.text.toUpperCase() === "TRUE" };
      else left = reference(token.text, origin);
    } else invalid("Expected a formula operand.");

    while (true) {
      if (peek().text === "%" && 60 >= minimum) { take(); left = { kind: "unary", operator: "%", operand: left }; continue; }
      const operator = peek().text, precedence = operators[operator];
      if (precedence === undefined || precedence < minimum) break;
      take(); const right = expression(precedence + 1, depth + 1);
      if (operator === ":") {
        if (left.kind !== "cell" || right.kind !== "cell") unsupported("This formula range form is not yet implemented.");
        if (right.sheet !== undefined && right.sheet !== left.sheet) unsupported("Three-dimensional formula ranges are not supported.");
        left = { kind: "range", from: left, to: right.sheet === undefined && left.sheet !== undefined ? { ...right, sheet: left.sheet } : right };
      } else left = { kind: "binary", operator, left, right };
    }
    return left;
  }
  const ast = expression(0, 0);
  if (peek().kind !== "end") unsupported("Unconsumed formula syntax is not supported.");
  // A long left-associative chain is shallow for the parser but deep as an AST.
  // Check it before recursive freezing or serialization can use the JS stack.
  const pending = [{ node: ast, depth: 0 }];
  while (pending.length) {
    const { node, depth } = pending.pop()!; checkLimit(depth, 128);
    const children = node.kind === "binary" ? [node.left, node.right] : node.kind === "unary" ? [node.operand]
      : node.kind === "call" ? node.arguments.filter((arg): arg is FormulaExpression => arg !== null) : node.kind === "range" ? [node.from, node.to] : [];
    for (const child of children) pending.push({ node: child, depth: depth + 1 });
  }
  return freeze(ast);
}

export interface FormulaTemplatePool {
  readonly size: number;
  readonly keyCodeUnits: number;
  intern(source: string, origin: FormulaOrigin): number;
  get(id: number): FormulaExpression | undefined;
}
/** Per-workbook derived cache, bounded independently from source bytes. */
export function createFormulaTemplatePool(maxTemplates = 250_000, maxKeyCodeUnits = 32 * 1024 * 1024): FormulaTemplatePool {
  if (![maxTemplates, maxKeyCodeUnits].every(value => Number.isSafeInteger(value) && value > 0)) invalid("Invalid formula template budget.");
  const ids = new Map<string, number>(), expressions: FormulaExpression[] = []; let keyCodeUnits = 0;
  return Object.freeze({
    get size() { return expressions.length; }, get keyCodeUnits() { return keyCodeUnits; },
    intern(source: string, origin: FormulaOrigin): number {
      const ast = parseFormula(source, origin), key = JSON.stringify(ast), existing = ids.get(key);
      if (existing !== undefined) return existing;
      checkLimit(expressions.length + 1, maxTemplates); checkLimit(keyCodeUnits + key.length, maxKeyCodeUnits);
      const id = expressions.length; ids.set(key, id); expressions.push(ast); keyCodeUnits += key.length; return id;
    },
    get(id: number) { return expressions[id]; },
  });
}
