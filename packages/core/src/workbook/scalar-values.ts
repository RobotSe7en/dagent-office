import type { DeepReadonly } from "../runtime.js";
import type { WorkbookValue } from "../contracts/workbook.js";

export type ScalarValue = DeepReadonly<WorkbookValue>;
export type ScalarError = Extract<ScalarValue, { type: "error" }>;
export const scalarBlank: ScalarValue = Object.freeze({ type: "blank" });
export function scalarError(value: ScalarError["value"]): ScalarError { return { type: "error", value }; }
export function scalarNumber(value: number): ScalarValue { return Number.isFinite(value) ? { type: "number", value: value === 0 ? 0 : value } : scalarError("#NUM!"); }
export function scalarBoolean(value: boolean): ScalarValue { return { type: "boolean", value }; }
export function scalarText(value: string): ScalarValue { return value.length <= 32_767 ? { type: "text", value } : scalarError("#VALUE!"); }
export function scalarToNumber(value: ScalarValue): number | ScalarError {
  if (value.type === "error") return value;
  if (value.type === "blank") return 0;
  if (value.type === "number") return value.value;
  if (value.type === "boolean") return Number(value.value);
  const source = value.value.trim(), match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)(%)?$/.exec(source);
  if (!match) return scalarError("#VALUE!");
  const number = Number(match[1]) / (match[2] ? 100 : 1);
  return Number.isFinite(number) ? number : scalarError("#NUM!");
}
export function scalarToBoolean(value: ScalarValue): boolean | ScalarError {
  if (value.type === "error") return value;
  if (value.type === "boolean") return value.value;
  if (value.type === "blank") return false;
  if (value.type === "number") return value.value !== 0;
  const source = value.value.toUpperCase();
  return source === "TRUE" ? true : source === "FALSE" ? false : scalarError("#VALUE!");
}
export function scalarToText(value: ScalarValue): string | ScalarError {
  if (value.type === "error") return value;
  return value.type === "blank" ? "" : value.type === "boolean" ? value.value ? "TRUE" : "FALSE" : value.type === "number" ? String(value.value).toUpperCase() : value.value;
}
export function scalarCompare(first: ScalarValue, second: ScalarValue): number | ScalarError {
  if (first.type === "error") return first;
  if (second.type === "error") return second;
  if (first.type === "blank" && second.type === "blank") return 0;
  const blankAs = (other: ScalarValue): ScalarValue => other.type === "text" ? { type: "text", value: "" } : other.type === "boolean" ? scalarBoolean(false) : scalarNumber(0);
  const a = first.type === "blank" ? blankAs(second) : first, b = second.type === "blank" ? blankAs(first) : second;
  const rank = { number: 0, text: 1, boolean: 2, blank: 0, error: 3 };
  if (a.type !== b.type) return rank[a.type] - rank[b.type];
  const left = a.type === "text" ? a.value.toUpperCase() : a.type === "blank" ? 0 : a.value;
  const right = b.type === "text" ? b.value.toUpperCase() : b.type === "blank" ? 0 : b.value;
  return left === right ? 0 : left < right ? -1 : 1;
}
export function scalarBinary(operator: string, a: ScalarValue, b: ScalarValue): ScalarValue {
  if (a.type === "error") return a;
  if (b.type === "error") return b;
  if (["=", "<>", "<", ">", "<=", ">="].includes(operator)) {
    const order = scalarCompare(a, b); if (typeof order !== "number") return order;
    return scalarBoolean(operator === "=" ? order === 0 : operator === "<>" ? order !== 0 : operator === "<" ? order < 0 : operator === ">" ? order > 0 : operator === "<=" ? order <= 0 : order >= 0);
  }
  if (operator === "&") {
    const left = scalarToText(a), right = scalarToText(b);
    if (typeof left !== "string") return left;
    if (typeof right !== "string") return right;
    return left.length + right.length > 32_767 ? scalarError("#VALUE!") : scalarText(left + right);
  }
  const left = scalarToNumber(a), right = scalarToNumber(b);
  if (typeof left !== "number") return left;
  if (typeof right !== "number") return right;
  if (operator === "/" && right === 0) return scalarError("#DIV/0!");
  if (operator === "^" && left === 0 && right === 0) return scalarError("#NUM!");
  return scalarNumber(operator === "+" ? left + right : operator === "-" ? left - right : operator === "*" ? left * right : operator === "/" ? left / right : left ** right);
}
