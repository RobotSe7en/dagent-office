import { DocumentError } from "./errors.js";
import { presetDefinitions } from "./presentation-geometry-data.js";

export type PresentationPathCommand =
  | { readonly kind: "move" | "line"; readonly x: number; readonly y: number }
  | { readonly kind: "arc"; readonly radiusX: number; readonly radiusY: number; readonly startAngle: number; readonly sweepAngle: number }
  | { readonly kind: "close" };
export interface PresentationPresetGeometry {
  readonly textRect: Readonly<{ x: number; y: number; width: number; height: number }> | null;
  readonly connections: readonly Readonly<{ x: number; y: number; angle: number }>[];
  readonly paths: readonly Readonly<{ fill: boolean; stroke: boolean; commands: readonly PresentationPathCommand[] }>[];
}
const finite = (n: number): number => {
  if (!Number.isFinite(n)) throw new DocumentError("UnsupportedFeature", "The preset geometry has a nonfinite result.");
  return n;
};

/** Evaluate only this SDK's selected, pinned standard definitions. Source input
 * can supply integer adjustments, never guide formulas or executable geometry. */
export function evaluatePresentationPreset(name: string, width: number, height: number, adjustments: Readonly<Record<string, number>> = {}): PresentationPresetGeometry {
  if (typeof width !== "number" || typeof height !== "number" || !Number.isFinite(width) || !Number.isFinite(height) || width < 0 || height < 0) throw new DocumentError("InvalidOperation", "Preset dimensions must be finite nonnegative numbers.");
  const mirror = name === "upArrow", key = mirror ? "downArrow" : name;
  if (!Object.hasOwn(presetDefinitions, key)) throw new DocumentError("UnsupportedFeature", "This preset geometry is not supported.");
  const definition = presetDefinitions[key]!;
  const values = new Map<string, number>([["w", width], ["h", height], ["l", 0], ["t", 0], ["r", width], ["b", height], ["hc", width / 2], ["vc", height / 2], ["ss", Math.min(width, height)], ["ls", Math.max(width, height)], ["cd2", 10800000], ["cd4", 5400000], ["cd8", 2700000], ["3cd4", 16200000]]);
  for (const divisor of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) { values.set("wd" + divisor, width / divisor); values.set("hd" + divisor, height / divisor); values.set("ssd" + divisor, Math.min(width, height) / divisor); }
  const read = (token: string): number => {
    if (/^-?\d+(?:\.\d+)?$/.test(token)) return finite(Number(token));
    const n = values.get(token);
    if (n === undefined) throw new DocumentError("InvalidDocument", "A built-in geometry guide is undefined.");
    return n;
  };
  const formula = (source: string): number => {
    const [op, ...tokens] = source.trim().split(/\s+/), args = tokens.map(read), [a = 0, b = 0, c = 0] = args;
    let result: number;
    switch (op) {
      case "val": result = a; break;
      case "pin": result = Math.max(a, Math.min(b, c)); break;
      case "*/": result = c === 0 ? 0 : a * b / c; break;
      case "+-": result = a + b - c; break;
      case "+/": result = c === 0 ? 0 : (a + b) / c; break;
      case "?:": result = a > 0 ? b : c; break;
      case "cos": result = a * Math.cos(b * Math.PI / 10800000); break;
      case "sin": result = a * Math.sin(b * Math.PI / 10800000); break;
      default: throw new DocumentError("UnsupportedFeature", "This built-in geometry formula is not supported.");
    }
    return finite(result);
  };
  const names = new Set(definition.adjustments.map(([id]) => id));
  for (const [id, source] of definition.adjustments) values.set(id, formula(source));
  for (const id in adjustments) if (Object.hasOwn(adjustments, id)) {
    if (!names.has(id) || !Number.isSafeInteger(adjustments[id])) throw new DocumentError("UnsupportedFeature", "A preset adjustment is unknown or is not a safe integer constant.");
    values.set(id, adjustments[id]!);
  }
  for (const [id, source] of definition.guides) values.set(id, formula(source));
  const point = (x: string, y: string): Readonly<{ x: number; y: number }> => Object.freeze({ x: read(x), y: mirror ? height - read(y) : read(y) });
  const text = definition.text?.map(read);
  const textRect = text ? Object.freeze({ x: text[0]!, y: mirror ? height - text[3]! : text[1]!, width: text[2]! - text[0]!, height: text[3]! - text[1]! }) : null;
  const connectionTemplates = mirror ? [definition.connections[2]!, definition.connections[1]!, definition.connections[0]!, definition.connections[3]!] : definition.connections;
  const connections = Object.freeze(connectionTemplates.map(([angle, x, y]) => Object.freeze({ ...point(x, y), angle: (((mirror ? -read(angle) : read(angle)) / 60000) % 360 + 360) % 360 })));
  const paths = Object.freeze(definition.paths.map(path => Object.freeze({ fill: path.fill, stroke: path.stroke,
    commands: Object.freeze(path.commands.map((command): PresentationPathCommand => {
      if (command[0] === "Z") return Object.freeze({ kind: "close" });
      if (command[0] === "M" || command[0] === "L") return Object.freeze({ kind: command[0] === "M" ? "move" : "line", ...point(command[1]!, command[2]!) });
      if (command[0] === "A") return Object.freeze({ kind: "arc", radiusX: read(command[1]!), radiusY: read(command[2]!), startAngle: (mirror ? -read(command[3]!) : read(command[3]!)) / 60000, sweepAngle: (mirror ? -read(command[4]!) : read(command[4]!)) / 60000 });
      throw new DocumentError("UnsupportedFeature", "This built-in geometry path command is not supported.");
    })) })));
  return Object.freeze({ textRect, connections, paths });
}
