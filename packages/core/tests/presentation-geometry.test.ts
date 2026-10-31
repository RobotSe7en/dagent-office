import { expect, test } from "vitest";
import { evaluatePresentationPreset, PresentationPresetSchema } from "@dagent-office/core";

for (const preset of [...PresentationPresetSchema.options, "straightConnector1", "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5"]) {
  test("evaluates selected preset geometry without platform or XML input: " + preset, () => {
    const first = evaluatePresentationPreset(preset, 195, 95);
    expect(evaluatePresentationPreset(preset, 195, 95)).toEqual(first);
    expect(first.paths.length).toBeGreaterThan(0);
    for (const path of first.paths) for (const command of path.commands) for (const v of Object.values(command)) if (typeof v === "number") expect(Number.isFinite(v)).toBe(true);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first.paths[0]!.commands)).toBe(true);
  });
}
test("retains the standard rectangle text area and indexed connection sites", () => {
  const shape = evaluatePresentationPreset("rect", 200, 100);
  expect(shape.textRect).toEqual({ x: 0, y: 0, width: 200, height: 100 });
  expect(shape.connections).toEqual([{ x: 100, y: 0, angle: 270 }, { x: 0, y: 50, angle: 180 }, { x: 100, y: 100, angle: 90 }, { x: 200, y: 50, angle: 0 }]);
  expect(shape.paths[0]!.commands).toEqual([{ kind: "move", x: 0, y: 0 }, { kind: "line", x: 200, y: 0 }, { kind: "line", x: 200, y: 100 }, { kind: "line", x: 0, y: 100 }, { kind: "close" }]);
});
test("applies bounded corner adjustments and retains quarter arcs instead of rasterizing them", () => {
  const sharp = evaluatePresentationPreset("roundRect", 200, 100, { adj: 0 });
  expect(sharp.textRect).toEqual({ x: 0, y: 0, width: 200, height: 100 });
  expect(sharp.paths[0]!.commands.filter(c => c.kind === "arc")).toHaveLength(4);
  const round = evaluatePresentationPreset("roundRect", 200, 100, { adj: 50000 });
  const arc = round.paths[0]!.commands.find(c => c.kind === "arc")!;
  expect(arc).toMatchObject({ radiusX: 50, radiusY: 50 });
  expect(evaluatePresentationPreset("roundRect", 200, 100, { adj: 999999 })).toEqual(round);
});
test("derives up-arrow geometry by vertical reflection and retains top/left/bottom/right sites", () => {
  const down = evaluatePresentationPreset("downArrow", 200, 100), up = evaluatePresentationPreset("upArrow", 200, 100);
  expect(up.textRect).toEqual({ ...down.textRect!, y: 100 - down.textRect!.y - down.textRect!.height });
  expect(up.connections[0]).toEqual({ x: 100, y: 0, angle: 270 });
  expect(up.paths[0]!.commands[0]).toEqual({ kind: "move", x: 0, y: 50 });
});
test("keeps elbow path adjustments and handles zero-width or zero-height straight lines", () => {
  expect(evaluatePresentationPreset("bentConnector4", 200, 100, { adj1: 25000, adj2: 75000 }).paths[0]!.commands).toEqual([
    { kind: "move", x: 0, y: 0 }, { kind: "line", x: 50, y: 0 }, { kind: "line", x: 50, y: 75 }, { kind: "line", x: 200, y: 75 }, { kind: "line", x: 200, y: 100 },
  ]);
  expect(evaluatePresentationPreset("line", 200, 0).textRect).toBeNull();
  expect(evaluatePresentationPreset("straightConnector1", 0, 100).paths[0]!.commands.at(-1)).toEqual({ kind: "line", x: 0, y: 100 });
});
test("rejects unsupported definitions, arbitrary formulas and invalid dimensions", () => {
  for (const action of [() => evaluatePresentationPreset("script", 10, 10), () => evaluatePresentationPreset("rect", 10, 10, { execute: 1 }), () => evaluatePresentationPreset("roundRect", 10, 10, { adj: 1.5 })]) expect(action).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  for (const n of [-1, NaN, Infinity]) expect(() => evaluatePresentationPreset("rect", n, 100)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
});
