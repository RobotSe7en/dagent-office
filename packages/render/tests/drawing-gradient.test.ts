import { describe, expect, it } from 'vitest';
import { drawingGradientLine } from '../src/drawing-gradient.js';
import type { DrawingGradientLine } from '../src/drawing-gradient.js';

function position(line: DrawingGradientLine, x: number, y: number): number {
  const dx = line.x2 - line.x1, dy = line.y2 - line.y1;
  return ((x - line.x1) * dx + (y - line.y1) * dy) / (dx * dx + dy * dy);
}
describe('DrawingML physical linear-gradient geometry', () => {
  it('distinguishes scaled diagonal angles from fixed physical angles in a nonsquare region', () => {
    const bounds = { x: 20, y: 10, width: 200, height: 100 };
    const scaled = drawingGradientLine(bounds, 45, true), fixed = drawingGradientLine(bounds, 45, false);
    for (const [key, value] of Object.entries({ x1: 20, y1: 10, x2: 220, y2: 110 })) expect(scaled[key as keyof DrawingGradientLine]).toBeCloseTo(value, 10);
    for (const [key, value] of Object.entries({ x1: 45, y1: -15, x2: 195, y2: 135 })) expect(fixed[key as keyof DrawingGradientLine]).toBeCloseTo(value, 10);
    expect(position(scaled, 220, 10)).toBeCloseTo(.8, 10); expect(position(fixed, 220, 10)).toBeCloseTo(2 / 3, 10);
    for (const line of [scaled, fixed]) {
      expect(position(line, 20, 10)).toBeCloseTo(0, 10); expect(position(line, 220, 110)).toBeCloseTo(1, 10); expect(position(line, 120, 60)).toBeCloseTo(.5, 10);
    }
  });
  it('covers the entire region at arbitrary angles and reverses stops at the opposite angle', () => {
    for (const [width, height] of [[100, 100], [200, 100], [100, 200], [1, 1000]]) for (const scaled of [false, true]) for (const angle of [0, 30, 45, 60, 90, 135, 210, 270, 360]) {
      const bounds = { x: -20, y: 15, width: width!, height: height! }, line = drawingGradientLine(bounds, angle, scaled), reverse = drawingGradientLine(bounds, angle + 180, scaled);
      const corners = [[-20, 15], [-20 + width!, 15], [-20, 15 + height!], [-20 + width!, 15 + height!]].map(([x, y]) => position(line, x!, y!));
      expect(Math.min(...corners)).toBeCloseTo(0, 10); expect(Math.max(...corners)).toBeCloseTo(1, 10);
      expect(reverse.x1).toBeCloseTo(line.x2, 9); expect(reverse.y1).toBeCloseTo(line.y2, 9);
    }
  });
  it('handles degenerate regions without creating nonfinite SVG coordinates and rejects invalid extents', () => {
    expect(drawingGradientLine({ x: 2, y: 3, width: 0, height: 0 }, 0, true)).toEqual({ x1: 2, y1: 3, x2: 2, y2: 3 });
    for (const width of [-1, NaN, Infinity]) expect(() => drawingGradientLine({ x: 0, y: 0, width, height: 1 }, 0, true)).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
    expect(() => drawingGradientLine({ x: Number.MAX_VALUE, y: 0, width: Number.MAX_VALUE, height: 1 }, 0, true)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
  it('keeps fixed fills in page axes across rotations and single or double reflections', () => {
    const bounds = { x: 0, y: 0, width: 200, height: 100 };
    expect(drawingGradientLine(bounds, 0, false, { rotationDegrees: 90, flipH: false, flipV: false })).toEqual({ x1: 100, y1: 100, x2: 100, y2: 0 });
    for (const rotationDegrees of [-90, 45, 90, 135, 180, 270, 360]) for (const flipH of [false, true]) for (const flipV of [false, true]) {
      const line = drawingGradientLine(bounds, 0, false, { rotationDegrees, flipH, flipV }), angle = rotationDegrees*Math.PI/180;
      const world = (x: number, y: number) => ({ x: (x-100)*(flipH ? -1 : 1)*Math.cos(angle) - (y-50)*(flipV ? -1 : 1)*Math.sin(angle), y: (x-100)*(flipH ? -1 : 1)*Math.sin(angle) + (y-50)*(flipV ? -1 : 1)*Math.cos(angle) });
      const first = world(line.x1, line.y1), last = world(line.x2, line.y2);
      expect(first.y).toBeCloseTo(0, 10); expect(last.y).toBeCloseTo(0, 10);
      expect(last.x-first.x).toBeCloseTo((200*Math.abs(Math.cos(angle))+100*Math.abs(Math.sin(angle)))*(flipH ? -1 : 1), 9);
      expect(Object.isFrozen(line)).toBe(true);
    }
    expect(() => drawingGradientLine(bounds, 0, false, { rotationDegrees: Infinity, flipH: false, flipV: false })).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
    expect(drawingGradientLine({ x: 2, y: 3, width: 0, height: 0 }, 45, true, { rotationDegrees: 45, flipH: true, flipV: false })).toEqual({ x1: 2, y1: 3, x2: 2, y2: 3 });
  });
});
