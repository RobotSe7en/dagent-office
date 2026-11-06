import { describe, expect, it } from "vitest";
import { measureNumberDisplay } from "../src/workbook/number-layout.js";

describe('measured workbook numeric display parts', () => {
  it('reserves exact character widths for question placeholders and explicit underscore spacing', () => {
    const widths: Record<string, number> = { '0': 7, ')': 4, '12': 12 };
    expect(measureNumberDisplay([{ kind: 'space', character: ')' }, { kind: 'text', text: '12' }, { kind: 'space', character: '0' }], 23, text => widths[text]!)).toEqual({
      parts: [{ kind: 'space', width: 4 }, { kind: 'text', text: '12', width: 12 }, { kind: 'space', width: 7 }], width: 23, overflow: false,
    });
  });
  it('uses the complete shaped fill width and leaves remaining fractions of a glyph unfilled', () => {
    const measure = (text: string) => /^f+$/.test(text) ? text.length === 1 ? 10 : text.length * 5 : text.length * 10;
    const output = measureNumberDisplay([{ kind: 'text', text: '$' }, { kind: 'fill', character: 'f' }, { kind: 'text', text: '12' }], 63, measure);
    expect(output).toEqual({ parts: [{ kind: 'text', text: '$', width: 10 }, { kind: 'text', text: 'ffffff', width: 30 }, { kind: 'text', text: '12', width: 20 }], width: 60, overflow: false });
    expect(measureNumberDisplay([{ kind: 'text', text: 'A' }, { kind: 'text', text: 'V' }], 15, text => text === 'AV' ? 15 : 10)).toEqual({ parts: [{ kind: 'text', text: 'AV', width: 15 }], width: 15, overflow: false });
  });
  it('does not repeat zero-width glyphs or conceal fixed content overflow', () => {
    const result = measureNumberDisplay([{ kind: 'fill', character: '\u200d' }, { kind: 'text', text: '1' }], 100, text => text === '\u200d' || !text ? 0 : 10);
    expect(result.width).toBe(10); expect(result.parts[0]).toEqual({ kind: 'text', text: '', width: 0 });
    expect(measureNumberDisplay([{ kind: 'fill', character: ' ' }, { kind: 'text', text: '123' }], 2, text => text.length).overflow).toBe(true);
  });
  it('keeps temporary shaping strings close to the visible fill width', () => {
    let longest = 0;
    const result = measureNumberDisplay([{ kind: 'fill', character: '.' }], 17, text => { longest = Math.max(longest, text.length); return text.length; });
    expect(result.width).toBe(17); expect(longest).toBeLessThanOrEqual(34);
  });
  it('bounds expanded fill before publishing, including supplementary Unicode characters', () => {
    expect(() => measureNumberDisplay([{ kind: 'fill', character: 'x' }], 32768, text => text.length)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    const exact = measureNumberDisplay([{ kind: 'text', text: '1' }, { kind: 'fill', character: '😀' }], 16384, text => [...text].length);
    expect(exact.width).toBe(16384); expect(exact.parts[1]?.kind === 'text' && exact.parts[1].text.length).toBe(32766);
    expect(() => measureNumberDisplay([{ kind: 'text', text: '1' }, { kind: 'fill', character: '😀' }], 16385, text => [...text].length)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    expect(measureNumberDisplay([{ kind: 'fill', character: ' ' }, { kind: 'fill', character: '.' }], 10, text => text.length).parts).toEqual([{ kind: 'text', text: '..........', width: 10 }]);
    expect(() => measureNumberDisplay([{ kind: 'text', text: '1' }], 10, () => Infinity)).toThrow();
  });
});
