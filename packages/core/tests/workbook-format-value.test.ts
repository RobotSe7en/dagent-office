import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { compileNumberFormat } from "../src/workbook/number-format.js";
import { formatWorkbookValue } from "../src/workbook/format-value.js";

const separators = { decimal: '.', group: ',', primaryGroupSize: 3, secondaryGroupSize: 3 };
const display = (value: number | string, code: string) => formatWorkbookValue(value, compileNumberFormat(code), separators);
interface Sample { value: number | string; code: string; text: string }
describe('workbook scalar number-format dispatch', () => {
  it('matches all implemented numeric/text results in the independent native evidence', () => {
    const numbers = JSON.parse(readFileSync('docs/evidence/xlsx-number-format-semantics.json', 'utf8')) as { cases: { samples: Sample[] }[] };
    const selection = JSON.parse(readFileSync('docs/evidence/xlsx-format-selection.json', 'utf8')) as { samples: Sample[] };
    let checked = 0, deferred = 0;
    for (const sample of [...numbers.cases.flatMap(entry => entry.samples), ...selection.samples]) {
      const result = display(sample.value, sample.code);
      if (result.kind === 'general' || result.kind === 'date') { deferred++; continue; }
      expect(result.kind, sample.code).toBe('parts'); if (result.kind !== 'parts') throw new Error('Missing display parts');
      const text = result.parts.map(part => part.kind === 'text' ? part.text : part.kind === 'space' ? part.character === '0' ? '\u2007' : ' ' : '<fill>').join('');
      expect(text, `${sample.value} ${sample.code}`).toBe(sample.text); checked++;
    }
    expect(checked).toBe(151); expect(deferred).toBe(26);
  });
  it('retains color and locale with literal or numeric content without mutating its scalar', () => {
    expect(display(-12.5, '0.0;[Red][$¥-804](0.0)')).toEqual({ kind: 'parts', parts: [{ kind: 'text', text: '¥(12.5)' }], color: { kind: 'named', name: 'red' }, locale: 0x804 });
    expect(display('001', '0;0;0;[Blue]"ID "@')).toEqual({ kind: 'parts', parts: [{ kind: 'text', text: 'ID 001' }], color: { kind: 'named', name: 'blue' } });
    expect(display(-1, '"constant"')).toEqual({ kind: 'parts', parts: [{ kind: 'text', text: 'constant' }] });
    expect(display(-.001, '0.00')).toEqual({ kind: 'parts', parts: [{ kind: 'text', text: '0.00' }] });
    expect(display(-.001, '0.00;(0.00)')).toEqual({ kind: 'parts', parts: [{ kind: 'text', text: '(0.00)' }] });
  });
  it('never reports unresolved General fitting, calendar display or unmatched conditions as completed formatting', () => {
    expect(display(60, 'yyyy-mm-dd')).toMatchObject({ kind: 'date', value: 60 });
    expect(display(1.25, '[h]:mm')).toMatchObject({ kind: 'date', value: 1.25 });
    expect(display(123456789012345, 'General')).toMatchObject({ kind: 'general', value: 123456789012345 });
    expect(display(5, '[>10]0;[<0]0')).toEqual({ kind: 'unmatchedCondition' });
    expect(() => display('x'.repeat(32768), '0')).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
});
