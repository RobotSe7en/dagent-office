import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { compileNumberFormat } from "../src/workbook/number-format.js";
import { formatFractionSection } from "../src/workbook/format-fraction.js";

const display = (value: number, code: string) => formatFractionSection(value, compileNumberFormat(code)[0]!);
const text = (value: number, code: string) => display(value, code).parts.map(part => part.kind === 'text' ? part.text : part.kind === 'space' ? part.character === '0' ? '\u2007' : ' ' : '<fill>').join('');
describe('fraction number-format sections', () => {
  it('matches independent native fractions, including whole numbers, zero, padding and rollover', () => {
    const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-number-format-semantics.json', 'utf8')) as { cases: { samples: { value: number; code: string; text: string }[] }[] };
    let checked = 0;
    for (const entry of evidence.cases) for (const sample of entry.samples) {
      if (!['# ?/?', '# ??/??', '# ?/8', '?/?', '0/0', '# ?/16', '# ?/10', '# ??/100'].includes(sample.code)) continue;
      expect(text(sample.value, sample.code), `${sample.value} ${sample.code}`).toBe(sample.text); checked++;
    }
    expect(checked).toBe(60);
  });
  it('keeps fixed denominators unreduced and supports whole or improper fractions', () => {
    expect(text(.5, '# ?/8')).toBe(' 4/8');
    expect(text(1.5, '?/8')).toBe('12/8');
    expect(text(1.99, '# ?/?')).toBe('2 \u2007 \u2007');
    expect(display(.01, '# ?/?').zero).toBe(true);
    expect(display(1.99, '# ?/?').zero).toBe(false);
    expect(text(-1.3, '"("# ??/??")"')).toBe('(1 \u20073/10)');
    expect(text(1e308, '# ?/?%').slice(0, 311)).toBe('1' + '0'.repeat(310));
  });
  it('does not conflate quoted digits with a native fixed denominator', () => {
    expect(compileNumberFormat('# ?/100')[0]?.tokens.at(-1)).toEqual({ kind: 'integer', digits: '100' });
    expect(compileNumberFormat('# ?/"100"')[0]?.tokens.at(-1)).toEqual({ kind: 'literal', text: '100' });
    for (const code of ['# ?/"8"', '# ?/\\8', '# ?/?/?', '# ?/', '# /?', '# ?/8?', '0.00', '#,##0 ?/?', '# 0/8']) expect(() => display(1, code), code).toThrowError(expect.objectContaining({ code: 'UnsupportedFeature' }));
  });
  it('keeps explicit padding and fill as measured characters and bounds large denominator fields', () => {
    expect(display(1, '_)* # ?/?').parts.slice(0, 2)).toEqual([{ kind: 'space', character: ')' }, { kind: 'fill', character: ' ' }]);
    expect(text(.125, '# ?/' + '?'.repeat(500)).endsWith('1/' + '\u2007'.repeat(499) + '8')).toBe(true);
    expect(() => display(Infinity, '# ?/?')).toThrow();
  });
});
