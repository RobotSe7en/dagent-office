import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { compileNumberFormat } from "../src/workbook/number-format.js";

describe('native number-format source compiler', () => {
  it('compiles every independent native format probe without claiming display evaluation', () => {
    const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-number-format-semantics.json', 'utf8')) as { cases: { samples: { code: string }[] }[] };
    let count = 0;
    for (const entry of evidence.cases) for (const sample of entry.samples) { expect(compileNumberFormat(sample.code).length).toBeGreaterThan(0); count++; }
    expect(count).toBe(130);
  });
  it('separates only unquoted, unescaped section delimiters', () => {
    const format = compileNumberFormat('0.00;[Red](0.00);"—;";"收入 "@');
    expect(format).toHaveLength(4); expect(format[1]?.color).toEqual({ kind: 'named', name: 'red' });
    expect(format[2]?.tokens).toEqual([{ kind: 'literal', text: '—;' }]);
    expect(format[3]?.tokens).toEqual([{ kind: 'literal', text: '收入 ' }, { kind: 'text' }]);
    expect(compileNumberFormat('0\\;0')).toHaveLength(1); expect(compileNumberFormat(';;;').map(section => section.tokens)).toEqual([[], [], [], []]);
    expect(Object.isFrozen(format[0]?.tokens[0])).toBe(true);
  });
  it('preserves measured spacing and fill characters including Unicode and semicolons', () => {
    expect(compileNumberFormat('_)* 0')[0]?.tokens).toEqual([{ kind: 'space', character: ')' }, { kind: 'fill', character: ' ' }, { kind: 'digit', value: '0' }]);
    expect(compileNumberFormat('_😀*;0')).toHaveLength(1);
    expect(compileNumberFormat('_😀*;0')[0]?.tokens.slice(0, 2)).toEqual([{ kind: 'space', character: '😀' }, { kind: 'fill', character: ';' }]);
    expect(compileNumberFormat('\\"0\\%')[0]?.tokens).toEqual([{ kind: 'literal', text: '"' }, { kind: 'digit', value: '0' }, { kind: 'literal', text: '%' }]);
  });
  it('retains conditions, native palette identities, currencies and locale flags', () => {
    const format = compileNumberFormat('[Color56][>=1e3][$¥-804]#,##0.00;[Red][<0](0)');
    expect(format[0]).toMatchObject({ condition: { operator: '>=', value: 1000 }, color: { kind: 'indexed', index: 56 }, locale: 0x804 });
    expect(format[0]?.tokens[0]).toEqual({ kind: 'literal', text: '¥' });
    expect(format[1]?.condition).toEqual({ operator: '<', value: 0 });
    expect(compileNumberFormat('[<>0]0')[0]?.condition?.operator).toBe('<>');
    expect(compileNumberFormat('[$-F800]dddd')[0]?.locale).toBe(0xf800);
  });
  it('distinguishes date fields, elapsed units, exponents and literal date letters', () => {
    expect(compileNumberFormat('yyyy-mm-dd')[0]?.tokens).toEqual([{ kind: 'date', field: 'y', length: 4 }, { kind: 'literal', text: '-' }, { kind: 'date', field: 'm', length: 2 }, { kind: 'literal', text: '-' }, { kind: 'date', field: 'd', length: 2 }]);
    expect(compileNumberFormat('[hh]:mm:ss.00 AM/PM')[0]?.tokens).toContainEqual({ kind: 'elapsed', field: 'h', length: 2 });
    expect(compileNumberFormat('h:mm a/p')[0]?.tokens.at(-1)).toEqual({ kind: 'dayPeriod', abbreviated: true, upperCase: false });
    expect(compileNumberFormat('0.00E+00')[0]?.tokens).toContainEqual({ kind: 'exponent', positiveSign: true, upperCase: true });
    expect(compileNumberFormat('"yyyy"0\\M')[0]?.tokens).toEqual([{ kind: 'literal', text: 'yyyy' }, { kind: 'digit', value: '0' }, { kind: 'literal', text: 'M' }]);
    expect(compileNumberFormat('General')[0]?.tokens).toEqual([{ kind: 'general' }]);
  });
  it('rejects unsupported extensions, invalid declarations, broken escapes and excessive sources explicitly', () => {
    for (const code of ['"oops', '0\\', '0_', '0*', '[Red', '[Red][Blue]0', '[Color57]0', '[>1e999]0', '[>0][<1]0', '0;0;[>0]0', '0;0;0;@;0']) expect(() => compileNumberFormat(code), code).toThrow();
    for (const code of ['[DBNum1]0', '[unknown]0', 'eeee']) expect(() => compileNumberFormat(code), code).toThrowError(expect.objectContaining({ code: 'UnsupportedFeature' }));
    expect(() => compileNumberFormat('0'.repeat(2049))).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    expect(() => compileNumberFormat('"\ud800"')).toThrow();
    expect(() => compileNumberFormat('[$\ud800]0')).toThrow();
  });
});
