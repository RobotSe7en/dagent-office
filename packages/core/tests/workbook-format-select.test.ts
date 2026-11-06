import { describe, expect, it } from "vitest";
import { compileNumberFormat } from "../src/workbook/number-format.js";
import { formatLiteralSection, selectNumberFormat } from "../src/workbook/format-select.js";

const select = (value: number | string, code: string) => selectNumberFormat(value, compileNumberFormat(code));
describe('source number-format section selection', () => {
  it('selects before rounding and applies automatic negative signs only to a shared numeric section', () => {
    expect(select(-.001, '0.00')).toMatchObject({ kind: 'number', index: 0, automaticNegative: true });
    expect(select(-.001, '0.00;(0.00);"zero"')).toMatchObject({ kind: 'number', index: 1, automaticNegative: false });
    expect(select(-0, '0.00;(0.00);"zero"')).toMatchObject({ kind: 'number', index: 2, automaticNegative: false });
    expect(select(0, '0.00;(0.00)')).toMatchObject({ kind: 'number', index: 0 });
    expect(select(-1, '0;')).toMatchObject({ kind: 'number', index: 1, section: { tokens: [] } });
    expect(select(0, '0;;')).toMatchObject({ kind: 'number', index: 2, section: { tokens: [] } });
  });
  it('separates final text sections from the number section count', () => {
    expect(select(-1, '0;@')).toMatchObject({ kind: 'number', index: 0, automaticNegative: true });
    expect(select(-1, '0;"neg"0;@')).toMatchObject({ kind: 'number', index: 1, automaticNegative: false });
    expect(select(-1, '@')).toEqual({ kind: 'general' });
    expect(select('001', '@')).toMatchObject({ kind: 'text', index: 0 });
    expect(select('001', '0.00')).toEqual({ kind: 'plainText' });
    expect(select('001', '0.00;0.00;0;')).toMatchObject({ kind: 'text', index: 3, section: { tokens: [] } });
    expect(select('001', '0.00;0.00;0;"receipt "@')).toMatchObject({ kind: 'text', index: 3 });
  });
  it('evaluates ordered conditions on the original signed scalar and distinguishes no match', () => {
    const code = '[>=10]"A"0;[<0]"B"0;"C"0';
    expect(select(10, code)).toMatchObject({ kind: 'number', index: 0 });
    expect(select(9.999, code)).toMatchObject({ kind: 'number', index: 2 });
    expect(select(-1, code)).toMatchObject({ kind: 'number', index: 1, automaticNegative: false });
    expect(select(30, '[>10]0;[>20]0;0')).toMatchObject({ kind: 'number', index: 0 });
    expect(select(5, '[>10]0;[>20]0')).toEqual({ kind: 'unmatchedCondition' });
    expect(select(0, '[>10]0')).toEqual({ kind: 'unmatchedCondition' });
    for (const [operator, value, expected] of [['<', 1, 0], ['<=', 2, 0], ['>', 3, 0], ['>=', 2, 0], ['=', 2, 0], ['<>', 3, 0], ['<>', 2, 1]] as const) expect(select(value, `[${operator}2]0;0`)).toMatchObject({ kind: 'number', index: expected });
  });
  it('preserves text, literals, blank sections and measured padding without parsing numeric-looking text', () => {
    expect(formatLiteralSection('001😀禰󠄀', compileNumberFormat('"ID "@')[0]!)).toEqual([{ kind: 'text', text: 'ID 001😀禰󠄀' }]);
    expect(formatLiteralSection(undefined, compileNumberFormat('"—"')[0]!)).toEqual([{ kind: 'text', text: '—' }]);
    expect(formatLiteralSection('hidden', compileNumberFormat(';;;')[3]!)).toEqual([]);
    expect(formatLiteralSection('001', compileNumberFormat('_)* @')[0]!)).toEqual([{ kind: 'space', character: ')' }, { kind: 'fill', character: ' ' }, { kind: 'text', text: '001' }]);
    expect(() => formatLiteralSection('1'.repeat(32767), compileNumberFormat('@@')[0]!)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
  it('rejects ambiguous or unsupported arrangements instead of silently replacing them with General', () => {
    for (const code of ['@;0', '0;[<0]0', '[<0]0;0;0', '[>0]@']) expect(() => select(1, code)).toThrowError(expect.objectContaining({ code: 'UnsupportedFeature' }));
    expect(() => select(Infinity, '0')).toThrow();
    expect(() => formatLiteralSection('x', compileNumberFormat('0')[0]!)).toThrowError(expect.objectContaining({ code: 'UnsupportedFeature' }));
  });
});
