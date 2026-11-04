import { describe, expect, it } from "vitest";
import { compileNumberFormat } from "../src/workbook/number-format.js";
import { formatScientificSection } from "../src/workbook/format-scientific.js";
const separators = { decimal: '.', group: ',', primaryGroupSize: 3, secondaryGroupSize: 3 };
function display(value: number, code: string) { return formatScientificSection(value, compileNumberFormat(code)[0]!, separators).parts; }
describe('scientific magnitude display', () => {
  it.each([
    [1234, '0.00E+00', '1.23E+03'], [1234, '00.0E+0', '12.3E+2'], [1234, '##0.0E+0', '1.2E+3'],
    [0, '0.00E+00', '0.00E+00'], [9.999, '0.00E+00', '1.00E+01'], [999.99, '##0.0E+00', '1.0E+03'],
    [1e308, '0.00E+00', '1.00E+308'], [Number.MIN_VALUE, '0.00E+00', '5.00E-324'],
    [1234, '0.0e-00', '1.2e03'], [.001, '0.0E-00', '1.0E-03'], [.5, '0.00E+00%', '5.00E+01%'],
    [.0001234, '##0.0E+0', '123.4E-6'], [-1234, '0.00E+00', '1.23E+03'],
  ] as const)('formats %s with %s', (value, code, expected) => { expect(display(value, code)).toEqual([{ kind: 'text', text: expected }]); });
  it('preserves layout tokens and rejects unsupported field combinations', () => {
    expect(display(12, '0.0E+00_)')).toEqual([{ kind: 'text', text: '1.2E+01' }, { kind: 'space', character: ')' }]);
    for (const code of ['0.0', '0.0E+##', '#.0E+0', '0.0E+00E+00', '0.0E+00@']) expect(() => display(1, code)).toThrowError(expect.objectContaining({ code: 'UnsupportedFeature' }));
  });
});
