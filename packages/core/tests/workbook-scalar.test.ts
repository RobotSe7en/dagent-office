import { describe, expect, it } from 'vitest';
import { createWorkbookScalarEvaluator } from '../src/index.js';
import type { WorkbookValue } from '../src/index.js';

import { scalarWorkbook } from './workbook-scalar-fixture.js';

const location = { sheetId: 'sheet-1', row: 0, column: 4 };
describe('read-only workbook scalar kernel', () => {
  it('evaluates typed operators, logical control and independent rounding examples', () => {
    const evaluator = createWorkbookScalarEvaluator(scalarWorkbook([]));
    const cases: [string, WorkbookValue][] = [
      ['=1+2*3', { type: 'number', value: 7 }], ['=-2^2', { type: 'number', value: 4 }], ['=2^3^2', { type: 'number', value: 64 }],
      ['=TRUE=1', { type: 'boolean', value: false }], ['="A"&TRUE', { type: 'text', value: 'ATRUE' }], ['=""+1', { type: 'error', value: '#VALUE!' }],
      ['=IF(TRUE,1,1/0)', { type: 'number', value: 1 }], ['=IF(FALSE,1/0,2)', { type: 'number', value: 2 }],
      ['=ROUND(1.005,2)', { type: 'number', value: 1.01 }], ['=ROUND(-1.005,2)', { type: 'number', value: -1.01 }],
      ['=ROUNDUP(-1.001,2)', { type: 'number', value: -1.01 }], ['=ROUNDDOWN(-1.009,2)', { type: 'number', value: -1 }],
      ['=ROUND(125,-1)', { type: 'number', value: 130 }], ['=TRUNC(9.99,1)', { type: 'number', value: 9.9 }],
      ['=ROUND(1e308,300)', { type: 'number', value: 1e308 }], ['=ROUND(1e-308,308)', { type: 'number', value: 1e-308 }],
      ['=ROUND(12,-1e10)', { type: 'number', value: 0 }], ['=ROUNDUP(12,-1e10)', { type: 'error', value: '#NUM!' }],
      ['=MOD(1e308,3)', { type: 'number', value: 2 }], ['=MOD(-5,3)', { type: 'number', value: 1 }],
    ];
    for (const [expression, value] of cases) {
      const result = evaluator.evaluate(expression, location); expect(result, expression).toMatchObject({ kind: 'value', documentId: 'scalar-kernel', revision: 0, value });
      expect(Object.isFrozen(result)).toBe(true); if (result.kind === 'value') expect(Object.isFrozen(result.value)).toBe(true);
    }
    evaluator.dispose(); expect(() => evaluator.evaluate('=1', location)).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
  });
  it('calculates current dependencies, ignores deliberately wrong caches and preserves source', () => {
    const model = scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 2 } },
      { row: 0, column: 1, value: { type: 'number', value: 999 }, formula: '=A1+1' },
      { row: 0, column: 2, value: { type: 'number', value: 999 }, formula: '=B1*2' }]);
    const before = model.serialize(), evaluator = createWorkbookScalarEvaluator(model);
    expect(evaluator.cell({ ...location, column: 2 })).toMatchObject({ kind: 'value', value: { type: 'number', value: 6 }, readCells: 3 });
    expect(evaluator.evaluate('=AND(C1=6,NOT(FALSE))', location)).toMatchObject({ kind: 'value', value: { type: 'boolean', value: true } });
    expect(evaluator.cell({ ...location, row: 10 })).toMatchObject({ kind: 'value', value: { type: 'blank' } });
    expect(evaluator.evaluate('=E11', location)).toMatchObject({ kind: 'value', value: { type: 'number', value: 0 } });
    expect(model.getCell('sheet-1', 0, 2)!.value).toEqual({ type: 'number', value: 999 });
    expect(model.serialize()).toEqual(before); evaluator.dispose();
  });
  it('keeps unsupported capability separate from legal formula errors and short circuits only conditional branches', () => {
    const evaluator = createWorkbookScalarEvaluator(scalarWorkbook([]));
    expect(evaluator.evaluate('=IF(TRUE,7,FUTUREFUNC())', location)).toMatchObject({ kind: 'value', value: { type: 'number', value: 7 } });
    expect(evaluator.evaluate('=IFERROR(FUTUREFUNC(),7)', location)).toMatchObject({ kind: 'unavailable', reason: 'unsupported-function', functionName: 'FUTUREFUNC' });
    expect(evaluator.evaluate('=AND(FALSE,1/0)', location)).toMatchObject({ kind: 'value', value: { type: 'error', value: '#DIV/0!' } });
    expect(evaluator.evaluate('=IFNA(1/0,3)', location)).toMatchObject({ kind: 'value', value: { type: 'error', value: '#DIV/0!' } }); evaluator.dispose();
    const presence = createWorkbookScalarEvaluator(scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 999 }, formula: '=FUTUREFUNC()' }]), { maxDepth: 1 });
    expect(presence.evaluate('=COUNTA(A:A)', location)).toMatchObject({ kind: 'value', value: { type: 'number', value: 1 }, readCells: 1 }); presence.dispose();
  });
  it('reads sparse whole axes without enumerating implicit blanks, preserving aggregate input types', () => {
    const model = scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 2 } }, { row: 2, column: 0, value: { type: 'text', value: '2' } },
      { row: 3, column: 0, value: { type: 'boolean', value: true } }, { row: 4, column: 0, value: { type: 'text', value: '' }, formula: '=""' },
      { row: 1_048_575, column: 0, value: { type: 'number', value: 5 } }]);
    const evaluator = createWorkbookScalarEvaluator(model, { maxCells: 10, maxSteps: 1000 });
    for (const [expression, expected] of [['=SUM(A:A)', 7], ['=COUNT(A:A)', 2], ['=COUNTA(A:A)', 5], ['=COUNTBLANK(A:A)', 1_048_572], ['=SUM("2",TRUE)', 3]] as const) {
      const result = evaluator.evaluate(expression, location); expect(result, expression).toMatchObject({ kind: 'value', value: { type: 'number', value: expected } }); expect(result.readCells).toBeLessThanOrEqual(5);
    }
    evaluator.dispose();
  });
  it('shifts relative templates, binds constant names and diagnoses actual reference and name cycles', () => {
    const model = scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 2 } }, { row: 1, column: 0, value: { type: 'number', value: 3 } },
      { row: 0, column: 1, value: { type: 'number', value: 999 }, formula: '=C1' }, { row: 0, column: 2, value: { type: 'number', value: 999 }, formula: '=B1' }], metadata => {
      metadata.workbook.names = [{ nameId: 'constant', name: 'Constant', expression: '=7', origin: { row: 0, column: 0 } },
        { nameId: 'cycle', name: 'Cycle', expression: '=Cycle', origin: { row: 0, column: 0 } }];
    });
    const evaluator = createWorkbookScalarEvaluator(model);
    expect(evaluator.evaluate('=A1+$A$2', { ...location, row: 2, column: 0 }, { at: { ...location, row: 3, column: 0 } })).toMatchObject({ kind: 'value', value: { type: 'number', value: 6 } });
    expect(evaluator.evaluate('=Constant', location)).toMatchObject({ kind: 'value', value: { type: 'number', value: 7 } });
    expect(evaluator.cell({ ...location, column: 1 })).toMatchObject({ kind: 'unavailable', reason: 'circular-reference' });
    expect(evaluator.evaluate('=Cycle', location)).toMatchObject({ kind: 'unavailable', reason: 'circular-name' }); evaluator.dispose();
  });
  it('fails bounded or cancelled queries without modifying the source or contaminating later queries', () => {
    const model = scalarWorkbook([{ row: 0, column: 0, value: { type: 'text', value: 'long text' } }, { row: 1, column: 0, value: { type: 'number', value: 2 } }]), before = model.serialize();
    for (const limits of [{ maxSteps: 1 }, { maxCodeUnits: 1 }, { maxDepth: 1 }, { maxCells: 1 }]) {
      const evaluator = createWorkbookScalarEvaluator(model, limits);
      expect(() => evaluator.evaluate('=SUM(A:A)+1', location)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' })); evaluator.dispose();
    }
    const evaluator = createWorkbookScalarEvaluator(model);
    expect(() => evaluator.evaluate('=1', location, { signal: { aborted: true } })).toThrowError(expect.objectContaining({ code: 'OperationCancelled' }));
    expect(evaluator.evaluate('=1', location)).toMatchObject({ kind: 'value', value: { type: 'number', value: 1 }, readCells: 0 });
    expect(() => evaluator.evaluate('=1+2+3', location, { maxSteps: 2 })).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    expect(model.serialize()).toEqual(before); evaluator.dispose();
  });
});
