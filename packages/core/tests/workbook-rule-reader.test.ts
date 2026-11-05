import { describe, expect, it } from 'vitest';
import { DEFAULT_WORKBOOK_RULE_LIMITS, readWorkbookRules } from '../src/index.js';
import type { WorkbookConditionalRule, WorkbookCoordinate, WorkbookDataValidation, WorkbookModel, WorkbookRuleReadLimits } from '../src/index.js';
import { scalarWorkbook } from './workbook-scalar-fixture.js';

const origin = { row: 0, column: 0 }, ranges = [{ from: origin, to: { row: 2, column: 0 } }];
const condition = (ruleId: string, priority: number, expression: string, stopIfTrue = false): WorkbookConditionalRule => ({ ruleId, priority, expression, stopIfTrue, kind: 'expression', origin, ranges });
const validation = { ruleId: 'validation', origin, ranges, allowBlank: true, showErrorMessage: true, showInputMessage: true,
  errorStyle: 'stop' as const, errorTitle: '', error: '', promptTitle: '', prompt: '' };
function request(model: WorkbookModel, coordinates: readonly WorkbookCoordinate[], limits: Partial<WorkbookRuleReadLimits> = {}) {
  return { documentId: model.snapshot().manifest.documentId, revision: 0, sheetId: 'sheet-1', coordinates, ...DEFAULT_WORKBOOK_RULE_LIMITS, ...limits };
}
describe('bounded current-source workbook rule reads', () => {
  it('uses shifted current values, ordered priorities and uncertain stopping without changing caches', async () => {
    const model = scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 999 }, formula: '=2' },
      { row: 1, column: 0, value: { type: 'number', value: 999 }, formula: '=FUTUREFUNC()' },
      { row: 2, column: 0, value: { type: 'number', value: 8 } }], metadata => {
      metadata.workbook.sheets[0]!.conditionalRules = [condition('last', 3, '=TRUE'), condition('first', 1, '=A1=2'), condition('stop', 2, '=A1>5', true)];
    });
    const before = model.serialize(), result = await readWorkbookRules(model, request(model, [origin, { row: 1, column: 0 }, { row: 2, column: 0 }]));
    expect(result.cells[0]).toMatchObject({ current: { kind: 'value', value: { type: 'number', value: 2 } }, conditional: [
      { ruleId: 'first', state: 'match' }, { ruleId: 'stop', state: 'miss' }, { ruleId: 'last', state: 'match' }] });
    expect(result.cells[1]!.conditional).toEqual([{ ruleId: 'first', state: 'unavailable', issue: { reason: 'unsupported-function', functionName: 'FUTUREFUNC' } },
      { ruleId: 'stop', state: 'unavailable', issue: { reason: 'unsupported-function', functionName: 'FUTUREFUNC' } }]);
    expect(result.cells[2]!.conditional.map(rule => rule.state)).toEqual(['miss', 'match']);
    expect(model.serialize()).toEqual(before); expect(Object.isFrozen(result.cells[0]!.conditional)).toBe(true);
  });
  it('retains actual blanks for allowBlank and evaluates custom and numeric validation from current source', async () => {
    for (const [kind, expected] of [['custom', ['valid', 'valid', 'invalid']], ['whole', ['valid', 'valid', 'invalid']]] as const) {
      const model = scalarWorkbook([{ row: 1, column: 0, value: { type: 'number', value: 999 }, formula: '=2' }, { row: 2, column: 0, value: { type: 'number', value: 2.5 } }], metadata => {
        metadata.workbook.sheets[0]!.dataValidations = [{ ...validation, kind, ...(kind === 'custom' ? { expression: '=AND(A1>=1,A1<=2)' } : { operator: 'between', formulas: ['=1', '=2'] }) } as WorkbookDataValidation];
      });
      const result = await readWorkbookRules(model, request(model, [origin, { row: 1, column: 0 }, { row: 2, column: 0 }]));
      expect(result.cells.map(cell => cell.validation?.state)).toEqual(expected);
    }
  });
  it('reads cross-sheet lists and diagnoses unknown members rather than trusting their caches', async () => {
    const model = scalarWorkbook([{ row: 0, column: 0, value: { type: 'text', value: 'yes' } }, { row: 1, column: 0, value: { type: 'text', value: 'no' } },
      { row: 0, column: 1, value: { type: 'text', value: 'YES' } }, { row: 1, column: 1, value: { type: 'text', value: 'no' }, formula: '=FUTUREFUNC()' }], metadata => {
      const sheet = metadata.workbook.sheets[0]!;
      metadata.workbook.sheets.push({ ...sheet, sheetId: 'second', name: 'Second', dataValidations: [{ ...validation, kind: 'list', expression: '=Sheet1!$B$1:$B$2', dropdown: true }] });
    });
    // Select empty cells on Second: unknown candidate remains explicit. A list
    // referencing Sheet1 must never read equal coordinates on Second instead.
    const input = { ...request(model, [origin]), sheetId: 'second' };
    const response = await readWorkbookRules(model, input);
    expect(response.cells[0]!.validation?.state).toBe('valid'); // allowBlank
    const files = new Map(model.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!));
    metadata.workbook.sheets[1].dataValidations[0].allowBlank = false;
    files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
    const { openDocument } = await import('../src/index.js'), strict = openDocument(files, { kind: 'workbook' });
    expect((await readWorkbookRules(strict, input)).cells[0]!.validation).toMatchObject({ state: 'unavailable', issue: { reason: 'unsupported-function' } });
  });
  it('budgets entire batches, checks cancellation between host yields and captures mutable requests', async () => {
    const model = scalarWorkbook([], metadata => { metadata.workbook.sheets[0]!.conditionalRules = [condition('rule', 1, '=TRUE')]; });
    const points = [origin, { row: 1, column: 0 }, { row: 2, column: 0 }];
    for (const limits of [{ maxSteps: 1 }, { maxCells: 1 }, { maxCodeUnits: 1 }])
      await expect(readWorkbookRules(model, request(model, points, limits))).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    await expect(readWorkbookRules(model, request(model, points), { signal: { aborted: true } })).rejects.toMatchObject({ code: 'OperationCancelled' });
    const many = scalarWorkbook([], metadata => { metadata.workbook.sheets[0]!.conditionalRules = [{ ...condition('rule', 1, '=TRUE'), ranges: [{ from: origin, to: { row: 99, column: 0 } }] }]; });
    const input = request(many, Array.from({ length: 100 }, (_, row) => ({ row, column: 0 }))), signal = { aborted: false };
    await expect(readWorkbookRules(many, input, { signal, yield: async () => { signal.aborted = true; } })).rejects.toMatchObject({ code: 'OperationCancelled' });
    signal.aborted = false;
    const result = await readWorkbookRules(many, input, { yield: async () => { input.sheetId = 'mutated'; input.maxSteps = 1; input.coordinates[0]!.row = 1000; } });
    expect(result.sheetId).toBe('sheet-1'); expect(result.cells).toHaveLength(100); expect(result.cells[0]!.row).toBe(0);
  });
});
