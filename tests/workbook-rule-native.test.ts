import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { expect, it } from 'vitest';
import { DEFAULT_WORKBOOK_RULE_LIMITS, openDocument, readWorkbookRules } from '@doffi/core';
import { importXlsx } from '@doffi/office';

it('reads the seven independent native validation kinds and four conditional rule kinds', async () => {
  const model = openDocument(await importXlsx(readFileSync('examples/xlsx/06-rules-validation.xlsx'), { documentId: 'rules' }), { kind: 'workbook' }), before = model.serialize();
  const coordinates = Array.from({ length: 5 }, (_, row) => Array.from({ length: 4 }, (_, column) => ({ row: row + 1, column }))).flat();
  coordinates.push(...Array.from({ length: 7 }, (_, row) => ({ row: row + 1, column: 4 })));
  const result = await readWorkbookRules(model, { documentId: 'rules', revision: 0, sheetId: 'sheet-1', coordinates, ...DEFAULT_WORKBOOK_RULE_LIMITS });
  const matches = (column: number) => result.cells.filter(cell => cell.column === column).map(cell => cell.conditional[0]!.state);
  expect(matches(0)).toEqual(['miss', 'miss', 'miss', 'match', 'match']);
  expect(matches(1)).toEqual(['miss', 'match', 'match', 'match', 'miss']);
  expect(result.cells.filter(cell => cell.column === 4).map(cell => cell.validation?.state)).toEqual(['valid', 'invalid', 'valid', 'valid', 'valid', 'valid', 'valid']);
  const expected = [
    { fromStop: 0, toStop: 0, ratio: 0 }, { fromStop: 0, toStop: 1, ratio: 40 / 190 }, { fromStop: 0, toStop: 1, ratio: 90 / 190 },
    { fromStop: 0, toStop: 1, ratio: 140 / 190 }, { fromStop: 1, toStop: 1, ratio: 0 }];
  for (const [index, cell] of result.cells.filter(cell => cell.column === 2).entries()) {
    expect(cell.conditional[0]!.scale).toMatchObject({ fromStop: expected[index]!.fromStop, toStop: expected[index]!.toStop });
    expect(cell.conditional[0]!.scale!.ratio).toBeCloseTo(expected[index]!.ratio, 14);
  }
  expect(model.serialize()).toEqual(before);
});
it('preserves native constant and repeated threshold behavior observed independently in Calc PDFs', async () => {
  const bytes = readFileSync('packages/office/tests/fixtures/xlsx-color-scales.xlsx'), evidence = JSON.parse(readFileSync('docs/evidence/xlsx-color-scale-semantics.json', 'utf8'));
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
  const model = openDocument(await importXlsx(bytes, { documentId: 'scales' }), { kind: 'workbook' });
  const expected = [[0, 0, 0], [0, 0, 0], [0, [0, 1, 0.5], 1], [2, 2, 2], [2, 2, 2], [0, [1, 2, 0], 2], [0, 0, 2], [0, 2, 2]];
  for (const [index, sheet] of model.snapshot().metadata.workbook.sheets.entries()) {
    const result = await readWorkbookRules(model, { documentId: 'scales', revision: 0, sheetId: sheet.sheetId, coordinates: [0, 1, 2].map(row => ({ row, column: 0 })), ...DEFAULT_WORKBOOK_RULE_LIMITS });
    for (const [row, cell] of result.cells.entries()) {
      const scale = cell.conditional[0]!.scale!, target = expected[index]![row]!;
      expect(cell.conditional[0]!.state, evidence.cases[index].name).toBe('match');
      expect(scale).toEqual(typeof target === 'number' ? { fromStop: target, toStop: target, ratio: 0 } : { fromStop: target[0], toStop: target[1], ratio: target[2] });
    }
  }
});
