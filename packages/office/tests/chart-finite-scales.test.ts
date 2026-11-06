import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@doffi/core';
import { importXlsx } from '@doffi/office';

const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-chart-finite-scales.json', 'utf8')) as { cases: { source: { axis: Record<string, string>; majorUnit: string | null; values: string[] }; producer: { axis: Record<string, string>; majorUnit: string | null; values: string[] } }[] };
describe('independent finite chart intervals', () => {
  it.each(['', '-producer'])('preserves%s finite bounds and subnormal current cells through package reopen', async suffix => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-chart-finite-scales' + suffix + '.xlsx');
    if (!suffix) expect(execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('finite','scripts/check-xlsx-chart-finite-scales.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"])).toEqual(bytes);
    const model = openDocument(await importXlsx(bytes, { documentId: 'finite-charts' }), { kind: 'workbook' }), snapshot = model.snapshot();
    expect(snapshot.metadata.workbook.charts).toHaveLength(14);
    for (const [index, chart] of snapshot.metadata.workbook.charts!.entries()) {
      const expected = evidence.cases[index]![suffix ? 'producer' : 'source'];
      expect(chart.minimum).toBe(expected.axis.min === undefined ? undefined : Number(expected.axis.min));
      expect(chart.maximum).toBe(expected.axis.max === undefined ? undefined : Number(expected.axis.max));
      expect(chart.majorUnit).toBe(expected.majorUnit === null ? undefined : Number(expected.majorUnit));
      expect([39, 40, 41].map(row => model.getCell(snapshot.metadata.workbook.sheets[index]!.sheetId, row, 1)!.value)).toEqual(expected.values.map(value => ({ type: 'number', value: Number(value) })));
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
});
