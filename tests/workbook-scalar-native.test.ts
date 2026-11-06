import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { createWorkbookScalarEvaluator, openDocument } from '@dagent-office/core';
import type { WorkbookValue } from '@dagent-office/core';
import { importXlsx } from '@dagent-office/office';

it('matches independent scalar expectations while recording actual producer differences and ignoring wrong native caches', async () => {
  const bytes = readFileSync('packages/office/tests/fixtures/xlsx-scalar-semantics.xlsx');
  const regenerated = execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('scalar','scripts/check-xlsx-scalar-semantics.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate())"]);
  expect(regenerated).toEqual(bytes);
  const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-scalar-semantics.json', 'utf8')) as { sourceSha256: string; cases: { address: string; expression: string; value: WorkbookValue; expected: WorkbookValue; matchesProducer: boolean; basis: string }[] };
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
  const model = openDocument(await importXlsx(bytes, { documentId: 'native-scalar' }), { kind: 'workbook' }), before = model.serialize(), evaluator = createWorkbookScalarEvaluator(model);
  try {
    for (const [index, input] of evidence.cases.entries()) {
      const result = evaluator.cell({ sheetId: 'sheet-1', row: index, column: 4 });
      expect(result, input.expression).toMatchObject({ kind: 'value', value: input.expected });
      if (input.matchesProducer) expect(input.expected).toEqual(input.value);
      else { expect(input.expected).not.toEqual(input.value); expect(input.basis).not.toContain('Pinned Calc observation'); }
    }
    expect(evidence.cases.filter(input => !input.matchesProducer)).toHaveLength(10);
    expect(model.getCell('sheet-1', 0, 4)!.value).toEqual({ type: 'number', value: 999 });
    expect(model.serialize()).toEqual(before);
  } finally { evaluator.dispose(); }
});
