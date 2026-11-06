import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { expect, it } from 'vitest';
import { openDocument } from '@doffi/core';
import { importXlsx } from '@doffi/office';

it('imports and reopens every independent cell-flow source without flattening native alignment or underlines', async () => {
  const bytes = readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), evidence = JSON.parse(readFileSync('docs/evidence/xlsx-cell-flow.json', 'utf8'));
  const regenerated = execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('flow','scripts/check-xlsx-cell-flow.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.cases()))"]);
  expect(regenerated).toEqual(bytes); expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
  const model = openDocument(await importXlsx(bytes, { documentId: 'native-flow' }), { kind: 'workbook' }), snapshot = model.snapshot();
  expect(snapshot.metadata.workbook.sheets).toHaveLength(51);
  const css = snapshot.styles.map(style => style.css).join('\n');
  for (const [index, input] of evidence.cases.entries()) {
    const cell = model.getCell('sheet-' + (index + 1), 1, 1)!;
    expect(cell.value).toEqual({ type: typeof input.value === 'number' ? 'number' : 'text', value: input.value });
    const declarations = css.match(new RegExp('\\.' + cell.cssClass + '\\{([^}]*)\\}'))![1]!;
    expect(declarations).toContain('--office-cell-horizontal:' + input.horizontal);
    if (input.underline) expect(declarations).toContain('--office-font-underline:' + input.underline);
  }
  const reopened = openDocument(model.serialize(), { kind: 'workbook' }); expect(reopened.snapshot()).toEqual(snapshot);
});

it('imports independent fixed-default, automatic-exception and far-row sources with explicit sizing semantics', async () => {
  const bytes = readFileSync('packages/office/tests/fixtures/xlsx-row-heights.xlsx'), evidence = JSON.parse(readFileSync('docs/evidence/xlsx-row-heights.json', 'utf8'));
  const regenerated = execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('rows','scripts/check-xlsx-row-heights.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.cases()))"]);
  expect(regenerated).toEqual(bytes); expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
  const model = openDocument(await importXlsx(bytes, { documentId: 'native-auto-rows' }), { kind: 'workbook' }), snapshot = model.snapshot(), sheets = snapshot.metadata.workbook.sheets;
  expect(sheets).toHaveLength(21); expect(sheets[1]).toMatchObject({ defaultRowHeightCustom: true, defaultRowHeightPt: 18, rowDimensions: [] });
  expect(sheets[2]).toMatchObject({ defaultRowHeightCustom: true, rowDimensions: [{ from: 1, to: 1, autoHeight: true }] });
  expect(sheets[3]?.rowDimensions).toEqual([{ from: 1, to: 1, heightPt: 12 }]); expect(sheets[4]?.rowDimensions).toEqual([{ from: 1, to: 1, hidden: true }]);
  expect(model.getCell('sheet-15', 1, 1)?.value).toEqual({ type: 'blank' });
  expect(sheets[15]?.rowDimensions[0]?.cssClass).toBeDefined(); expect(sheets[16]?.columnDimensions[0]?.cssClass).toBeDefined();
  expect(model.getCell('sheet-14', 299, 1)?.value).toEqual({ type: 'text', value: evidence.cases[13].value });
  expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
});
