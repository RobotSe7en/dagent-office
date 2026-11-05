import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { expect, it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import { importXlsx } from '@dagent-office/office';

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
