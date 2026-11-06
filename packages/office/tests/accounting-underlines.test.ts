import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@doffi/core';
import { importXlsx } from '@doffi/office';

const report = JSON.parse(readFileSync('docs/evidence/xlsx-accounting-underlines.json', 'utf8')) as { cases: { source: { underline: string }; producer: { underline: string } }[] };
describe('native accounting underline sources', () => {
  it.each(['', '-producer'])('preserves%s actual font types and automatic row sources on package reopen', async suffix => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-accounting-underlines' + suffix + '.xlsx');
    if (!suffix) expect(execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('accounting','scripts/check-xlsx-accounting-underlines.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"])).toEqual(bytes);
    const model = openDocument(await importXlsx(bytes, { documentId: 'accounting-source' }), { kind: 'workbook' }), snapshot = model.snapshot();
    expect(snapshot.metadata.workbook.sheets).toHaveLength(20);
    const css = snapshot.styles.map(style => style.css).join('\n');
    for (const [index, sheet] of snapshot.metadata.workbook.sheets.entries()) {
      const cell = model.getCell(sheet.sheetId, 1, 1)!;
      const expected = report.cases[index]![suffix ? 'producer' : 'source'].underline;
      const rules = (cell.cssClass?.split(' ') ?? []).flatMap(name => [...css.matchAll(new RegExp('\\.' + name + '\\{([^}]*)\\}', 'g'))].map(match => match[1]));
      expect(rules.join(';'), sheet.name).toContain('--office-font-underline:' + expected);
      expect(cell.value.type).toBe(index < 8 && index % 2 || index === 9 ? 'number' : 'text');
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('records Calc rewriting accounting types without inferring them back from ordinary underlines', () => {
    expect(report.cases[0]).toMatchObject({ source: { underline: 'singleAccounting' }, producer: { underline: 'single' } });
    expect(report.cases[4]).toMatchObject({ source: { underline: 'doubleAccounting' }, producer: { underline: 'double' } });
  });
});
