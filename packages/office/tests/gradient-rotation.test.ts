import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import { importXlsx } from '@dagent-office/office';

const report = JSON.parse(readFileSync('docs/evidence/xlsx-gradient-rotation.json', 'utf8')) as { cases: { source: { gradient: Record<string, string>; linear: Record<string, string> }; producer: { gradient: Record<string, string>; linear: Record<string, string> } }[] };
describe('native gradient rotation sources', () => {
  it.each(['', '-producer'])('preserves%s actual rotation, angle and scaled properties without restoring Calc losses', async suffix => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-gradient-rotation' + suffix + '.xlsx');
    if (!suffix) expect(execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('gradients','scripts/check-xlsx-gradient-rotation.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"])).toEqual(bytes);
    const model = openDocument(await importXlsx(bytes, { documentId: 'gradient-rotation' }), { kind: 'workbook' }), snapshot = model.snapshot(), css = snapshot.styles.map(style => style.css).join('\n');
    expect(snapshot.metadata.workbook.objects).toHaveLength(32);
    for (const [index, sheet] of snapshot.metadata.workbook.sheets.entries()) {
      const object = snapshot.metadata.workbook.objects!.find(object => object.sheetId === sheet.sheetId)!;
      const native = report.cases[index]![suffix ? 'producer' : 'source'], classes = model.getNode(object.nodeId)!.attributes.class!.split(' ');
      const rules = classes.flatMap(name => [...css.matchAll(new RegExp('\\.' + name + '\\{([^}]+)\\}', 'g'))].map(match => match[1])).join(';');
      expect(rules).toContain('--office-fill-kind:linear'); expect(rules).toContain('--office-fill-angle:' + Number(native.linear.ang ?? 0)/60000);
      for (const [raw, property] of [[native.gradient.rotWithShape, '--office-fill-rotate-with-shape'], [native.linear.scaled, '--office-fill-scaled']] as const) {
        if (raw !== undefined) expect(rules).toContain(property + ':' + String(['1', 'true'].includes(raw)));
        else expect(rules).not.toContain(property + ':');
      }
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
});
