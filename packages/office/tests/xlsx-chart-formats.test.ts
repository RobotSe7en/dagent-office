import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import { importXlsx } from '../src/xlsx-import.js';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

const path = 'packages/office/tests/fixtures/xlsx-chart-formats.xlsx';
const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-chart-format-semantics.json', 'utf8')) as { sourceSha256: string; cases: { role: string; linked: boolean | string }[] };
const declarations = (css: string, id: string | undefined) => [...css.matchAll(new RegExp('\\.' + id + '\\{([^}]*)\\}', 'g'))].map(match => match[1]!).join(';');

describe('native chart source-linked number formats', () => {
  it('preserves explicit, omitted and absent formats from independent sources across package reopen', async () => {
    const generated = execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('chart_formats','scripts/check-xlsx-chart-formats.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"]);
    expect(createHash('sha256').update(generated).digest('hex')).toBe(evidence.sourceSha256); expect(generated).toEqual(readFileSync(path));
    const model = openDocument(await importXlsx(generated, { documentId: 'chart-formats' }), { kind: 'workbook' }), snapshot = model.snapshot();
    const charts = snapshot.metadata.workbook.charts!, css = snapshot.styles.map(style => style.css).join('\n');
    expect(charts).toHaveLength(evidence.cases.length);
    for (const [index, entry] of evidence.cases.entries()) {
      const chart = charts[index]!, label = entry.role.startsWith('label'), id = label ? chart.labelCssClass : entry.role === 'category-axis' ? chart.categoryAxisCssClass : chart.valueAxisCssClass;
      const declaration = declarations(css, id), flag = `--office-${label ? 'label' : 'axis'}-number-format-source-linked`;
      if (entry.linked === 'no-format') expect(declaration).not.toContain(flag);
      else expect(declaration).toContain(flag + ':' + String(entry.linked !== false));
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('resets a point format with omitted linkage instead of inheriting its parent false', async () => {
    const files = new Map(readDocumentZip(readFileSync(path))), part = 'xl/charts/chart9.xml', source = new TextDecoder().decode(files.get(part)!);
    const changed = source.replace('<c:dLbls>', '<c:dLbls><c:dLbl><c:idx val="1"/><c:numFmt formatCode="0.000"/></c:dLbl>'); expect(changed).not.toBe(source);
    files.set(part, new TextEncoder().encode(changed));
    const snapshot = openDocument(await importXlsx(writeOpcZip(files), { documentId: 'point-format' }), { kind: 'workbook' }).snapshot();
    const chart = snapshot.metadata.workbook.charts![8]!, css = snapshot.styles.map(style => style.css).join('\n');
    expect(declarations(css, chart.labelCssClass)).toContain('--office-label-number-format-source-linked:false');
    expect(declarations(css, chart.pointStyles?.[0]?.labelCssClass)).toContain('--office-label-number-format-source-linked:true');
  });
});
