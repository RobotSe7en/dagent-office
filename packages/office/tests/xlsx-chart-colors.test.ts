import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import { importXlsx } from '../src/xlsx-import.js';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-chart-colors.json', 'utf8')) as {
  workbooks: { theme: string; sourceSha256: string; cases: { kind: string; vary: boolean; series: number; fill?: string }[] }[];
};
describe('native chart automatic colors and explicit overrides', () => {
  it.each(evidence.workbooks)('preserves independent $theme theme, variation flags and source styles on reopen', async workbook => {
    const bytes = readFileSync(`packages/office/tests/fixtures/xlsx-chart-colors-${workbook.theme}.xlsx`);
    const regenerated = execFileSync('python3', ['-c', `import importlib.util,sys;s=importlib.util.spec_from_file_location('chart_colors','scripts/check-xlsx-chart-colors.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs(),${workbook.theme === 'alternate' ? 'True' : 'False'}))`]);
    expect(regenerated).toEqual(bytes); expect(createHash('sha256').update(bytes).digest('hex')).toBe(workbook.sourceSha256);
    const model = openDocument(await importXlsx(bytes, { documentId: 'colors-' + workbook.theme }), { kind: 'workbook' }), snapshot = model.snapshot();
    const charts = snapshot.metadata.workbook.charts!; expect(charts).toHaveLength(workbook.cases.length);
    for (const [index, chart] of charts.entries()) {
      const source = workbook.cases[index]!; expect(chart).toMatchObject({ kind: source.kind, varyColors: source.vary }); expect(chart.series).toHaveLength(source.series);
      if (source.kind === 'line') for (const series of chart.seriesStyles!) expect(series).toMatchObject({ marker: 'circle', markerSize: 7 });
      if (source.fill === 'point') expect(chart.pointStyles).toHaveLength(1);
    }
    const css = snapshot.styles.map(item => item.css).join('\n'); expect(css).toContain('#112233'); expect(css).toContain('#123456');
    if (workbook.theme === 'alternate') expect(css).toContain('--office-theme-accent1:#aa5500');
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('distinguishes absent variation from a present flag with omitted true value and rejects invalid values', async () => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-chart-colors-default.xlsx'), path = 'xl/charts/chart1.xml';
    for (const [replacement, expected] of [['', undefined], ['<c:varyColors/>', true]] as const) {
      const files = new Map(readDocumentZip(bytes)), xml = new TextDecoder().decode(files.get(path)!);
      files.set(path, new TextEncoder().encode(xml.replace(/<c:varyColors[^>]*\/>/, replacement)));
      const model = openDocument(await importXlsx(writeOpcZip(files), { documentId: 'variation-flag' }), { kind: 'workbook' });
      expect(model.snapshot().metadata.workbook.charts![0]!.varyColors).toBe(expected);
    }
    const files = new Map(readDocumentZip(bytes)); files.set(path, new TextEncoder().encode(new TextDecoder().decode(files.get(path)!).replace(/<c:varyColors[^>]*\/>/, '<c:varyColors val="2"/>')));
    await expect(importXlsx(writeOpcZip(files), { documentId: 'invalid-variation' })).rejects.toMatchObject({ code: 'InvalidDocument' });
  });
});
