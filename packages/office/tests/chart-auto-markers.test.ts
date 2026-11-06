import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import { exportPptx, importPptx, importXlsx } from '@dagent-office/office';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

const report = JSON.parse(readFileSync('docs/evidence/xlsx-chart-auto-markers.json', 'utf8')) as {
  cases: { series: number; symbol: string; pointSymbol?: string; pointSize?: number; pointPaint?: string }[];
  producerMarkers: { series: { symbol: string | null; size: string | null } | null; points: { index: string; marker: { symbol: string | null; size: string | null } | null }[] }[][];
  producerLegacyFlags: { val?: string }[][];
};
describe('native explicit automatic chart markers', () => {
  it('preserves independent source auto and partial point properties through public import and package reopen', async () => {
    const source = readFileSync('packages/office/tests/fixtures/xlsx-chart-auto-markers.xlsx');
    const generated = execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('auto','scripts/check-xlsx-chart-auto-markers.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"]);
    expect(generated).toEqual(source);
    const model = openDocument(await importXlsx(source, { documentId: 'auto-markers' }), { kind: 'workbook' }), snapshot = model.snapshot();
    for (const [index, chart] of snapshot.metadata.workbook.charts!.entries()) {
      const expected = report.cases[index]!; expect(chart.series).toHaveLength(expected.series);
      for (const style of chart.seriesStyles!) expect(style).toMatchObject({ marker: expected.symbol, markerSize: 10 });
      for (const point of chart.pointStyles!) { expect(point.marker).toBe(expected.pointSymbol); expect(point.markerSize).toBe(expected.pointSize); }
    }
    expect(snapshot.styles.some(style => style.css.includes('--office-marker-fill-kind:none'))).toBe(true);
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('reports the producer legacy Boolean without silently selecting its pending product semantics', async () => {
    // Calc drops the first series marker but writes lineChart/marker=true.
    // Preserve the original producer bytes as evidence; do not normalize that
    // flag or restore source auto to manufacture a successful reopen.
    expect(report.producerMarkers[0]![0]!.series).toBeNull();
    expect(report.producerLegacyFlags).toEqual(report.cases.map(() => [{ val: '1' }]));
    await expect(importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-auto-markers-producer.xlsx'), { documentId: 'producer-auto' }))
      .rejects.toMatchObject({ code: 'UnsupportedFeature', message: expect.stringContaining('Implicit automatic marker shapes') });
  });
  it('retains explicit auto at series and point levels through native editable PPT export and reopen', () => {
    const files = new Map(readDocumentZip(readFileSync('examples/pptx/09-charts.pptx'))), path = 'ppt/charts/chart2.xml';
    const xml = new TextDecoder().decode(files.get(path)!).replace(/<c:symbol val="[^"]+"\/>/g, '<c:symbol val="auto"/>')
      .replace('<c:cat>', '<c:dPt><c:idx val="1"/><c:marker><c:symbol val="auto"/><c:size val="20"/></c:marker></c:dPt><c:cat>');
    files.set(path, new TextEncoder().encode(xml));
    const model = openDocument(importPptx(writeOpcZip(files), { documentId: 'ppt-auto' }), { kind: 'presentation' });
    const chart = model.snapshot().metadata.presentation!.charts.find(chart => chart.kind === 'line')!;
    expect(chart.seriesStyles!.every(series => series.marker === 'auto')).toBe(true); expect(chart.pointStyles![0]).toMatchObject({ marker: 'auto', markerSize: 20 });
    const output = exportPptx(model), native = new Map(readDocumentZip(output.bytes)); expect(new TextDecoder().decode(native.get(path)!)).toContain('<c:symbol val="auto"');
    const after = openDocument(importPptx(output.bytes, { documentId: 'ppt-auto-after' }), { kind: 'presentation' }).snapshot().metadata.presentation!.charts.find(chart => chart.kind === 'line')!;
    expect(after.seriesStyles!.every(series => series.marker === 'auto')).toBe(true); expect(after.pointStyles![0]).toMatchObject({ marker: 'auto', markerSize: 20 });
  });
});
