import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@doffi/core';
import { exportPptx, importPptx, importXlsx } from '@doffi/office';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

const report = JSON.parse(readFileSync('docs/evidence/xlsx-chart-auto-markers.json', 'utf8')) as {
  cases: { series: number; symbol: string; pointSymbol?: string; pointSize?: number; pointPaint?: string }[];
  producerMarkers: { series: { symbol: string | null; size: string | null } | null; points: { index: string; marker: { symbol: string | null; size: string | null } | null }[] }[][];
  producerLegacyFlags: { val?: string }[][];
};
function legacyFlag(xml: string, flag: string, removeSeries = false): string {
  const boundary = xml.lastIndexOf('</c:ser>') + '</c:ser>'.length;
  let head = xml.slice(0, boundary), tail = xml.slice(boundary).replace(/<c:marker(?:\s[^>]*)?\/>/g, '');
  if (removeSeries) head = head.replace(/<c:marker>[\s\S]*?<\/c:marker>/g, '');
  tail = tail.replace(/<c:(?:smooth|axId)\b/, match => flag + match);
  return head + tail;
}
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
  it('reopens the lossy producer source while preserving and ignoring its legacy marker flag', async () => {
    expect(report.producerMarkers[0]![0]!.series).toBeNull();
    expect(report.producerLegacyFlags).toEqual(report.cases.map(() => [{ val: '1' }]));
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-auto-markers-producer.xlsx'), { documentId: 'producer-auto' }), { kind: 'workbook' }), snapshot = model.snapshot();
    for (const [index, chart] of snapshot.metadata.workbook.charts!.entries()) {
      expect(chart.nativeShowMarkers).toBe(true);
      for (const [series, style] of chart.seriesStyles!.entries()) {
        const expected = report.producerMarkers[index]![series]!;
        expect(style.marker).toBe(expected.series?.symbol ?? 'none'); expect(style.markerSize).toBe(Number(expected.series?.size ?? 5));
        for (const point of expected.points) {
          const actual = chart.pointStyles!.find(item => item.seriesId === style.seriesId && item.index === Number(point.index))!;
          expect(actual.marker).toBe(point.marker?.symbol ?? undefined); expect(actual.markerSize).toBe(point.marker?.size == null ? undefined : Number(point.marker.size));
        }
      }
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('distinguishes absent, empty, false and true native flags without inventing implicit series markers', async () => {
    for (const [flag, expected] of [['', undefined], ['<c:marker/>', true], ['<c:marker val="0"/>', false], ['<c:marker val="false"/>', false], ['<c:marker val="1"/>', true], ['<c:marker val="true"/>', true], ['<c:marker val=" true "/>', true], ['<c:marker val=" 0 "/>', false]] as const) {
      const files = new Map(readDocumentZip(readFileSync('packages/office/tests/fixtures/xlsx-chart-auto-markers.xlsx'))), path = 'xl/charts/chart1.xml';
      files.set(path, new TextEncoder().encode(legacyFlag(new TextDecoder().decode(files.get(path)!), flag, true)));
      const model = openDocument(await importXlsx(writeOpcZip(files), { documentId: 'native-flag' }), { kind: 'workbook' });
      const chart = model.snapshot().metadata.workbook.charts![0]!;
      expect(chart.nativeShowMarkers).toBe(expected); expect(Object.hasOwn(chart, 'nativeShowMarkers')).toBe(expected !== undefined);
      expect(chart.seriesStyles!.every(series => series.marker === 'none')).toBe(true);
      expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(model.snapshot());
    }
  });
  it('rejects malformed, duplicated and unconsumed legacy flags in both native readers', async () => {
    for (const [flag, code] of [['<c:marker val="yes"/>', 'InvalidDocument'], ['<c:marker val="TRUE"/>', 'InvalidDocument'], ['<c:marker val="1" unexpected="1"/>', 'UnsupportedFeature'], ['<c:marker>unexpected</c:marker>', 'InvalidDocument'], ['<c:marker> </c:marker>', 'InvalidDocument'], ['<c:marker val="1"/><c:marker val="0"/>', 'InvalidDocument']] as const) {
      for (const kind of ['xlsx', 'pptx']) {
        const files = new Map(readDocumentZip(readFileSync(kind === 'xlsx' ? 'packages/office/tests/fixtures/xlsx-chart-auto-markers.xlsx' : 'examples/pptx/09-charts.pptx'))), path = kind === 'xlsx' ? 'xl/charts/chart1.xml' : 'ppt/charts/chart2.xml';
        files.set(path, new TextEncoder().encode(legacyFlag(new TextDecoder().decode(files.get(path)!), flag)));
        const read = async () => kind === 'xlsx' ? importXlsx(writeOpcZip(files), { documentId: 'bad-native-flag' }) : importPptx(writeOpcZip(files), { documentId: 'bad-native-flag' });
        await expect(read()).rejects.toMatchObject({ code });
      }
    }
  });
  it.each([['', undefined], ['<c:marker/>', true], ['<c:marker val="0"/>', false], ['<c:marker val="1"/>', true], ['<c:marker val=" false "/>', false]] as const)('preserves PPT native flag %s through editable export', (flag, expected) => {
    const files = new Map(readDocumentZip(readFileSync('examples/pptx/09-charts.pptx'))), path = 'ppt/charts/chart2.xml';
    files.set(path, new TextEncoder().encode(legacyFlag(new TextDecoder().decode(files.get(path)!), flag)));
    const model = openDocument(importPptx(writeOpcZip(files), { documentId: 'ppt-native-flag' }), { kind: 'presentation' });
    expect(model.snapshot().metadata.presentation!.charts.find(chart => chart.kind === 'line')!.nativeShowMarkers).toBe(expected);
    const output = exportPptx(model), after = openDocument(importPptx(output.bytes, { documentId: 'ppt-flag-reopen' }), { kind: 'presentation' });
    expect(after.snapshot().metadata.presentation!.charts.find(chart => chart.kind === 'line')!.nativeShowMarkers).toBe(expected);
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
