import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import { exportPptx, importPptx, importXlsx } from '@dagent-office/office';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-chart-markers.json', 'utf8')) as {
  cases: { marker?: string; size?: number }[];
  producerMarkers: { series: { symbol: string; size: string }; points: { index: string; marker: { symbol: string | null; size: string | null } | null }[] }[][];
};
const source = readFileSync('packages/office/tests/fixtures/xlsx-chart-markers.xlsx');
describe('native chart markers', () => {
  it('preserves independent point symbols, sizes, partial inheritance and paint through the HTML package', async () => {
    const regenerated = execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('markers','scripts/check-xlsx-chart-markers.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"]);
    expect(regenerated).toEqual(source);
    const model = openDocument(await importXlsx(source, { documentId: 'markers' }), { kind: 'workbook' }), snapshot = model.snapshot();
    for (const [index, chart] of snapshot.metadata.workbook.charts!.entries()) {
      expect(chart.seriesStyles![0]).toMatchObject({ marker: 'circle', markerSize: 10 });
      const point = chart.pointStyles![0]!, item = evidence.cases[index]!;
      expect(point.index).toBe(1); expect(point.marker).toBe(item.marker); expect(point.markerSize).toBe(item.size);
      expect(chart.pointStyles![1]).toMatchObject({ index: 2, marker: 'none' });
    }
    const css = snapshot.styles.map(style => style.css).join('\n');
    expect(css).toContain('--office-marker-fill-kind:linear'); expect(css).toContain('--office-marker-fill-kind:none');
    expect(css).toContain('--office-marker-line-width:0pt'); expect(css).toContain('--office-marker-line-width:1.5pt');
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('opens independent producer point markers, preserving the producer source without replacing it with the original', async () => {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-markers-producer.xlsx'), { documentId: 'producer-markers' }), { kind: 'workbook' }), snapshot = model.snapshot();
    for (const [index, chart] of snapshot.metadata.workbook.charts!.entries()) {
      const expected = evidence.producerMarkers[index]![0]!;
      expect(chart.seriesStyles![0]).toMatchObject({ marker: expected.series.symbol, markerSize: Number(expected.series.size) });
      for (const entry of expected.points) if (entry.marker) {
        const point = chart.pointStyles!.find(point => point.index === Number(entry.index))!;
        expect(point.marker).toBe(entry.marker.symbol ?? undefined); expect(point.markerSize).toBe(entry.marker.size === null ? undefined : Number(entry.marker.size));
      }
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it.each(['default', 'alternate'])('opens all 13 %s producer color charts, including varied per-point appearance', async theme => {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-colors-' + theme + '-producer.xlsx'), { documentId: 'producer-' + theme }), { kind: 'workbook' });
    const snapshot = model.snapshot(); expect(snapshot.metadata.workbook.charts).toHaveLength(13);
    expect(snapshot.metadata.workbook.charts!.filter(chart => chart.kind === 'line').some(chart => chart.pointStyles?.some(point => point.index === 1))).toBe(true);
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('rejects invalid native marker bounds, identities and unconsumed properties', async () => {
    for (const marker of ['<c:symbol val="picture"/>', '<c:symbol val="circle"/><c:symbol val="square"/>', '<c:symbol val="circle" unknown="1"/>', '<c:size val="1"/>', '<c:size val="73"/>', '<c:size val="2.5"/>', '<c:size val="NaN"/>', '<c:size val="20" unknown="1"/>', '<c:unknown/>']) {
      const files = new Map(readDocumentZip(source)), path = 'xl/charts/chart1.xml', xml = new TextDecoder().decode(files.get(path)!);
      files.set(path, new TextEncoder().encode(xml.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/, '<c:dPt><c:idx val="1"/><c:marker>' + marker + '</c:marker></c:dPt>')));
      await expect(importXlsx(writeOpcZip(files), { documentId: 'invalid-marker' })).rejects.toBeDefined();
    }
  });
  it('keeps PPT series and point marker appearance editable through native export and reopen', () => {
    const files = new Map(readDocumentZip(readFileSync('examples/pptx/09-charts.pptx'))), path = 'ppt/charts/chart2.xml';
    const xml = new TextDecoder().decode(files.get(path)!);
    const marker = '<c:marker><c:symbol val="star"/><c:size val="18"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr></c:marker>';
    files.set(path, new TextEncoder().encode(xml.replace(/<c:marker>[\s\S]*?<\/c:marker>/, marker).replace('<c:cat>', '<c:dPt><c:idx val="1"/>' + marker.replace('val="star"', 'val="dot"') + '</c:dPt><c:cat>')));
    const model = openDocument(importPptx(writeOpcZip(files), { documentId: 'ppt-markers' }), { kind: 'presentation' });
    const chart = model.snapshot().metadata.presentation!.charts.find(chart => chart.kind === 'line')!;
    expect(chart.seriesStyles![0]).toMatchObject({ marker: 'star', markerSize: 18 }); expect(chart.pointStyles![0]).toMatchObject({ marker: 'dot', markerSize: 18 });
    const output = exportPptx(model), native = new Map(readDocumentZip(output.bytes));
    const exported = new TextDecoder().decode(native.get(path)!); expect(exported).toContain('<c:symbol val="dot"'); expect(exported).toContain('val="ff0000"'); expect(exported).toContain('val="0000ff"');
    const after = openDocument(importPptx(output.bytes, { documentId: 'ppt-markers-after' }), { kind: 'presentation' });
    const reopened = after.snapshot().metadata.presentation!.charts.find(chart => chart.kind === 'line')!;
    expect(reopened.seriesStyles![0]).toMatchObject({ marker: 'star', markerSize: 18 }); expect(reopened.pointStyles![0]).toMatchObject({ marker: 'dot', markerSize: 18 });
    expect(after.snapshot().styles.map(style => style.css).join('\n')).toContain('--office-marker-line-color:#0000ff');
  });
});
