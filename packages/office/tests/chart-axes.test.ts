import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@doffi/core';
import { importXlsx, importPptx, exportPptx } from '@doffi/office';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

type NativeAxis = { crosses?: string; crossesAt?: string; axPos?: string };
const report = JSON.parse(readFileSync('docs/evidence/xlsx-chart-axes.json', 'utf8')) as { cases: { categoryAt?: number; valueAt?: number; categoryCrosses?: string; valueCrosses?: string; categoryPosition?: string; valuePosition?: string; producer: { category: NativeAxis; value: NativeAxis } }[] };
function axisXml(xml: string, category: boolean, crossing: string): string {
  const tag = category ? 'catAx' : 'valAx';
  return xml.replace(new RegExp('<c:' + tag + '>[\\s\\S]*?</c:' + tag + '>'), axis => axis.replace(/<c:crosses(?:At)?\b[^>]*\/>/g, '').replace(/(<c:crossAx\b[^>]*\/>)/, '$1' + crossing));
}
describe('Office-profile native chart axes', () => {
  it.each(['', '-producer'])('preserves independent%s axis coordinates and source flags on package reopen', async suffix => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-chart-axes' + suffix + '.xlsx');
    if (!suffix) expect(execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('axes','scripts/check-xlsx-chart-axes.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"])).toEqual(bytes);
    const model = openDocument(await importXlsx(bytes, { documentId: 'native-axes' }), { kind: 'workbook' }), snapshot = model.snapshot();
    expect(snapshot.metadata.workbook.charts).toHaveLength(report.cases.length);
    for (const [index, chart] of snapshot.metadata.workbook.charts!.entries()) {
      const source = report.cases[index]!;
      const css = snapshot.styles.map(style => style.css).join('\n');
      for (const category of [true, false]) {
        const name = category ? chart.categoryAxisCssClass : chart.valueAxisCssClass;
        const rule = css.match(new RegExp('\\.' + name + '\\{([^}]+)\\}'))![1]!;
        const at = category ? source.categoryAt : source.valueAt;
        const native = suffix ? source.producer[category ? 'category' : 'value'] : { crossesAt: at === undefined ? undefined : String(at), crosses: at === undefined ? (category ? source.categoryCrosses : source.valueCrosses) ?? 'autoZero' : undefined, axPos: (category ? source.categoryPosition : source.valuePosition) ?? (category ? 'b' : 'l') };
        expect(rule).toContain('--office-axis-position:' + native.axPos);
        if (native.crossesAt !== undefined) { expect(rule).toContain('--office-axis-crosses-at:' + Number(native.crossesAt)); expect(rule).not.toMatch(/--office-axis-crosses:/); }
        else expect(rule).toContain('--office-axis-crosses:' + native.crosses);
      }
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('preserves decimal value crossings and integral category crossings through editable PPT export', () => {
    const files = new Map(readDocumentZip(readFileSync('examples/pptx/09-charts.pptx'))), path = 'ppt/charts/chart2.xml';
    let xml = new TextDecoder().decode(files.get(path)!);
    xml = axisXml(axisXml(xml, true, '<c:crossesAt val=" -1.25e0 "/>'), false, '<c:crossesAt val="2"/>');
    files.set(path, new TextEncoder().encode(xml));
    const model = openDocument(importPptx(writeOpcZip(files), { documentId: 'axis-ppt' }), { kind: 'presentation' });
    const bytes = exportPptx(model).bytes, native = new TextDecoder().decode(new Map(readDocumentZip(bytes)).get(path)!);
    expect(native).toContain('<c:crossesAt val="-1.25"'); expect(native).toContain('<c:crossesAt val="2"');
    expect(native.match(/<c:crosses\b/g)).toBeNull();
    const after = openDocument(importPptx(bytes, { documentId: 'axis-ppt-after' }), { kind: 'presentation' });
    expect(after.snapshot().styles.map(style => style.css).join('\n')).toContain('--office-axis-crosses-at:-1.25');
    expect(after.snapshot().styles.map(style => style.css).join('\n')).toContain('--office-axis-crosses-at:2');
  });
  it('rejects malformed, duplicate, conflicting and unconsumed crossing choices in both native readers', async () => {
    for (const [crossing, code] of [
      ['<c:crossesAt/>', 'InvalidDocument'], ['<c:crossesAt val="NaN"/>', 'InvalidDocument'], ['<c:crossesAt val="INF"/>', 'InvalidDocument'],
      ['<c:crossesAt val="0x10"/>', 'InvalidDocument'], ['<c:crossesAt val="1"/><c:crosses val="min"/>', 'InvalidDocument'],
      ['<c:crossesAt val="1"/><c:crossesAt val="2"/>', 'InvalidDocument'], ['<c:crossesAt val="1" unexpected="1"/>', 'UnsupportedFeature'],
      ['<c:crossesAt val="1"> </c:crossesAt>', 'InvalidDocument'], ['<c:crosses val="minimum"/>', 'InvalidDocument'],
    ] as const) for (const kind of ['xlsx', 'pptx']) {
      const files = new Map(readDocumentZip(readFileSync(kind === 'xlsx' ? 'packages/office/tests/fixtures/xlsx-chart-axes.xlsx' : 'examples/pptx/09-charts.pptx'))), path = kind === 'xlsx' ? 'xl/charts/chart1.xml' : 'ppt/charts/chart2.xml';
      files.set(path, new TextEncoder().encode(axisXml(new TextDecoder().decode(files.get(path)!), true, crossing)));
      const read = async () => kind === 'xlsx' ? importXlsx(writeOpcZip(files), { documentId: 'bad-crossing' }) : importPptx(writeOpcZip(files), { documentId: 'bad-crossing' });
      await expect(read()).rejects.toMatchObject({ code });
    }
  });
  it('does not infer a category rounding policy from Calc', async () => {
    const files = new Map(readDocumentZip(readFileSync('packages/office/tests/fixtures/xlsx-chart-axes.xlsx'))), path = 'xl/charts/chart1.xml';
    files.set(path, new TextEncoder().encode(axisXml(new TextDecoder().decode(files.get(path)!), false, '<c:crossesAt val="2.5"/>')));
    await expect(importXlsx(writeOpcZip(files), { documentId: 'fractional-category-crossing' })).rejects.toMatchObject({ code: 'UnsupportedFeature' });
  });
  it('rejects invalid persisted crossing CSS before emitting editable native XML', () => {
    const model = openDocument(importPptx(readFileSync('examples/pptx/09-charts.pptx'), { documentId: 'persisted-axis-css' }), { kind: 'presentation' });
    const snapshot = model.snapshot(), chart = snapshot.metadata.presentation!.charts.find(chart => chart.kind === 'line')!, path = snapshot.styles[0]!.path;
    for (const raw of ['NaN', '1e309', '0x10', '2.5']) {
      const files = new Map(model.serialize()), css = new TextDecoder().decode(files.get(path)!);
      files.set(path, new TextEncoder().encode(css + '\n.' + chart.valueAxisCssClass + '{--office-axis-crosses-at:' + raw + '}'));
      const changed = openDocument(files, { kind: 'presentation' });
      expect(() => exportPptx(changed)).toThrowError(expect.objectContaining({ code: 'UnsupportedFeature' }));
    }
  });
});
