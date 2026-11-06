import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@doffi/core';
import { importXlsx, importPptx, exportPptx } from '@doffi/office';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-chart-grid.json', 'utf8')) as { cases: { between: string; reverse: boolean; delete?: boolean }[] };
describe('category chart grids and crossing', () => {
  it.each(['', '-producer'])('preserves independent%s category grids, reversals and axis deletion on package reopen', async suffix => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-chart-grid' + suffix + '.xlsx');
    if (!suffix) expect(execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('grid','scripts/check-xlsx-chart-grid.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"])).toEqual(bytes);
    const model = openDocument(await importXlsx(bytes, { documentId: 'chart-grid' }), { kind: 'workbook' }), snapshot = model.snapshot();
    expect(snapshot.metadata.workbook.charts).toHaveLength(12);
    for (const [index, chart] of snapshot.metadata.workbook.charts!.entries()) {
      expect(chart.categoryGridCssClass).toBeDefined(); expect(chart.reverseCategories).toBe(evidence.cases[index]!.reverse);
      expect(chart.showCategoryAxis).toBe(!evidence.cases[index]!.delete);
      const css = snapshot.styles.map(style => style.css).join('\n');
      expect(css).toContain('.' + chart.categoryGridCssClass + '{--office-line-kind:solid;--office-line-color:#00ff00');
      expect(css).toContain('--office-axis-cross-between:' + evidence.cases[index]!.between);
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('preserves PPT category grids through the shared native writer', () => {
    const files = new Map(readDocumentZip(readFileSync('examples/pptx/09-charts.pptx'))), path = 'ppt/charts/chart2.xml';
    const source = new TextDecoder().decode(files.get(path)!), grid = '<c:majorGridlines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>';
    files.set(path, new TextEncoder().encode(source.replace(/(<c:catAx>[\s\S]*?<c:axPos[^>]*\/>)/, '$1' + grid)));
    const model = openDocument(importPptx(writeOpcZip(files), { documentId: 'ppt-grid' }), { kind: 'presentation' });
    const chart = model.snapshot().metadata.presentation!.charts.find(chart => chart.kind === 'line')!; expect(chart.categoryGridCssClass).toBeDefined();
    const target = openDocument(importPptx(exportPptx(model).bytes, { documentId: 'ppt-grid-after' }), { kind: 'presentation' });
    const after = target.snapshot().metadata.presentation!.charts.find(chart => chart.kind === 'line')!; expect(after.categoryGridCssClass).toBeDefined();
    expect(target.snapshot().styles.map(style => style.css).join('\n')).toContain('.' + after.categoryGridCssClass + '{--office-line-kind:solid;--office-line-color:#00ff00');
  });
});
