import type { WorkbookSnapshot } from '@dagent-office/core';
import type { BasicChartData } from '../../../packages/render/src/chart.js';
import { createWorkbookChartView } from '../../../packages/render/src/workbook/chart-view.js';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookChartText } from '../../../packages/render/src/workbook/drawing-text.js';
import type { WorkbookChartTextRole } from '../../../packages/render/src/workbook/drawing-text.js';
import { layoutWorkbookChartTexts } from '../../../packages/render/src/workbook/drawing-text-layout.js';

export async function measureWorkbookCharts(snapshot: WorkbookSnapshot, plots: readonly BasicChartData[]) {
  const container = document.createElement('div'); document.body.append(container);
  const css = '.chart-text-probe{font-family:"Liberation Sans";font-size:13pt;color:rgb(12,34,56);--office-text-color:#0c2238;--office-text-source:direct;--office-text-lum-mod:1;--office-text-lum-off:0;--office-text-tint:0;--office-text-shade:1;--office-text-alpha:1}svg{width:1px!important;height:1px!important}text{font-size:99pt!important;fill:red!important}';
  const surface = await createWorkbookSurface(container, { ...snapshot, styles: [...snapshot.styles, { path: 'chart-probe.css', css }] }, { widthPt: 600, heightPt: 300 });
  const sources = new WorkbookContentSources(surface.document, snapshot), results = [];
  try {
    const definitions = snapshot.metadata.workbook.charts!, projection = sources.project(surface.source, definitions.map(chart => chart.nodeId));
    for (const [index, definition] of definitions.entries()) {
      const chart = { ...definition, textCssClass: 'chart-text-probe', categoryAxisCssClass: 'chart-text-probe', valueAxisCssClass: 'chart-text-probe' };
      const source = projection.get(chart.nodeId)!, before = source.outerHTML, data = plots[index]!;
      const title = () => createWorkbookChartText(source, chart, 'title').element;
      // The test supplies an explicit numeric-only format policy. Workbook
      // locale/date policy and rich title projection are separate integration.
      const formatNumber = (value: number, format?: string) => format === '0%' ? `${value * 100}%` : String(Number(value.toPrecision(12)));
      const options = () => ({ formatNumber, title: title() });
      const result = createWorkbookChartView(source, chart, data, 360, 220, 'chart-' + index, options());
      surface.view.append(result.element);
      const walker = surface.document.createTreeWalker(result.element, 0xffff_ffff); let nodes = 0; do { nodes++; } while (walker.nextNode());
      const svg = result.element.querySelector('svg')!, text = [...svg.querySelectorAll('text')];
      const errors = [], unchanged = () => before === source.outerHTML;
      for (const kind of ['node-budget', 'point-budget', 'text-budget', 'formatter', 'invalid-data', 'source-title', 'missing-title', 'axis-scale', 'formatted-text-budget']) {
        try {
          const selected = kind === 'invalid-data' ? { ...data, series: data.series.map(series => ({ ...series, values: [NaN, ...series.values.slice(1)] })) } : data;
          const appearance = kind === 'formatter' ? { ...chart, showLabels: true, labelContents: ['value' as const] } : kind === 'axis-scale' ? { ...chart, kind: 'column' as const, minimum: 1e16, maximum: 1e16 + 4, majorUnit: .1, showValueAxis: false } : kind === 'formatted-text-budget' ? { ...chart, kind: 'column' as const } : chart;
          createWorkbookChartView(source, appearance, selected, 360, 220, 'failure-' + index, { formatNumber: kind === 'formatter' ? () => { throw Error('format rejected'); } : kind === 'formatted-text-budget' ? () => 'x'.repeat(10000) : formatNumber,
            ...(kind === 'missing-title' ? {} : { title: kind === 'source-title' ? source.querySelector('div')! : title() }),
            ...(kind === 'node-budget' ? { maxNodes: result.peakNodes - 1 } : {}), ...(kind === 'point-budget' ? { maxPoints: 1 } : {}), ...(kind === 'text-budget' ? { maxCodeUnits: 1 } : kind === 'formatted-text-budget' ? { maxCodeUnits: 1000 } : {}) });
          errors.push('accepted');
        } catch (error) { errors.push((error as { code?: string; message: string }).code ?? (error as Error).message); }
        if (!unchanged()) throw Error('Chart probe changed source after failure');
      }
      const exact = createWorkbookChartView(source, chart, data, 360, 220, 'exact-' + index, { ...options(), maxNodes: result.peakNodes });
      const titleText = createWorkbookChartText(source, chart, 'title'), richSource = source.querySelector<HTMLElement>('[data-node-id="' + chart.titleNodeId + '"]')!;
      const titleProjection = titleText.projections.map(item => ({ from: item.from, to: item.to, text: item.node.data, owned: richSource.querySelector('[data-node-id="' + item.sourceNodeId + '"]') !== null }));
      let negativePie = '';
      if (chart.kind === 'pie') {
        try { createWorkbookChartView(source, chart, { ...data, series: data.series.map(series => ({ ...series, values: [-2, ...series.values.slice(1)] })) }, 360, 220, 'negative-pie', options()); negativePie = 'accepted'; }
        catch (error) { negativePie = (error as { code: string }).code; }
      }
      results.push({ kind: chart.kind, nodes, reported: result.nodes, peak: result.peakNodes, exact: exact.nodes === nodes, unchanged: unchanged(), errors, negativePie, titleProjection, titleText: richSource.textContent,
        sourceClasses: result.element.querySelectorAll('[class],[data-node-id]').length,
        svgSize: [svg.clientWidth, svg.clientHeight], text: text.map(node => node.textContent),
        fonts: text.map(node => surface.document.defaultView!.getComputedStyle(node).fontSize), colors: text.map(node => surface.document.defaultView!.getComputedStyle(node).fill),
        bars: [...svg.querySelectorAll('rect')].filter(node => Number(node.getAttribute('width')) > 20 && Number(node.getAttribute('width')) < 100).map(node => Number(node.getAttribute('height'))) });
      result.element.remove();
    }
    return results;
  } finally { sources.dispose(); surface.dispose(); container.remove(); }
}

export async function measureWorkbookChartTitles(snapshot: WorkbookSnapshot) {
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 600, heightPt: 300 });
  const sources = new WorkbookContentSources(surface.document, snapshot), native = [];
  try {
    await surface.document.fonts.ready;
    const definitions = snapshot.metadata.workbook.charts!, projection = sources.project(surface.source, definitions.map(chart => chart.nodeId));
    for (const chart of definitions) for (const role of ['title', 'category-axis-title', 'value-axis-title'] as const) {
      const nodeId = role === 'title' ? chart.titleNodeId : role === 'category-axis-title' ? chart.categoryAxisTitle?.nodeId : chart.valueAxisTitle?.nodeId;
      if (!nodeId) continue;
      const source = projection.get(chart.nodeId)!, before = source.outerHTML, text = createWorkbookChartText(source, chart, role);
      const batch = layoutWorkbookChartTexts(surface.view, [{ source, chart, role, text, widthPt: 240, heightPt: 60 }]);
      const result = batch.objects[0]!; surface.view.append(result.element);
      const walker = surface.document.createTreeWalker(result.element, 0xffff_ffff); let nodes = 1; while (walker.nextNode()) nodes++;
      native.push({ role, unchanged: source.outerHTML === before, nodes, declared: result.nodes, fonts: batch.fonts.length,
        sourceId: text.element.dataset.sourceNodeId, expectedId: nodeId, widthPt: result.widthPt, scale: result.scale, text: text.projections.map(piece => piece.node.data).join(''),
        units: text.projections.map(piece => piece.to - piece.from), lengths: text.projections.map(piece => piece.node.data.length),
        owned: text.projections.every(piece => source.querySelector('[data-node-id="' + piece.sourceNodeId + '"]')),
        reflected: surface.document.defaultView!.getComputedStyle(text.element).transform });
      result.element.remove();
    }
    const original = definitions[0]!, source = projection.get(original.nodeId)!;
    const body = source.querySelector<HTMLElement>('[data-node-id="' + original.titleNodeId + '"]')!;
    body.innerHTML = '<p data-node-id="chart-title-fit" style="font-family:Liberation Sans;--office-font-ascii:&#34;Liberation Sans&#34;;font-size:18pt;line-height:24pt">Long title with 中文😀禰󠄀 and several words that need multiple lines.</p>';
    const before = source.outerHTML, fits = [];
    for (const fit of ['none', 'normal', 'shape'] as const) {
      const chart = { ...original, titleText: { autoFit: fit, wrap: true, fontScale: 1, lineSpacingReduction: 0 } }, text = createWorkbookChartText(source, chart, 'title');
      const result = layoutWorkbookChartTexts(surface.view, [{ source, chart, role: 'title', text, widthPt: 120, heightPt: 30 }]).objects[0]!;
      fits.push({ fit, width: result.widthPt, height: result.heightPt, scale: result.scale, overflow: result.overflow, unchanged: source.outerHTML === before,
        sourceText: text.projections.map(piece => piece.node.data).join('') });
      result.element.remove();
    }
    const errors = [], first = createWorkbookChartText(source, original, 'title'), second = createWorkbookChartText(source, original, 'title');
    const input = { source, chart: original, role: 'title' as WorkbookChartTextRole, text: first, widthPt: 120, heightPt: 30 };
    for (const kind of ['budget', 'duplicate', 'foreign-title', 'missing-title', 'later-overflow']) {
      try {
        const items = kind === 'duplicate' ? [input, input] : kind === 'foreign-title' ? [{ ...input, chart: { ...original, titleNodeId: 'foreign-title' } }]
          : kind === 'missing-title' ? [{ ...input, role: 'value-axis-title' as const, chart: { ...original, valueAxisTitle: { reference: { expression: '=Data!A1', valueType: 'text' as const } } } }]
          : kind === 'later-overflow' ? [input, { ...input, text: second, chart: { ...original, titleText: { autoFit: 'shape' as const } } }] : [input];
        layoutWorkbookChartTexts(surface.view, items, kind === 'budget' ? { maxNodes: 1 } : kind === 'later-overflow' ? { maxExtentPt: 125 } : {});
        errors.push('accepted');
      } catch (error) { errors.push((error as { code: string }).code); }
    }
    return { native, fits, errors, unchanged: source.outerHTML === before, cleaned: !first.element.parentNode && !second.element.parentNode && !surface.view.querySelector('[data-office-view="workbook-drawing-text-measurement"]') };
  } finally { sources.dispose(); surface.dispose(); container.remove(); }
}
