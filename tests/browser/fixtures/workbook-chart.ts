import type { WorkbookSnapshot } from '@dagent-office/core';
import type { BasicChartData } from '../../../packages/render/src/chart.js';
import { createWorkbookChartView } from '../../../packages/render/src/workbook/chart-view.js';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';

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
      const title = () => {
        const result = surface.document.createElement('div'); result.textContent = '标题😀'; result.style.cssText = 'font:13pt "Liberation Sans";color:rgb(12,34,56)'; return result;
      };
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
      let negativePie = '';
      if (chart.kind === 'pie') {
        try { createWorkbookChartView(source, chart, { ...data, series: data.series.map(series => ({ ...series, values: [-2, ...series.values.slice(1)] })) }, 360, 220, 'negative-pie', options()); negativePie = 'accepted'; }
        catch (error) { negativePie = (error as { code: string }).code; }
      }
      results.push({ kind: chart.kind, nodes, reported: result.nodes, peak: result.peakNodes, exact: exact.nodes === nodes, unchanged: unchanged(), errors, negativePie,
        sourceClasses: result.element.querySelectorAll('[class],[data-node-id]').length,
        svgSize: [svg.clientWidth, svg.clientHeight], text: text.map(node => node.textContent),
        fonts: text.map(node => surface.document.defaultView!.getComputedStyle(node).fontSize), colors: text.map(node => surface.document.defaultView!.getComputedStyle(node).fill),
        bars: [...svg.querySelectorAll('rect')].filter(node => Number(node.getAttribute('width')) > 20 && Number(node.getAttribute('width')) < 100).map(node => Number(node.getAttribute('height'))) });
      result.element.remove();
    }
    return results;
  } finally { sources.dispose(); surface.dispose(); container.remove(); }
}
