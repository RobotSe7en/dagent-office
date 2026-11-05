import type { WorkbookSnapshot } from '@dagent-office/core';
import { createWorkbookNumberFormatter } from '@dagent-office/core';
import type { BasicChartData } from '../../../packages/render/src/chart.js';
import { createWorkbookChartView } from '../../../packages/render/src/workbook/chart-view.js';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookChartText } from '../../../packages/render/src/workbook/drawing-text.js';
import type { WorkbookChartTextRole } from '../../../packages/render/src/workbook/drawing-text.js';

export async function measureWorkbookChartColors(snapshot: WorkbookSnapshot, plots: readonly BasicChartData[]) {
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 520, heightPt: 340 }), sources = new WorkbookContentSources(surface.document, snapshot);
  const charts = snapshot.metadata.workbook.charts!, projected = sources.project(surface.source, charts.map(chart => chart.nodeId)), before = surface.source.innerHTML;
  const read = (index: number, data = plots[index]!) => {
    const chart = charts[index]!, source = projected.get(chart.nodeId)!;
    const result = createWorkbookChartView(source, chart, data, 480, 300, 'chart-colors-' + index, { formatNumber: value => String(value) });
    surface.view.append(result.element); const svg = result.element.querySelector('svg')!, visible = data.series.filter(series => !series.hidden);
    const varied = (chart.varyColors ?? chart.kind === 'pie') && chart.series.length === 1;
    const entries = chart.kind === 'pie' || varied && visible.length === 1 ? data.categories.length : visible.length;
    const color = (node: Element) => surface.document.defaultView!.getComputedStyle(node).getPropertyValue('fill');
    const points = chart.kind === 'column' ? [...svg.querySelectorAll('rect')].slice(1, 1 + visible.length * data.categories.length)
      : chart.kind === 'line' ? [...svg.querySelectorAll('circle')] : [...svg.querySelectorAll('path,circle')].filter(node => node.getAttribute('fill') !== 'none');
    const keys = chart.kind === 'line' ? [...svg.querySelectorAll('line')].slice(-entries).map(node => surface.document.defaultView!.getComputedStyle(node).stroke)
      : [...svg.querySelectorAll('rect')].slice(-entries).map(color);
    const text = [...svg.querySelectorAll('text')].slice(-entries).map(node => node.textContent);
    const words: Record<string, number> = {};
    for (const node of svg.querySelectorAll('text')) if (['One', 'Two', 'Three', 'Series1', 'Series2'].includes(node.textContent ?? '')) words[node.textContent!] = (words[node.textContent!] ?? 0) + 1;
    const value = { points: points.map(color), keys, text, words, nodes: result.nodes, actualNodes: result.element.querySelectorAll('*').length + [...svg.querySelectorAll('text')].reduce((sum, node) => sum + (node.textContent ? 1 : 0), 0) + 1 };
    result.element.remove(); return value;
  };
  try {
    const normal = charts.map((_, index) => read(index)), multiple = charts.findIndex(chart => chart.kind === 'column' && chart.varyColors && chart.series.length === 2);
    const hidden = read(multiple, { ...plots[multiple]!, series: plots[multiple]!.series.map((series, index) => ({ ...series, hidden: index > 0 })) });
    const pie = charts.findIndex(chart => chart.kind === 'pie' && chart.varyColors), source = plots[pie]!;
    const compacted = read(pie, { ...source, categories: [source.categories[0]!, source.categories[2]!], sourceIndexes: [0, 2], series: source.series.map(series => ({ ...series, values: [series.values[0]!, series.values[2]!], ...(series.missing ? { missing: [series.missing[0], series.missing[2]] } : {}) })) });
    return { normal, hidden, compacted, unchanged: surface.source.innerHTML === before };
  } finally { sources.dispose(); surface.dispose(); container.remove(); }
}
import { layoutWorkbookChartTexts, layoutWorkbookChartTitleContents } from '../../../packages/render/src/workbook/drawing-text-layout.js';
import type { WorkbookChartTitleContentInput } from '../../../packages/render/src/workbook/drawing-text-layout.js';
import type { WorkbookChartBinding } from '../../../packages/render/src/workbook/chart-values.js';
import { layoutWorkbookChart } from '../../../packages/render/src/workbook/chart-layout.js';

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
      for (const kind of ['node-budget', 'point-budget', 'text-budget', 'formatter', 'invalid-data', 'source-title', 'missing-title', 'axis-scale', 'formatted-text-budget', 'coordinate-scale']) {
        try {
          const selected = kind === 'invalid-data' || kind === 'coordinate-scale' ? { ...data, series: data.series.map(series => ({ ...series, values: [kind === 'invalid-data' ? NaN : Number.MAX_VALUE, ...series.values.slice(1)] })) } : data;
          const appearance = kind === 'formatter' ? { ...chart, showLabels: true, labelContents: ['value' as const] } : kind === 'axis-scale' ? { ...chart, kind: 'column' as const, minimum: 1e16, maximum: 1e16 + 4, majorUnit: .1, showValueAxis: false } : kind === 'coordinate-scale' ? { ...chart, kind: 'column' as const, minimum: 0, maximum: 1, majorUnit: .1 } : kind === 'formatted-text-budget' ? { ...chart, kind: 'column' as const } : chart;
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
      const negativePies = [];
      if (chart.kind === 'pie') {
        for (const values of [[-2, 5, 8], [2, -1, -1], [-Number.MAX_VALUE, Number.MAX_VALUE, 0], [-4, 0, null], [0, 0, null], [-Number.MIN_VALUE, Number.MIN_VALUE, 0]]) {
          const signed = { ...data, series: [{ ...data.series[0]!, values }] }, positive = { ...data, series: [{ ...data.series[0]!, values: values.map(value => value === null ? null : Math.abs(value)) }] };
          const numbers: number[] = [], appearance = { ...chart, showLabels: true, labelContents: ['value', 'percent'] as const };
          const graphics = createWorkbookChartView(source, appearance, signed, 360, 220, 'negative-pie', { title: title(), formatNumber(value, format) { numbers.push(value); return formatNumber(value, format); } });
          const absolute = createWorkbookChartView(source, { ...chart, showLabels: false }, positive, 360, 220, 'absolute-pie', options());
          const geometry = (element: HTMLElement) => [...element.querySelectorAll('svg > path, svg > circle')].map(node => [node.tagName, node.getAttribute('d'), node.getAttribute('cx'), node.getAttribute('cy'), node.getAttribute('r')]);
          negativePies.push({ values, labels: numbers.filter((_, index) => index % 2 === 0), fractions: numbers.filter((_, index) => index % 2 === 1), geometry: geometry(graphics.element), absolute: geometry(absolute.element),
            diagnostics: graphics.diagnostics, positiveDiagnostics: absolute.diagnostics, immutable: Object.isFrozen(graphics.diagnostics) && graphics.diagnostics.every(Object.isFrozen) });
        }
      }
      const smallScales = [], zeroScales = [];
      if (chart.kind === 'column') for (const sign of [1, -1, 0, null]) {
        const appearance = { ...chart }; delete appearance.minimum; delete appearance.maximum; delete appearance.majorUnit;
        const values = { ...data, series: data.series.map(series => ({ ...series, values: series.values.map(value => value === null || sign === null ? null : Math.abs(value) * sign / 128) })) };
        const graphics = createWorkbookChartView(source, appearance, values, 360, 220, 'small-' + sign, options());
        const rects = [...graphics.element.querySelectorAll('svg > rect')], plot = rects[0]!;
        const bars = rects.filter(node => Number(node.getAttribute('width')) > 20 && Number(node.getAttribute('width')) < 100);
        const fraction = Math.max(0, ...bars.map(node => Number(node.getAttribute('height')))) / Number(plot.getAttribute('height'));
        if (sign) smallScales.push(fraction); else zeroScales.push(fraction);
      }
      results.push({ kind: chart.kind, nodes, reported: result.nodes, peak: result.peakNodes, exact: exact.nodes === nodes, unchanged: unchanged(), errors, negativePies, smallScales, zeroScales, titleProjection, titleText: richSource.textContent,
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


export async function measureWorkbookChartPositions(snapshot: WorkbookSnapshot) {
  const container = document.createElement('div'); document.body.append(container);
  const css = '.chart-index-base{font-family:"Liberation Sans";font-size:11pt;color:#0000ff;--office-text-color:#0000ff;--office-text-source:direct;--office-text-lum-mod:1;--office-text-lum-off:0;--office-text-tint:0;--office-text-shade:1;--office-text-alpha:1}.chart-index-red{color:#ff0000;--office-text-color:#ff0000;--office-text-source:direct;--office-text-lum-mod:1;--office-text-lum-off:0;--office-text-tint:0;--office-text-shade:1;--office-text-alpha:1}';
  const surface = await createWorkbookSurface(container, { ...snapshot, styles: [...snapshot.styles, { path: 'chart-index.css', css }] }, { widthPt: 400, heightPt: 260 });
  const sources = new WorkbookContentSources(surface.document, snapshot);
  try {
    const original = snapshot.metadata.workbook.charts![0]!, series = original.series[0]!, projection = sources.project(surface.source, [original.nodeId]);
    const source = projection.get(original.nodeId)!, before = source.outerHTML;
    const chart = { ...original, kind: 'line' as const, titleVisible: false, series: [series], legend: 'none' as const, blankAs: 'zero' as const,
      showLabels: true, labelContents: ['value'] as const, textCssClass: 'chart-index-base', pointStyles: [{ seriesId: series.seriesId, index: 2, cssClass: 'chart-index-base', labelCssClass: 'chart-index-red' }] };
    const options = { formatNumber: (value: number) => String(value) };
    const compact: BasicChartData = { categories: ['A', 'C'], sourceIndexes: [0, 2], series: [{ seriesId: series.seriesId, name: 'Values', values: [2, 8] }] };
    const graphics = createWorkbookChartView(source, chart, compact, 360, 220, 'indexes', options); surface.view.append(graphics.element);
    const colors = [...graphics.element.querySelectorAll('svg > text')].filter(node => ['2', '8'].includes(node.textContent ?? '')).slice(-2).map(node => surface.document.defaultView!.getComputedStyle(node).fill);
    const plain: BasicChartData = { categories: ['A', 'B', 'C', 'D', 'E'], series: [{ seriesId: series.seriesId, name: 'Values', values: [2, null, 8, null, 5] }] };
    const missing: BasicChartData = { ...plain, series: [{ ...plain.series[0]!, missing: [undefined, 'gap', undefined, 'span', undefined] }] };
    const path = (data: BasicChartData) => [...createWorkbookChartView(source, { ...chart, showLabels: false }, data, 360, 220, 'gaps', options).element.querySelectorAll('svg > path')].map(node => node.getAttribute('d')).join(' ');
    const hidden = createWorkbookChartView(source, { ...chart, kind: 'pie' }, { ...compact, series: [{ ...compact.series[0]!, hidden: true, values: [-2, 8] }] }, 360, 220, 'hidden', options);
    const second = { ...series, seriesId: 'second-series' };
    const masked = createWorkbookChartView(source, { ...chart, series: [series, second], legend: 'right', seriesStyles: [], showLabels: false },
      { ...compact, series: [{ ...compact.series[0]!, hidden: true }, { ...compact.series[0]!, seriesId: second.seriesId }] }, 360, 220, 'masked-series', options);
    const maskedColors = [...masked.element.querySelectorAll('svg > path,svg > line')].filter(node => node.localName === 'path' || node.getAttribute('x2') !== node.getAttribute('x1') && Number(node.getAttribute('x2')) - Number(node.getAttribute('x1')) === 22).map(node => node.getAttribute('stroke'));
    const errors: string[] = [];
    for (const data of [{ ...compact, sourceIndexes: [2, 2] }, { ...compact, sourceIndexes: [0] }, { ...compact, series: [{ ...compact.series[0]!, missing: ['gap' as const, undefined] }] }]) {
      try { createWorkbookChartView(source, chart, data, 360, 220, 'invalid', options); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    }
    return { colors, maskedColors, plain: path(plain), missing: path(missing), hiddenPaths: hidden.element.querySelectorAll('svg > path').length, hiddenDiagnostics: hidden.diagnostics, errors, unchanged: source.outerHTML === before };
  } finally { sources.dispose(); surface.dispose(); container.remove(); }
}

export async function measureWorkbookChartTitleContents(snapshot: WorkbookSnapshot, bindings: readonly WorkbookChartBinding[]) {
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 600, heightPt: 300 });
  const sources = new WorkbookContentSources(surface.document, snapshot);
  try {
    const charts = snapshot.metadata.workbook.charts!, projection = sources.project(surface.source, charts.map(chart => chart.nodeId)), before = surface.source.innerHTML;
    const inputs: WorkbookChartTitleContentInput[] = [];
    for (const chart of charts) for (const role of ['title', 'category-axis-title', 'value-axis-title'] as const) {
      const definition = role === 'title' ? { nodeId: chart.titleNodeId, reference: chart.titleReference } : role === 'category-axis-title' ? chart.categoryAxisTitle : chart.valueAxisTitle;
      if (!definition) continue;
      const binding = bindings.find(binding => binding.chartId === chart.chartId)!;
      if (binding.kind !== 'available') throw Error('Current title source unavailable');
      inputs.push({ source: projection.get(chart.nodeId)!, chart, role, widthPt: 180, heightPt: 48, ...(definition.reference ? { value: binding.titles[role]! } : {}) });
    }
    const batch = layoutWorkbookChartTitleContents(surface.view, inputs);
    for (const [index, object] of batch.objects.entries()) { object.element.style.top = index * 50 + 'pt'; surface.view.append(object.element); }
    const stable = surface.view.innerHTML;
    const results = batch.objects.map((object, index) => {
      const walker = surface.document.createTreeWalker(object.element, 0xffff_ffff); let nodes = 1; while (walker.nextNode()) nodes++;
      return { reference: inputs[index]!.value !== undefined, role: inputs[index]!.role, text: object.text.element.textContent,
        projections: object.text.projections.map(projection => ({ from: projection.from, to: projection.to, text: projection.node.data })),
        nodes, declared: object.nodes, sourceId: object.text.element.dataset.sourceNodeId, expectedId: inputs[index]!.value !== undefined ? inputs[index]!.chart.nodeId : inputs[index]!.chart.titleNodeId,
        scale: object.scale, fontSizes: [...object.text.element.querySelectorAll('[data-office-script]')].map(node => surface.document.defaultView!.getComputedStyle(node).fontSize),
        sourceClasses: object.element.querySelectorAll('[class],[data-node-id]').length };
    });
    const errors = [];
    for (const kind of ['nodes', 'text', 'missing', 'rich-value', 'foreign', 'late-extent']) {
      try {
        const first = inputs[0]!, { value: _value, ...missing } = first;
        const items = kind === 'missing' ? [missing] : kind === 'rich-value' ? [{ ...inputs.at(-1)!, value: 'wrong' }]
          : kind === 'foreign' ? [{ ...first, chart: { ...first.chart, nodeId: 'other-chart' } }]
          : kind === 'late-extent' ? [first, { ...first, widthPt: 201 }] : inputs;
        layoutWorkbookChartTitleContents(surface.view, items, kind === 'nodes' ? { maxNodes: 1 } : kind === 'text' ? { maxCodeUnits: 1 } : kind === 'late-extent' ? { maxExtentPt: 200 } : {});
        errors.push('accepted');
      } catch (error) { errors.push((error as { code: string }).code); }
      if (surface.view.innerHTML !== stable || surface.source.innerHTML !== before) throw Error('Failed title candidate changed the active view or source');
    }
    const empty = layoutWorkbookChartTitleContents(surface.view, [{ ...inputs[0]!, value: '' }]).objects[0]!;
    return { results, errors, empty: { projections: empty.text.projections.length, visible: empty.text.element.textContent?.replace(/\u200b/g, '') },
      fonts: batch.fonts.length, unchanged: surface.source.innerHTML === before, cleaned: !surface.source.querySelector('[data-office-view="workbook-reference-title-source"]') && !surface.view.querySelector('[data-office-view="workbook-drawing-text-measurement"]') };
  } finally { sources.dispose(); surface.dispose(); container.remove(); }
}

export async function measureWorkbookComposedChart(snapshot: WorkbookSnapshot, bindings: readonly WorkbookChartBinding[]) {
  const container = document.createElement('div'); document.body.append(container);
  const css = '.axis-top{--office-axis-position:t}.axis-right{--office-axis-position:r}svg,text,div{font-size:99pt!important}';
  const surface = await createWorkbookSurface(container, { ...snapshot, styles: [...snapshot.styles, { path: 'chart-layout-probe.css', css }] }, { widthPt: 600, heightPt: 300 });
  const sources = new WorkbookContentSources(surface.document, snapshot), formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
  try {
    const original = snapshot.metadata.workbook.charts![0]!, binding = bindings.find(item => item.chartId === original.chartId)!;
    if (binding.kind !== 'available') throw Error('Current chart must be available');
    const source = sources.project(surface.source, [original.nodeId]).get(original.nodeId)!, before = source.outerHTML, results = [];
    let active: ReturnType<typeof layoutWorkbookChart> | undefined;
    for (const opposite of [false, true]) for (const legend of ['none', 'left', 'right', 'top', 'bottom'] as const) {
      const chart = { ...original, legend, ...(opposite ? { categoryAxisCssClass: 'axis-top', valueAxisCssClass: 'axis-right' } : {}) };
      const result = layoutWorkbookChart(surface.view, source, chart, binding, 420, 300, 'composed-chart', formatter);
      surface.view.replaceChildren(result.element); active = result;
      const base = result.element.getBoundingClientRect(), rect = (element: Element) => { const box = element.getBoundingClientRect(); return { x: (box.x - base.x) * .75, y: (box.y - base.y) * .75, width: box.width * .75, height: box.height * .75 }; };
      const walker = surface.document.createTreeWalker(result.element, 0xffff_ffff); let nodes = 1; while (walker.nextNode()) nodes++;
      results.push({ opposite, legend, layout: result.layout, nodes, declared: result.nodes, unchanged: before === source.outerHTML,
        plot: rect(result.element.querySelector('svg > rect')!), titles: result.texts.map(text => ({ role: text.text.element.dataset.officeChartTextRole, rect: rect(text.element), text: text.text.element.textContent, projections: text.text.projections.length })),
        titleOverflow: result.overflowingTitles, styles: result.element.querySelectorAll('[class],[data-node-id]').length });
    }
    const current = surface.view.innerHTML, errors = [];
    for (const kind of ['nodes', 'text', 'points', 'small', 'foreign-binding', 'formatter']) {
      try {
        layoutWorkbookChart(surface.view, source, original, kind === 'foreign-binding' ? { ...binding, chartId: 'foreign-chart' } : binding,
          kind === 'small' ? 30 : 420, 300, 'failed-chart', kind === 'formatter' ? { ...formatter, format: () => { throw Error('format rejected'); } } : formatter,
          kind === 'nodes' ? { maxNodes: active!.nodes - 1 } : kind === 'text' ? { maxCodeUnits: 1 } : kind === 'points' ? { maxPoints: 1 } : {});
        errors.push('accepted');
      } catch (error) { errors.push((error as { code?: string }).code ?? (error as Error).message); }
      if (surface.view.innerHTML !== current || source.outerHTML !== before) throw Error('Chart candidate changed active source or view on failure');
    }
    return { results, errors, unchanged: before === source.outerHTML, cleaned: !surface.source.querySelector('[data-office-view="workbook-reference-title-source"]') && !surface.view.querySelector('[data-office-view="workbook-drawing-text-measurement"]') };
  } finally { formatter.dispose(); sources.dispose(); surface.dispose(); container.remove(); }
}

export async function measureWorkbookChartFormats(snapshot: WorkbookSnapshot, bindings: readonly WorkbookChartBinding[]) {
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 600, heightPt: 400 });
  const sources = new WorkbookContentSources(surface.document, snapshot), formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
  try {
    const charts = snapshot.metadata.workbook.charts!, projection = sources.project(surface.source, charts.map(chart => chart.nodeId)), before = surface.source.innerHTML;
    return charts.map(chart => {
      const source = projection.get(chart.nodeId)!, binding = bindings.find(item => item.chartId === chart.chartId)!;
      if (binding.kind !== 'available') throw Error('Native formats must have current data');
      const result = layoutWorkbookChart(surface.view, source, chart, binding, 480, 300, 'formatted-chart', formatter); surface.view.replaceChildren(result.element);
      const labels = (role: string) => [...result.element.querySelectorAll('[data-sdk-chart-label="' + role + '"]')].map(element => element.textContent);
      return { valueAxis: labels('value-axis'), categoryAxis: labels('category-axis'), labels: labels('data-label'),
        unchanged: surface.source.innerHTML === before, sourceIndexes: binding.data.sourceIndexes };
    });
  } finally { formatter.dispose(); sources.dispose(); surface.dispose(); container.remove(); }
}
