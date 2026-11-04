import { createWorkbook, openDocument } from "@dagent-office/core";
import type { WorkbookDocumentMetadata, WorkbookManifest } from "@dagent-office/core";
import { AxisDimensions } from "../../../packages/render/src/workbook/dimensions.js";
import { WorkbookWindowLoader } from "../../../packages/render/src/workbook/window-loader.js";
import { modelWindowSource } from "../../../packages/render/src/workbook/window-source.js";
import { replaceWorkbookWindow } from "../../../packages/render/src/workbook/window-dom.js";
import { hitWorkbookCell, WorkbookMergeIndex } from "../../../packages/render/src/workbook/projection.js";
import type { WorkbookViewport } from "../../../packages/render/src/workbook/viewport.js";
import { measureNumberDisplay } from "../../../packages/render/src/workbook/number-layout.js";
import { formatGeneralNumber } from "../../../packages/core/src/workbook/format-general.js";
import { measureWorkbookColumnFont } from "../../../packages/render/src/workbook/column-font.js";
import { resolvedWorkbookColor } from "../../../packages/render/src/workbook/color.js";
import { WorkbookCellSources } from "../../../packages/render/src/workbook/cell-source.js";
import { workbookTextFonts } from "../../../packages/render/src/workbook/text-fonts.js";
import { createWorkbookTextView } from "../../../packages/render/src/workbook/text-view.js";
import { workbookCellFill } from "../../../packages/render/src/workbook/cell-fill.js";
import { createWorkbookBorderView } from "../../../packages/render/src/workbook/border-view.js";
import type { WorkbookBorderEdge } from "../../../packages/render/src/workbook/border-view.js";
import { WorkbookTextDirection } from "../../../packages/render/src/workbook/text-direction.js";

const encode = (text: string) => new TextEncoder().encode(text);
async function main() {
  const files = new Map(createWorkbook({ documentId: 'grid-geometry' }));
  const manifest = JSON.parse(new TextDecoder().decode(files.get('manifest.json')!)) as WorkbookManifest;
  const metadata = JSON.parse(new TextDecoder().decode(files.get(manifest.metadata)!)) as WorkbookDocumentMetadata;
  const sheet = metadata.workbook.sheets[0]!;
  metadata.workbook.columnWidthFontCssClass = 'normal-metric';
  metadata.workbook.themeCssClass = 'grid-theme';
  files.set('styles.css', encode(new TextDecoder().decode(files.get('styles.css')!) + '.normal-metric{font-family:"Liberation Serif";font-size:20pt;font-weight:normal}.grid-theme{--office-theme-accent1:#4472c4}.cell-highlight{font-family:"Liberation Serif";font-size:13pt;font-style:italic}.rich-style{font-weight:bold;color:#993355}table[data-row="9984"] tr[data-row="9999"] td[data-column="9"]{background-color:var(--office-theme-accent1)}'));
  sheet.frozen = { rows: 2, columns: 1 }; sheet.merges = [{ from: { row: 0, column: 0 }, to: { row: 6, column: 6 } }];
  for (const [row, column, text] of [[0, 0, '合并😀禰󠄀'], [9999, 9, '一万行'], [1_048_575, 16_383, '末格']] as const) {
    const blockRow = Math.floor(row / 256) * 256, blockColumn = Math.floor(column / 32) * 32, path = `sheets/sheet-1/${blockRow}-${blockColumn}.html`;
    const bytes = encode(`<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="${blockRow}" data-column="${blockColumn}"><tbody><tr data-row="${row}"><td data-column="${column}" data-value-type="text"${row === 9999 ? ' class="cell-highlight"' : ''}>${row === 9999 ? '<span class="rich-style">一万</span>行' : row === 0 ? '合并😀<span>禰</span><span>󠄀</span>' : text}</td></tr></tbody></table></body></html>`);
    const digest = await crypto.subtle.digest('SHA-256', bytes);
    manifest.chunks.push({ sheetId: 'sheet-1', row: blockRow, column: blockColumn, path, sha256: [...new Uint8Array(digest)].map(byte => byte.toString(16).padStart(2, '0')).join('') });
    files.set(path, bytes);
  }
  files.set('manifest.json', encode(JSON.stringify(manifest))); files.set(manifest.metadata, encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: 'workbook' }), before = model.serialize(), loader = new WorkbookWindowLoader(modelWindowSource(model));
  const dimensions = { rows: new AxisDimensions(1_048_576, 10, []), columns: new AxisDimensions(16_384, 20, []) }, merges = new WorkbookMergeIndex(sheet.merges);
  const container = document.createElement('div'); container.id = 'workbook-grid-probe'; container.style.cssText = 'position:absolute;left:0;top:0;background:white;color:black'; document.body.replaceChildren(container);
  const style = document.createElement('style'); style.textContent = model.snapshot().styles.map(style => style.css).join('\n'); document.head.append(style);
  let rightToLeft = false;
  const host = {
    async render(viewport: WorkbookViewport, rtl = false) {
      const current = await loader.load('sheet-1', dimensions, viewport);
      if (current) {
        rightToLeft = rtl; const sources = new WorkbookCellSources(container, model.snapshot(), loader.frame!);
        try {
          let remainingNodes = 50_000;
          replaceWorkbookWindow(container, loader.frame!, (body, cell, projection) => {
            const source = sources.get(projection.row, projection.column)!;
            body.style.backgroundColor = workbookCellFill(document, getComputedStyle(source.element));
            if (cell?.value.type === 'text') { const text = createWorkbookTextView(source, remainingNodes); remainingNodes -= text.nodes; body.append(text.element); }
          }, rtl);
        } finally { sources.dispose(); }
      }
      return current;
    },
    inspect() {
      const root = container.firstElementChild as HTMLElement, bounds = root.getBoundingClientRect();
      const cells = [...root.querySelectorAll<HTMLElement>('[data-office-view="workbook-cell"]')];
      return { cellCount: cells.length, accessibleCells: root.querySelectorAll('[role="gridcell"]').length, rows: [...root.querySelectorAll('[role="row"]')].map(row => Number(row.getAttribute('aria-rowindex'))),
        sourceUnchanged: [...before].every(([path, bytes]) => { const after = model.serialize().get(path)!; return bytes.length === after.length && bytes.every((byte, index) => byte === after[index]); }),
        merges: cells.filter(cell => cell.dataset.row === '0' && cell.dataset.column === '0').map(cell => {
          const clip = cell.getBoundingClientRect(), body = cell.firstElementChild!.getBoundingClientRect();
          const x = (clip.x - bounds.x + clip.width / 2) * .75, y = (clip.y - bounds.y + clip.height / 2) * .75;
          return { pane: cell.dataset.pane, text: cell.textContent, clip: { x: (clip.x - bounds.x) * .75, y: (clip.y - bounds.y) * .75, width: clip.width * .75, height: clip.height * .75 },
            body: { x: (body.x - bounds.x) * .75, y: (body.y - bounds.y) * .75, width: body.width * .75, height: body.height * .75 },
            hit: hitWorkbookCell(dimensions, loader.frame!.layout, merges, rightToLeft ? loader.frame!.layout.widthPt - x : x, y) };
        }) };
    },
    failedPaint() {
      const previous = container.firstChild; let failed = false;
      try { replaceWorkbookWindow(container, loader.frame!, () => { throw new Error('paint failed'); }, rightToLeft); } catch { failed = true; }
      return failed && previous === container.firstChild;
    },
    async numericMeasurement() {
      await document.fonts.ready;
      const canvas = document.createElement('canvas'), context = canvas.getContext('2d')!, font = '24px "Liberation Serif"'; context.font = font;
      const measure = (text: string) => context.measureText(text).width;
      const rounded = formatGeneralNumber(25.76, '.', measure('26') + .25, measure), scientific = formatGeneralNumber(123456789012, '.', 1000, measure);
      const source = [{ kind: 'text' as const, text: '$' }, { kind: 'fill' as const, character: 'f' }, { kind: 'text' as const, text: '12.5' }, { kind: 'space' as const, character: ')' }];
      const fitted = measureNumberDisplay(source, 140, measure), box = document.createElement('div');
      box.style.cssText = 'position:absolute;left:0;top:0;white-space:pre;display:flex;width:max-content'; box.style.font = font;
      for (const part of fitted.parts) {
        const span = document.createElement('span'); span.style.flex = 'none';
        if (part.kind === 'text') span.textContent = part.text; else span.style.width = part.width + 'px';
        box.append(span);
      }
      container.append(box);
      const actual = box.getBoundingClientRect().width; box.remove();
      return { rounded, scientific, actual, predicted: fitted.width, overflow: fitted.overflow, fill: fitted.parts[1] };
    },
    async columnFontMeasurement() {
      await document.fonts.ready;
      const unrelated = document.createElement('span'); unrelated.style.fontFamily = '"Unrelated Missing Probe Font"'; unrelated.textContent = 'outside'; container.append(unrelated);
      const originalFamily = unrelated.style.fontFamily;
      const book = model.snapshot().metadata.workbook, measured = measureWorkbookColumnFont(container, book, 'sheet-1');
      const { columnWidthFontCssClass: omitted, ...legacy } = book;
      const fallback = measureWorkbookColumnFont(container, legacy, 'sheet-1');
      const spans = [...'0123456789'].map(digit => {
        const span = document.createElement('span'); span.className = 'normal-metric'; span.textContent = digit; span.style.cssText = 'position:absolute;display:inline-block;white-space:pre'; container.append(span); return span;
      });
      const actual = Math.max(...spans.map(span => span.getBoundingClientRect().width)); for (const span of spans) span.remove();
      const unrelatedUnchanged = unrelated.style.fontFamily === originalFamily; unrelated.remove();
      return { measured, fallback, actual, unrelatedUnchanged, retainedProbes: container.querySelectorAll('.normal-metric').length };
    },
    colorMeasurement() {
      const wrapper = document.createElement('div'); wrapper.style.setProperty('--office-theme-accent1', '#4472c4'); container.append(wrapper);
      const values: (string | undefined)[] = [], errors: string[] = [], fills: string[] = [], fillErrors: string[] = []; let unchanged = true;
      try {
        for (const [color, tint, nativeAlpha] of [['var(--office-theme-accent1)', '.4', '0'], ['var(--office-theme-accent1)', '-.2', '128'], ['currentColor', '0', '255'], ['red', '0', ''], ['rgba(20,40,60,.5)', '0', ''], ['hsl(120,100%,50%)', '0', ''], ['', '', '']]) {
          const node = document.createElement('span'); node.style.color = 'rgb(20,40,60)';
          node.style.setProperty('--office-cell-fill-color', color!); node.style.setProperty('--office-cell-fill-tint', tint!); node.style.setProperty('--office-cell-fill-native-alpha', nativeAlpha!); wrapper.append(node);
          const before = node.style.cssText; values.push(resolvedWorkbookColor(document, getComputedStyle(node), '--office-cell-fill')); unchanged &&= before === node.style.cssText;
        }
        for (const [color, tint] of [['not-a-color', '0'], ['red', '1px'], ['red', '2']]) {
          const node = document.createElement('span'); node.style.setProperty('--office-cell-fill-color', color!); node.style.setProperty('--office-cell-fill-tint', tint!); wrapper.append(node);
          try { resolvedWorkbookColor(document, getComputedStyle(node), '--office-cell-fill'); } catch (error) { errors.push((error as { code: string }).code); }
        }
        for (const [kind, color] of [['', ''], ['none', 'red'], ['solid', 'var(--office-theme-accent1)'], ['solid', ''], ['darkGrid', 'red']]) {
          const node = document.createElement('span'); node.style.backgroundColor = 'rgb(1,2,3)'; node.style.setProperty('--office-cell-fill-kind', kind!);
          node.style.setProperty('--office-cell-fill-color', color!); node.style.setProperty('--office-cell-fill-tint', '.4'); wrapper.append(node);
          const before = node.style.cssText;
          try { fills.push(workbookCellFill(document, getComputedStyle(node))); } catch (error) { fillErrors.push((error as { code: string }).code); }
          unchanged &&= before === node.style.cssText;
        }
        return { values, errors, fills, fillErrors, unchanged };
      } finally { wrapper.remove(); }
    },
    async sourceStyleMeasurement() {
      await loader.load('sheet-1', dimensions, { widthPt: 240, heightPt: 80, scrollLeftPt: 0, scrollTopPt: 99_970 });
      const previous = container.firstChild, source = new WorkbookCellSources(container, model.snapshot(), loader.frame!);
      const cell = source.get(9999, 9)!, plain = source.get(9999, 8)!, style = getComputedStyle(cell.element), rich = getComputedStyle(cell.element.querySelector('span')!);
      const result = { text: cell.element.textContent, font: style.fontFamily, size: style.fontSize, italic: style.fontStyle, background: style.backgroundColor,
        bold: rich.fontWeight, foreground: rich.color, plain: plain.element.className, blockRow: cell.element.closest('table')!.dataset.row, row: cell.element.parentElement!.dataset.row,
        bound: source.get(9999, 16_384) === undefined, before: cell.element.outerHTML, disposed: false, noPartial: false, retained: -1, rejected: [] as string[] };
      source.dispose(); source.dispose();
      try { source.get(9999, 9); } catch { result.disposed = true; }
      try { new WorkbookCellSources(container, model.snapshot(), loader.frame!, 5); } catch { result.noPartial = previous === container.firstChild; }
      const frame = loader.frame!, original = frame.getCell(0, 0)!;
      for (const content of ['<img src="https://example.invalid/workbook-injected" onerror="window.workbookInjected=true">', '<script>window.workbookInjected=true</script>', '<span onclick="window.workbookInjected=true">bad</span>', '<svg><text>bad</text></svg>', '<span>wrong source text</span>']) {
        const corrupt = { ...frame, getCell: (row: number, column: number) => row === 0 && column === 0 ? { ...original, html: '<td data-column="0" data-value-type="text">' + content + '</td>' } : frame.getCell(row, column) };
        try { new WorkbookCellSources(container, model.snapshot(), corrupt); } catch (error) { result.rejected.push((error as { code: string }).code); }
      }
      result.retained = container.querySelectorAll('[data-office-view="workbook-source-measurement"]').length;
      return result;
    },
    async textFontMeasurement() {
      await loader.load('sheet-1', dimensions, { widthPt: 240, heightPt: 80, scrollLeftPt: 0, scrollTopPt: 99_970 });
      const source = new WorkbookCellSources(container, model.snapshot(), loader.frame!);
      try {
        const merged = source.get(0, 0)!, before = merged.element.outerHTML, runs = workbookTextFonts(merged), rich = workbookTextFonts(source.get(9999, 9)!);
        const copy = merged.element.cloneNode(true) as HTMLTableCellElement; copy.append(document.createElement('br'), document.createTextNode('A')); merged.element.parentElement!.append(copy);
        const withBreak = workbookTextFonts({ ...merged, element: copy }); copy.remove();
        let limited = false; try { workbookTextFonts(merged, 1); } catch (error) { limited = (error as { code: string }).code === 'ResourceLimitExceeded'; }
        return { runs, rich, withBreak, limited, unchanged: before === merged.element.outerHTML, frozen: Object.isFrozen(runs) && Object.isFrozen(runs[0]!.font) };
      } finally { source.dispose(); }
    },
    async textViewMeasurement() {
      await document.fonts.ready;
      await loader.load('sheet-1', dimensions, { widthPt: 240, heightPt: 80, scrollLeftPt: 0, scrollTopPt: 99_970 });
      const sources = new WorkbookCellSources(container, model.snapshot(), loader.frame!), cell = sources.get(9999, 9)!, merged = sources.get(0, 0)!;
      const original = cell.element.outerHTML, rich = createWorkbookTextView(cell), mixed = createWorkbookTextView(merged);
      const copy = cell.element.cloneNode(true) as HTMLTableCellElement;
      copy.style.setProperty('--office-cell-foreground-color', 'var(--office-theme-accent1)'); copy.style.setProperty('--office-cell-foreground-tint', '.4');
      copy.style.setProperty('--office-font-underline', 'double'); copy.style.setProperty('--office-font-strike', 'true');
      copy.replaceChildren(document.createTextNode('A'), document.createElement('br'));
      const plain = document.createElement('span'); plain.style.setProperty('--office-font-underline', 'none'); plain.style.setProperty('--office-font-strike', 'false'); plain.textContent = 'B'; copy.append(plain);
      cell.element.parentElement!.append(copy);
      try {
        const decorated = createWorkbookTextView({ ...cell, element: copy });
        const box = document.createElement('div'); box.id = 'workbook-text-view-probe'; box.style.cssText = 'position:absolute;left:400px;top:50px;width:500px;height:300px;background:white';
        for (const result of [rich, mixed, decorated]) { const row = document.createElement('div'); row.style.cssText = 'display:block;width:400px;min-height:50px'; row.append(result.element); box.append(row); }
        container.append(box);
        const describe = (result: typeof rich) => result.projections.map(projection => {
          const range = document.createRange(); range.selectNode(projection.node); const bounds = range.getBoundingClientRect(), style = getComputedStyle(projection.node.parentElement!);
          return { text: projection.node.data, from: projection.from, to: projection.to, family: style.fontFamily, weight: style.fontWeight, color: style.color, x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height };
        });
        const failures: string[] = [];
        try { createWorkbookTextView(cell, 2); } catch (error) { failures.push((error as { code: string }).code); }
        copy.style.setProperty('--office-font-underline', 'doubleAccounting');
        try { createWorkbookTextView({ ...cell, element: copy }); } catch (error) { failures.push((error as { code: string }).code); }
        copy.style.setProperty('--office-font-underline', 'unknown');
        try { createWorkbookTextView({ ...cell, element: copy }); } catch (error) { failures.push((error as { code: string }).code); }
        const a = decorated.projections[0]!.node.parentElement!, b = decorated.projections[2]!.node.parentElement!;
        return { rich: describe(rich), mixed: describe(mixed), decorated: describe(decorated), decoration: { strike: getComputedStyle(a).textDecorationLine, strikeStyle: getComputedStyle(a).textDecorationStyle,
          underline: getComputedStyle(a.parentElement!).textDecorationLine, underlineStyle: getComputedStyle(a.parentElement!).textDecorationStyle, plain: getComputedStyle(b).textDecorationLine },
          text: decorated.element.textContent, failures, unchanged: original === cell.element.outerHTML,
          frozen: Object.isFrozen(rich) && Object.isFrozen(rich.projections) && Object.isFrozen(rich.projections[0]) };
      } finally { copy.remove(); sources.dispose(); }
    },
    async borderViewMeasurement(edges: readonly WorkbookBorderEdge[]) {
      const layer = createWorkbookBorderView(document, { x: 20, y: 20, width: 200, height: 500 }, edges); layer.id = 'workbook-border-view-probe'; container.append(layer);
      const strokes = [...layer.querySelectorAll('line')].map(line => ({ from: [Number(line.getAttribute('x1')), Number(line.getAttribute('y1'))], to: [Number(line.getAttribute('x2')), Number(line.getAttribute('y2'))],
        width: Number(line.getAttribute('stroke-width')), dash: (line.getAttribute('stroke-dasharray') ?? '').split(' ').filter(Boolean).map(Number), cap: line.getAttribute('stroke-linecap') }));
      const rasterImages: string[] = [], rasterSizes: number[][] = [];
      const raster = async (x: number, kind: WorkbookBorderEdge['kind']) => {
        const svg = createWorkbookBorderView(document, { x, y: 0, width: 40, height: 20 }, [{ kind, color: 'rgba(255,0,0,1)', from: { x: -3, y: 10 }, to: { x: 80, y: 10 } }]);
        // Rasterize a standalone image, without the pane's DOM placement. SVG
        // image intrinsic dimensions come from attributes, not CSS box sizes.
        svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg'); svg.setAttribute('width', '160'); svg.setAttribute('height', '80');
        svg.style.cssText = 'width:160px;height:80px;overflow:hidden';
        const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(svg)], { type: 'image/svg+xml' }));
        try {
          const image = new Image(); image.src = url; await image.decode();
          const canvas = document.createElement('canvas'); canvas.width = 160; canvas.height = 80; const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
          rasterImages.push(canvas.toDataURL()); rasterSizes.push([image.naturalWidth, image.naturalHeight]);
          return context.getImageData(0, 0, 160, 80).data;
        } finally { URL.revokeObjectURL(url); }
      };
      const first = await raster(0, 'dashed'), clipped = await raster(7, 'dashed'), double = await raster(0, 'double');
      let difference = 0, painted = 0, maskDifference = 0;
      for (let y = 0; y < 80; y++) for (let x = 0; x < 132; x++) {
        const a = first[(y * 160 + x + 28) * 4 + 3]!, b = clipped[(y * 160 + x) * 4 + 3]!;
        difference = Math.max(difference, Math.abs(a - b)); if ((a > 128) !== (b > 128)) maskDifference++;
      }
      for (let x = 0; x < 132; x++) if (clipped[(40 * 160 + x) * 4 + 3]! > 128) painted++;
      const doubleRows = Array.from({ length: 80 }, (_, y) => y).filter(y => double[(y * 160 + 80) * 4 + 3]! > 128);
      const failures: string[] = [], basic = { kind: 'double' as const, color: 'rgba(255,0,0,1)', from: { x: 0, y: 0 }, to: { x: 10, y: 0 } }, bounds = { x: 0, y: 0, width: 10, height: 10 };
      for (const edge of [{ ...basic, kind: 'slantDashDot' as const }, { ...basic, color: 'url(https://example.invalid/workbook-border-injected)' }, { ...basic, to: { x: 1, y: 1 } }]) {
        try { createWorkbookBorderView(document, bounds, [edge]); } catch (error) { failures.push((error as { code: string }).code); }
      }
      try { createWorkbookBorderView(document, bounds, [basic], 2); } catch (error) { failures.push((error as { code: string }).code); }
      return { strokes, difference, maskDifference, painted, doubleRows, failures, rasterImages, rasterSizes };
    },
    textDirectionMeasurement() {
      const scope = document.createElement('div'); scope.dir = 'rtl'; container.append(scope);
      const direction = new WorkbookTextDirection(scope), failures: string[] = [];
      try {
        const values = [' abcאבג', '\u00a0\nאבג abc', '123 אבג', '😀 אבג', '\u200fabc', '\u200eאבג', ' \n', '𞤢', 'ا', '١ا'].map(text => direction.read(text, 'context'));
        const explicit = [direction.read('אבג', 'ltr'), direction.read('abc', 'rtl')];
        const whole = document.createElement('span'); whole.dir = 'auto'; whole.textContent = '123 אבג'; scope.append(whole);
        const wholeStringDirection = getComputedStyle(whole).direction; whole.remove();
        for (let character = 0x4e00; character < 0x4e00 + 300; character++) direction.read(String.fromCodePoint(character), 'context');
        values.push(direction.read('אבג', 'context'));
        const retainedText = scope.querySelector('[data-office-view="workbook-direction-measurement"]')!.textContent;
        try { direction.read('abc', 'invalid' as 'context'); } catch (error) { failures.push((error as { code: string }).code); }
        try { new WorkbookTextDirection(document.createElement('div')); } catch (error) { failures.push((error as { code: string }).code); }
        direction.dispose(); direction.dispose();
        try { direction.read('abc', 'ltr'); } catch (error) { failures.push((error as { code: string }).code); }
        return { values, explicit, wholeStringDirection, retainedText, failures, remaining: scope.childNodes.length, ownerDirection: getComputedStyle(scope).direction };
      } finally { direction.dispose(); scope.remove(); }
    },
    dispose() { loader.dispose(); container.remove(); style.remove(); },
  };
  Object.assign(window, { workbookGridHost: host });
}
void main();
