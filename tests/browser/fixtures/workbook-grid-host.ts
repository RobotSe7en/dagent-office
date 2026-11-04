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
  const container = document.createElement('div'); container.id = 'workbook-grid-probe'; container.style.cssText = 'position:absolute;left:0;top:0;background:white'; document.body.replaceChildren(container);
  const style = document.createElement('style'); style.textContent = model.snapshot().styles.map(style => style.css).join('\n'); document.head.append(style);
  let rightToLeft = false;
  const paint = (body: HTMLElement, cell: { readonly text: string } | undefined) => {
    body.style.border = '1px solid #cbd5e1'; body.style.font = '10pt "Liberation Sans"'; body.style.whiteSpace = 'pre'; body.textContent = cell?.text ?? '';
  };
  const host = {
    async render(viewport: WorkbookViewport, rtl = false) {
      const current = await loader.load('sheet-1', dimensions, viewport);
      if (current) { rightToLeft = rtl; replaceWorkbookWindow(container, loader.frame!, paint, rtl); }
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
      const values: (string | undefined)[] = [], errors: string[] = []; let unchanged = true;
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
        return { values, errors, unchanged };
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
    dispose() { loader.dispose(); container.remove(); style.remove(); },
  };
  Object.assign(window, { workbookGridHost: host });
}
void main();
