import { workbookCellDiagonals } from '../../../packages/render/src/workbook/cell-borders.js';
import { createWorkbookBorderView } from '../../../packages/render/src/workbook/border-view.js';
import { projectWorkbookDiagonals } from '../../../packages/render/src/workbook/diagonal-projection.js';
import type { WorkbookViewportLayout } from '../../../packages/render/src/workbook/viewport.js';

async function raster(svg: SVGSVGElement, width: number, height: number) {
  svg.setAttribute('width', String(width)); svg.setAttribute('height', String(height)); svg.style.cssText = `width:${width}px;height:${height}px`;
  const image = new Image(); image.src = 'data:image/svg+xml,' + encodeURIComponent(new XMLSerializer().serializeToString(svg)); await image.decode();
  const canvas = document.createElement('canvas'); canvas.width = width; canvas.height = height; const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
  return context.getImageData(0, 0, width, height).data;
}
export async function diagonalMeasurement(css: string, classes: readonly string[]) {
  const root = document.createElement('div'), style = document.createElement('style'); root.className = 'xlsx-theme'; style.textContent = css; root.append(style); document.body.append(root);
  const result = [];
  try {
    for (const cssClass of classes) {
      const source = document.createElement('span'); source.className = cssClass; root.append(source); const before = source.outerHTML;
      const paint = workbookCellDiagonals(document, getComputedStyle(source));
      const project = (x: number, y: number, width: number, height: number) => {
        const layout: WorkbookViewportLayout = { widthPt: 72, heightPt: 54, scrollLeftPt: 0, scrollTopPt: 0, maxScrollLeftPt: 0, maxScrollTopPt: 0, cellCount: 1,
          panes: [{ kind: 'body', inlineStartPt: x, blockStartPt: y, widthPt: width, heightPt: height, sourceLeftPt: x, sourceTopPt: y, rows: [], columns: [] }] };
        const [pane] = projectWorkbookDiagonals([{ pane: 'body', row: 1, column: 1, range: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } }, rect: { x: 0, y: 0, width: 72, height: 54 }, clip: { x, y, width, height } }], layout, () => paint, false);
        return createWorkbookBorderView(document, pane!.bounds, pane!.edges);
      };
      const svg = project(0, 0, 72, 54), strokes = [...svg.querySelectorAll('line')].map(line => ({ from: [Number(line.getAttribute('x1')), Number(line.getAttribute('y1'))], to: [Number(line.getAttribute('x2')), Number(line.getAttribute('y2'))], width: Number(line.getAttribute('stroke-width')), color: line.getAttribute('stroke') }));
      const full = await raster(svg, 96, 72), part = await raster(project(15.75, 12.75, 24, 18), 32, 24);
      let difference = 0, maskDifference = 0;
      for (let y = 0; y < 24; y++) for (let x = 0; x < 32; x++) {
        const a = full[((y + 17) * 96 + x + 21) * 4 + 3]!, b = part[(y * 32 + x) * 4 + 3]!;
        difference = Math.max(difference, Math.abs(a - b)); if ((a > 128) !== (b > 128)) maskDifference++;
      }
      result.push({ strokes, difference, maskDifference, unchanged: source.outerHTML === before }); source.remove();
    }
    const sample = createWorkbookBorderView(document, { x: 0, y: 0, width: 81, height: 1.5 }, [{ kind: 'slantDashDot', color: 'rgba(0,0,0,1)', from: { x: 0, y: .75 }, to: { x: 81, y: .75 } }]);
    const pixels = await raster(sample, 108, 2), rows = [0, 1].map(y => Array.from({ length: 108 }, (_, x) => pixels[(y * 108 + x) * 4 + 3]! > 128 ? '1' : '0').join(''));
    const failures = [];
    for (const css of ['--office-cell-border-diagonal-up:yes', '--office-cell-border-diagonal-down:1', '--office-cell-border-diagonal-kind:constructor', '--office-cell-border-diagonal-up:true;--office-cell-border-diagonal-kind:thin;--office-cell-border-diagonal-color:url(https://example.invalid/diagonal)']) {
      const source = document.createElement('span'); source.style.cssText = css; root.append(source);
      try { workbookCellDiagonals(document, getComputedStyle(source)); failures.push('accepted'); } catch (error) { failures.push((error as { code: string }).code); }
    }
    return { result, rows, failures };
  } finally { root.remove(); }
}
