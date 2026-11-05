import { workbookCellFill } from '../../../packages/render/src/workbook/cell-fill.js';
import { createWorkbookFillView } from '../../../packages/render/src/workbook/cell-fill-view.js';

async function raster(svg: SVGSVGElement, width: number, height: number): Promise<Uint8ClampedArray> {
  const clone = svg.cloneNode(true) as SVGSVGElement; clone.setAttribute('width', String(width)); clone.setAttribute('height', String(height));
  clone.style.width = width + 'px'; clone.style.height = height + 'px';
  const image = new Image(); image.src = 'data:image/svg+xml,' + encodeURIComponent(new XMLSerializer().serializeToString(clone)); await image.decode();
  const canvas = document.createElement('canvas'); canvas.width = width; canvas.height = height;
  const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0, width, height); return context.getImageData(0, 0, width, height).data;
}
export async function measureWorkbookFills(css: string, classes: readonly string[]) {
  const root = document.createElement('div'), style = document.createElement('style'); style.textContent = css; root.className = 'xlsx-theme'; root.append(style); document.body.append(root);
  const reports = [];
  try {
    for (const [index, cssClass] of classes.entries()) {
      const source = document.createElement('span'); source.className = cssClass; root.append(source); const before = source.outerHTML;
      const paint = workbookCellFill(document, getComputedStyle(source)), view = createWorkbookFillView(document, paint, 72, 54, 'paint-' + index);
      const pixels = await raster(view.element, 96, 72), at = (x: number, y: number) => [...pixels.slice((y * 96 + x) * 4, (y * 96 + x) * 4 + 4)];
      const fragment = createWorkbookFillView(document, paint, 72, 54, 'fragment-' + index, { xPt: 15.75, yPt: 12.75, widthPt: 24, heightPt: 18 });
      const clipped = await raster(fragment.element, 32, 24); let difference = 0;
      for (let y = 0; y < 24; y++) for (let x = 0; x < 32; x++) for (let c = 0; c < 4; c++) difference = Math.max(difference, Math.abs(clipped[(y * 32 + x) * 4 + c]! - pixels[((y + 17) * 96 + x + 21) * 4 + c]!));
      const rows = Array.from({ length: 32 }, (_, y) => Array.from({ length: 32 }, (_, x) => at(x, y).slice(0, 3).join(',') === '32,64,128' ? '1' : '0').join(''));
      const samples = [[12, 9], [48, 9], [84, 9], [12, 36], [48, 36], [84, 36], [12, 63], [48, 63], [84, 63]].map(([x, y]) => ({ x: x!, y: y!, color: at(x!, y!) }));
      const errors = [];
      try { createWorkbookFillView(document, paint, 72, 54, 'low-budget', undefined, view.nodes - 1); errors.push('accepted'); }
      catch (error) { errors.push((error as { code: string }).code); }
      const large = createWorkbookFillView(document, paint, 1e12, 1e12, 'large-cell', { xPt: 1e12 - 30, yPt: 1e12 - 24, widthPt: 24, heightPt: 18 });
      reports.push({ kind: paint.kind, pattern: paint.kind === 'pattern' ? paint.pattern : undefined, nodes: view.nodes, actual: view.element.querySelectorAll('*').length + 1, largeNodes: large.nodes,
        extent: [large.element.style.width, large.element.style.height], sourceUnchanged: source.outerHTML === before, immutable: Object.isFrozen(paint) && (!('stops' in paint) || Object.isFrozen(paint.stops)), rows, samples, difference, errors });
      source.remove();
    }
    const invalid = [];
    for (const declarations of ['--office-cell-fill-kind:unknown', '--office-cell-fill-kind:gradient;--office-cell-gradient-stops:5000', '--office-cell-fill-kind:gradient;--office-cell-gradient-stops:1;--office-cell-gradient-stop-0-position:2;--office-cell-gradient-stop-0-color:red', '--office-cell-fill-kind:gradient;--office-cell-gradient-stops:0;--office-cell-gradient-left:2']) {
      const source = document.createElement('span'); source.style.cssText = declarations; root.append(source);
      try { workbookCellFill(document, getComputedStyle(source)); invalid.push('accepted'); } catch (error) { invalid.push((error as { code: string }).code); }
    }
    return { reports, invalid };
  } finally { root.remove(); }
}
