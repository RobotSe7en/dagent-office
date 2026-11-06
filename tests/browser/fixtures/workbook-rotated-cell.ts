import { createWorkbookTextView } from "../../../packages/render/src/workbook/text-view.js";
import { layoutWorkbookTexts } from "../../../packages/render/src/workbook/text-layout.js";
import { createWorkbookShapedFillView } from "../../../packages/render/src/workbook/shaped-cell-fill.js";
import { createWorkbookBorderView } from "../../../packages/render/src/workbook/border-view.js";
import { workbookCellShape, workbookCellShapeOffset } from "../../../packages/render/src/workbook/cell-shape.js";
import type { WorkbookCellFill } from "../../../packages/render/src/workbook/cell-fill.js";
import { rasterWorkbookFill } from "./workbook-fill.js";

export async function shapedFillMeasurement() {
  const paints: WorkbookCellFill[] = [
    { kind: 'solid', color: '#00ffff' },
    { kind: 'pattern', pattern: 'darkGrid', foreground: '#204080', background: '#ffe0a0' },
    { kind: 'linear', degree: 45, stops: [{ position: 0, color: '#f00' }, { position: 1, color: '#00f' }] },
    { kind: 'path', left: .2, right: .8, top: .3, bottom: .7, stops: [{ position: 0, color: '#f00' }, { position: 1, color: '#00f' }] },
  ], reports = [];
  for (const angle of [-45, 45]) for (const [index, paint] of paints.entries()) {
    const shape = workbookCellShape(72, 54, angle, true), left = Math.min(0, shape.topOffsetPt);
    const whole = createWorkbookShapedFillView(document, paint, shape, 'shaped-full-' + index, { xPt: left, yPt: 0, widthPt: 126, heightPt: 54 });
    const clip = { xPt: left + 18, yPt: 12, widthPt: 54, heightPt: 30 }, fragment = createWorkbookShapedFillView(document, paint, shape, 'shaped-clip-' + index, clip);
    const full = await rasterWorkbookFill(whole.element, 168, 72), pixels = await rasterWorkbookFill(fragment.element, 72, 40); let difference = 0;
    for (let y = 0; y < 40; y++) for (let x = 0; x < 72; x++) for (let channel = 0; channel < 4; channel++) {
      difference = Math.max(difference, Math.abs(pixels[(y * 72 + x) * 4 + channel]! - full[((y + 16) * 168 + x + 24) * 4 + channel]!));
    }
    const huge = workbookCellShape(1e12, 1e12, angle, true);
    const bounded = createWorkbookShapedFillView(document, paint, huge, 'shaped-large-' + index, { xPt: 12, yPt: 1e12 - 42, widthPt: 54, heightPt: 30 });
    const failures = [];
    try { createWorkbookShapedFillView(document, paint, shape, 'shaped-low-budget', clip, fragment.nodes - 1); }
    catch (error) { failures.push((error as { code: string }).code); }
    reports.push({ angle, kind: paint.kind, difference, nodes: fragment.nodes, hugeNodes: bounded.nodes, extent: [bounded.element.style.width, bounded.element.style.height], failures });
  }
  return reports;
}

export interface RotatedCellCase { rotation: number; mode: string; horizontal: string; vertical: string; text: string; cellSizePt: readonly number[] }
export async function rotatedCellMeasurement(scope: HTMLElement, cases: readonly RotatedCellCase[]) {
  await document.fonts.ready;
  const sources = document.createElement('div'), gallery = document.createElement('div');
  sources.style.cssText = 'position:absolute;visibility:hidden;width:0;height:0;overflow:hidden'; scope.append(sources);
  gallery.id = 'workbook-rotated-cell-probe'; gallery.style.cssText = 'position:absolute;left:0;top:0;display:flex;flex-direction:column;gap:4px;background:white';
  const inputs = cases.map(item => {
    const table = document.createElement('table'), cell = table.insertRow().insertCell(); cell.textContent = item.text;
    cell.style.cssText = 'font-family:"Noto Sans CJK SC";font-size:12pt;line-height:normal;color:black;background:#00ffff';
    cell.style.setProperty('--office-cell-text-rotation', String(item.rotation));
    cell.style.setProperty('--office-cell-horizontal', item.horizontal); cell.style.setProperty('--office-cell-vertical', item.vertical);
    for (const side of item.mode === 'frame' ? ['top', 'bottom', 'left', 'right'] : ['top', 'left', 'diagonal'].includes(item.mode) ? [item.mode] : []) {
      cell.style.setProperty(`--office-cell-border-${side}-kind`, 'thin'); cell.style.setProperty(`--office-cell-border-${side}-color`, '#cc00cc');
    }
    sources.append(table); const source = { element: cell, format: { cssClass: '' } };
    return { source, text: createWorkbookTextView(source), valueType: 'text' as const, widthPt: item.cellSizePt[0]!, heightPt: item.cellSizePt[1]! };
  });
  const original = sources.innerHTML;
  try {
    const result = layoutWorkbookTexts(scope, inputs, 4);
    const measurements = result.cells.map((cell, index) => {
      const shape = cell.shape, item = cases[index]!, outer = document.createElement('div');
      outer.style.cssText = `position:relative;overflow:hidden;width:300pt;height:${shape.heightPt}pt;background:white`;
      const x = Math.max(-100, Math.min(0, shape.topOffsetPt)), right = Math.min(200, shape.widthPt + Math.max(0, shape.topOffsetPt));
      const clip = { xPt: x, yPt: 0, widthPt: right - x, heightPt: shape.heightPt };
      const fill = createWorkbookShapedFillView(document, { kind: 'solid', color: '#00ffff' }, shape, 'rotated-fill-' + index, clip);
      fill.element.style.left = `${x + 100}pt`; outer.append(fill.element);
      const point = (x: number, y: number) => ({ x: x + 100 + workbookCellShapeOffset(shape, y), y });
      const a = point(0, 0), b = point(shape.widthPt, 0), c = point(shape.widthPt, shape.heightPt), d = point(0, shape.heightPt);
      const segments = item.mode === 'frame' ? [[a, b], [b, c], [d, c], [a, d]] : item.mode === 'top' ? [[a, b]] : item.mode === 'left' ? [[a, d]] : item.mode === 'diagonal' ? [[d, b]] : [];
      outer.append(createWorkbookBorderView(document, { x: 0, y: 0, width: 300, height: shape.heightPt }, segments.map(([from, to]) => ({ kind: 'thin', color: 'rgb(204,0,204)', from: from!, to: to! }))));
      cell.element.style.position = 'absolute'; cell.element.style.left = '100pt'; outer.append(cell.element); gallery.append(outer);
      return { shape, points: [a, b, c, d].map(point => [point.x - 100, point.y]), projections: cell.text.projections.map(part => [part.from, part.to]), nodes: fill.nodes };
    });
    scope.append(gallery);
    return { measurements, unchanged: original === sources.innerHTML, retained: scope.querySelectorAll('[data-office-view="workbook-text-measurement"]').length };
  } finally { sources.remove(); }
}
