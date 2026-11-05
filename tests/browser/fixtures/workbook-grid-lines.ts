import { AxisDimensions } from '../../../packages/render/src/workbook/dimensions.js';
import { workbookViewport } from '../../../packages/render/src/workbook/viewport.js';
import { WorkbookMergeIndex } from '../../../packages/render/src/workbook/projection.js';
import { projectWorkbookBorders } from '../../../packages/render/src/workbook/border-projection.js';
import { createWorkbookBorderView } from '../../../packages/render/src/workbook/border-view.js';
import { projectWorkbookGridLines, WORKBOOK_GRID_PAINT } from '../../../packages/render/src/workbook/grid-projection.js';

export function gridLineMeasurement(scope: HTMLElement) {
  const document = scope.ownerDocument, gallery = document.createElement('div'); gallery.id = 'workbook-grid-line-probe';
  gallery.style.cssText = 'position:relative;width:480px;height:1200px;background:white';
  const dimensions = { rows: new AxisDimensions(100, 17.3, [{ from: 2, to: 4, sizePt: 0 }, { from: 8, to: 8, sizePt: 29.1 }]),
    columns: new AxisDimensions(50, 23.7, [{ from: 1, to: 1, sizePt: 0 }, { from: 5, to: 6, sizePt: 41.2 }]) };
  const results = [], paint = WORKBOOK_GRID_PAINT, merges = new WorkbookMergeIndex([]);
  for (const rtl of [false, true]) for (const frozen of [{ rows: 0, columns: 0 }, { rows: 2, columns: 3 }]) for (const offset of [0, 9.1]) {
    const layout = workbookViewport(dimensions, frozen, { widthPt: 160, heightPt: 100, scrollLeftPt: offset, scrollTopPt: offset }, 5000);
    const original = projectWorkbookBorders(dimensions, layout, merges, () => ({ before: paint, after: paint, top: paint, bottom: paint }), rtl);
    const full = projectWorkbookGridLines(layout, rtl), index = results.length;
    for (const [column, panes] of [original, full].entries()) {
      const frame = document.createElement('div'); frame.style.cssText = `position:absolute;overflow:hidden;left:${column * 240}px;top:${index * 150}px;width:160pt;height:100pt`;
      for (const pane of panes) frame.append(createWorkbookBorderView(document, pane.bounds, pane.edges));
      gallery.append(frame);
    }
    results.push({ rtl, panes: full.length, original: original.reduce((sum, pane) => sum + pane.edges.length, 0), full: full.reduce((sum, pane) => sum + pane.edges.length, 0),
      endpoints: original.flatMap(pane => pane.edges.flatMap(edge => [edge.from, edge.to])) });
  }
  scope.querySelector('#workbook-grid-line-probe')?.remove(); scope.append(gallery); return results;
}
