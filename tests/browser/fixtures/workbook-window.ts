import { createWorkbookNumberFormatter, openDocument } from '@dagent-office/core';
import type { WorkbookDocumentMetadata } from '@dagent-office/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { measureWorkbookColumnFont } from '../../../packages/render/src/workbook/column-font.js';
import { worksheetDimensions } from '../../../packages/render/src/workbook/dimensions.js';
import { WorkbookWindowLoader } from '../../../packages/render/src/workbook/window-loader.js';
import { modelWindowSource } from '../../../packages/render/src/workbook/window-source.js';
import { layoutWorkbookWindow } from '../../../packages/render/src/workbook/window-layout.js';
import type { WorkbookViewport } from '../../../packages/render/src/workbook/viewport.js';

export async function windowLayoutMeasurement(entries: readonly [string, number[]][], views: readonly { sheetId: string; viewport: WorkbookViewport }[], rtl = false) {
  const files = new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)]));
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json'))) as WorkbookDocumentMetadata;
  for (const sheet of metadata.workbook.sheets) sheet.rightToLeft = rtl;
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot();
  const originalSource = model.serialize();
  const formatter = createWorkbookNumberFormatter({ dateSystem: metadata.workbook.dateSystem, hostLocale: 'zh-CN' });
  const container = document.createElement('div'); container.style.cssText = 'position:relative;width:1000px;height:800px'; document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 750, heightPt: 600 });
  const loader = new WorkbookWindowLoader(modelWindowSource(model)), results = [];
  try {
    for (const { sheetId, viewport } of views) {
      const normal = measureWorkbookColumnFont(surface.source, snapshot.metadata.workbook, sheetId);
      const dimensions = worksheetDimensions(snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!, normal.maximumDigitWidthPx);
      await loader.load(sheetId, dimensions, viewport);
      const rendered = layoutWorkbookWindow(surface, snapshot, loader.frame!, dimensions, formatter, { normalSpaceWidthPx: normal.spaceWidthPx });
      const before = surface.view.firstChild; const detached = !rendered.element.isConnected && surface.view.firstChild === before;
      surface.view.replaceChildren(rendered.element); surface.frame.style.setProperty('visibility', 'visible', 'important');
      const view = surface.document.defaultView!;
      const cells = rendered.cells.filter(cell => cell.text).map(cell => {
        const text = cell.text!, bounds = text.text.element.getBoundingClientRect();
        return { row: cell.projection.row, column: cell.projection.column, pane: cell.projection.pane, text: text.text.element.textContent,
          rect: cell.projection.rect, clipRect: cell.projection.clip, offsets: text.text.projections.map(part => [part.from, part.to, part.node.data]), shape: text.shape, scale: text.scale,
          runs: text.text.projections.map(part => { const style = view.getComputedStyle(part.node.parentElement!); return { from: part.from, to: part.to, text: part.node.data, color: style.color, weight: style.fontWeight, style: style.fontStyle, family: style.fontFamily }; }),
          color: view.getComputedStyle(text.text.element).color, weight: view.getComputedStyle(text.text.element).fontWeight, family: view.getComputedStyle(text.text.element).fontFamily,
          breaks: text.text.element.querySelectorAll("br").length, spaces: [...text.text.element.querySelectorAll<HTMLElement>("[data-office-number-space]")].map(node => node.getBoundingClientRect().width), bounds: { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height },
          clip: text.element.parentElement!.parentElement!.style.cssText };
      });
      const errors: string[] = [];
      for (const limits of [{ maxNodes: 5 }, { maxSourceNodes: 1 }, { maxTextCodeUnits: 1 }]) {
        try { layoutWorkbookWindow(surface, snapshot, loader.frame!, dimensions, formatter, { normalSpaceWidthPx: normal.spaceWidthPx, ...limits }); errors.push('accepted'); }
        catch (error) { errors.push((error as { code: string }).code); }
      }
      const same = rendered.element === surface.view.firstChild;
      const fillNodes = [...rendered.element.querySelector('[data-office-view="workbook-fills"]')!.children];
      results.push({ sheetId, cells, detached, same, errors, fonts: rendered.fonts.length,
        fillPaints: frameFills(),
        borderColors: [...rendered.element.querySelectorAll('line')].filter(line => !line.closest('[data-office-view="workbook-gridlines"]')).map(line => line.getAttribute('stroke')),
        gridLines: rendered.element.querySelectorAll('[data-office-view="workbook-gridlines"]').length,
        widthPt: rendered.element.getBoundingClientRect().width * .75, heightPt: rendered.element.getBoundingClientRect().height * .75,
        accessible: rendered.element.querySelectorAll('[role="gridcell"]').length,
        fills: [...rendered.element.querySelectorAll<SVGSVGElement>('[data-office-view="workbook-cell-fill"]')].map(fill => ({ viewBox: fill.getAttribute('viewBox'), transform: fill.querySelector('g')?.getAttribute('transform'), width: fill.clientWidth })),
        borderCount: rendered.element.querySelectorAll('line').length,
        retainedSources: surface.source.querySelectorAll('[data-office-view="workbook-source-measurement"]').length,
        retainedMeasurements: surface.view.querySelectorAll('[data-office-view$="measurement"]').length,
      });
      function frameFills() {
        return loader.frame!.projections.map((cell, index) => ({ row: cell.row, column: cell.column, pane: cell.pane, color: fillNodes[index]!.querySelector('rect')?.getAttribute('fill') ?? null }));
      }
    }
    const after = model.serialize(), unchanged = [...originalSource].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index]));
    return { results, unchanged };
  } finally { loader.dispose(); formatter.dispose(); surface.dispose(); container.remove(); }
}
