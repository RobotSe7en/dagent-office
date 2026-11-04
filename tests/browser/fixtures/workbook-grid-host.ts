import { createWorkbook, openDocument } from "@dagent-office/core";
import type { WorkbookDocumentMetadata, WorkbookManifest } from "@dagent-office/core";
import { AxisDimensions } from "../../../packages/render/src/workbook/dimensions.js";
import { WorkbookWindowLoader } from "../../../packages/render/src/workbook/window-loader.js";
import { modelWindowSource } from "../../../packages/render/src/workbook/window-source.js";
import { replaceWorkbookWindow } from "../../../packages/render/src/workbook/window-dom.js";
import { hitWorkbookCell, WorkbookMergeIndex } from "../../../packages/render/src/workbook/projection.js";
import type { WorkbookViewport } from "../../../packages/render/src/workbook/viewport.js";

const encode = (text: string) => new TextEncoder().encode(text);
async function main() {
  const files = new Map(createWorkbook({ documentId: 'grid-geometry' }));
  const manifest = JSON.parse(new TextDecoder().decode(files.get('manifest.json')!)) as WorkbookManifest;
  const metadata = JSON.parse(new TextDecoder().decode(files.get(manifest.metadata)!)) as WorkbookDocumentMetadata;
  const sheet = metadata.workbook.sheets[0]!;
  sheet.frozen = { rows: 2, columns: 1 }; sheet.merges = [{ from: { row: 0, column: 0 }, to: { row: 6, column: 6 } }];
  for (const [row, column, text] of [[0, 0, '合并😀禰󠄀'], [9999, 9, '一万行'], [1_048_575, 16_383, '末格']] as const) {
    const blockRow = Math.floor(row / 256) * 256, blockColumn = Math.floor(column / 32) * 32, path = `sheets/sheet-1/${blockRow}-${blockColumn}.html`;
    const bytes = encode(`<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="${blockRow}" data-column="${blockColumn}"><tbody><tr data-row="${row}"><td data-column="${column}" data-value-type="text">${text}</td></tr></tbody></table></body></html>`);
    const digest = await crypto.subtle.digest('SHA-256', bytes);
    manifest.chunks.push({ sheetId: 'sheet-1', row: blockRow, column: blockColumn, path, sha256: [...new Uint8Array(digest)].map(byte => byte.toString(16).padStart(2, '0')).join('') });
    files.set(path, bytes);
  }
  files.set('manifest.json', encode(JSON.stringify(manifest))); files.set(manifest.metadata, encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: 'workbook' }), before = model.serialize(), loader = new WorkbookWindowLoader(modelWindowSource(model));
  const dimensions = { rows: new AxisDimensions(1_048_576, 10, []), columns: new AxisDimensions(16_384, 20, []) }, merges = new WorkbookMergeIndex(sheet.merges);
  const container = document.createElement('div'); container.id = 'workbook-grid-probe'; container.style.cssText = 'position:absolute;left:0;top:0;background:white'; document.body.replaceChildren(container);
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
    dispose() { loader.dispose(); container.remove(); },
  };
  Object.assign(window, { workbookGridHost: host });
}
void main();
