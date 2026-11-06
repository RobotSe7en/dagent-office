import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createWorkbookNumberFormatter, openDocument } from '@doffi/core';
import type { WorkbookDocumentMetadata } from '@doffi/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { modelWindowSource } from '../../../packages/render/src/workbook/window-source.js';
import { modelAssetSource } from '../../../packages/render/src/workbook/assets.js';
import { modelChartSource } from '../../../packages/render/src/workbook/chart-source.js';
import { WorkbookWindowController } from '../../../packages/render/src/workbook/window-controller.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { layoutWorkbookNoteBoxes } from '../../../packages/render/src/workbook/note-boxes.js';

export async function nativeNoteMeasurement(entries: TestPackage, control?: { vertical: string; rtl: boolean }) {
  const files = decodeTestPackage(entries), decoder = new TextDecoder(), encoder = new TextEncoder();
  if (control) {
    const metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata, sheet = metadata.workbook.sheets[0]!;
    sheet.frozen = { rows: 2, columns: 1 }; sheet.rightToLeft = control.rtl;
    const comment = metadata.workbook.comments!.find(comment => comment.format === 'classic')!;
    if (comment.format !== 'classic' || !comment.box) throw Error('No native note box');
    comment.box.visible = true; comment.box.cssClass = 'controlled-note'; comment.box.anchor = { kind: 'absolute', xPt: 12, yPt: 12, widthPt: 240, heightPt: 180 };
    files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
    const manifest = JSON.parse(decoder.decode(files.get('manifest.json')!)) as { styles: string[]; content: string };
    files.set(manifest.styles[0]!, encoder.encode(decoder.decode(files.get(manifest.styles[0]!)!) + `.controlled-note{--office-note-fill:#ff0000;--office-note-fill-opacity:0.5;--office-note-stroke:#008000;--office-note-stroke-width:2pt;--office-note-stroke-opacity:1;--office-note-stroke-dashstyle:dash;--office-note-shadow:#0000ff;--office-note-shadow-x:4pt;--office-note-shadow-y:4pt;--office-note-shadow-opacity:0.75;--office-note-vertical:${control.vertical};${control.vertical === 'distributed' ? '--office-note-start-arrow:diamond;--office-note-end-arrow:open;' : ''}padding:6pt 8pt}`));
    const template = document.createElement('template'); template.innerHTML = decoder.decode(files.get(manifest.content)!);
    template.content.querySelector('[data-node-id="' + (JSON.parse(decoder.decode(files.get('annotations.json')!)) as {annotationId: string; bodyNodeId: string}[]).find(annotation => annotation.annotationId === comment.annotationId)!.bodyNodeId + '"] p')!.textContent = 'Note A😀\nNote B禰\u{e0100}\nNote C';
    files.set(manifest.content, encoder.encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="' + manifest.styles[0]! + '"></head><body>' + template.content.querySelector('[data-office-kind="workbook-content"]')!.outerHTML + '</body></html>'));
  }
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), before = model.serialize();
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 }), formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
  const controller = new WorkbookWindowController(surface, { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model) }, formatter);
  const content = new WorkbookContentSources(surface.document, snapshot);
  try {
    await controller.render(snapshot.metadata.workbook.activeSheetId, { widthPt: 420, heightPt: 300, scrollLeftPt: 0, scrollTopPt: 0 });
    const current = controller.current!, walker = surface.document.createTreeWalker(current.element, 0xffff_ffff); let actual = 1; while (walker.nextNode()) actual++;
    const notes = current.notes.map(note => {
      const root = note.text.element, area = root.parentElement!, areaBox = area.getBoundingClientRect(), range = surface.document.createRange();
      const projections = note.text.projections.map(part => { range.selectNodeContents(part.node); const rect = range.getBoundingClientRect(); return { text: part.node.data, from: part.from, to: part.to, sourceNodeId: part.sourceNodeId, y: rect.y-areaBox.y, height: rect.height }; });
      return { pane: note.projection.pane, text: root.textContent, exact: projections.every(part => part.to-part.from === part.text.length && model.getNode(part.sourceNodeId)!.text!.slice(part.from,part.to) === part.text), projections, areaHeight: areaBox.height };
    });
    const graphic = current.element.querySelector<SVGSVGElement>('[data-office-view=workbook-note-box] > svg'), paint = graphic ? [...graphic.children].filter(node => ['rect','path'].includes(node.localName)).map(rect => ({ fill: rect.getAttribute('fill'), opacity: rect.getAttribute('fill-opacity') ?? rect.getAttribute('opacity'), stroke: rect.getAttribute('stroke'), dash: rect.getAttribute('stroke-dasharray') })) : [];
    let pixels: number[][] = [];
    if (graphic) {
      const clone = graphic.cloneNode(true) as SVGSVGElement; clone.setAttribute('width', '240'); clone.setAttribute('height', '180');
      const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(clone)], { type: 'image/svg+xml' })), image = new Image();
      try { image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 240; canvas.height = 180; const context = canvas.getContext('2d')!; context.drawImage(image,0,0); pixels = [[1,1],[120,90]].map(([x,y]) => [...context.getImageData(x!,y!,1,1).data]); }
      finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
    }
    const stable = surface.view.innerHTML, errors = [];
    if (notes.length) for (const kind of ['nodes', 'text', 'extent', 'objects', 'cancel']) {
      const abort = new AbortController(); if (kind === 'cancel') abort.abort();
      try { await layoutWorkbookNoteBoxes(surface,snapshot,current.frame,current.dimensions,content,abort.signal,{maxNodes:kind==='nodes'?1:50000,maxSourceNodes:50000,maxCodeUnits:kind==='text'?1:100000,maxExtentPt:kind==='extent'?1:1000000,maxObjects:kind==='objects'?0:5000}); errors.push('accepted'); }
      catch(error) { errors.push((error as {code:string}).code); }
    }
    const after = model.serialize(), unchanged = [...before].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((value,index) => value === after.get(path)![index]));
    const result = { notes, paint, pixels, markers: graphic?.querySelectorAll('marker').length ?? 0, actual, nodes: current.nodes, errors, unchanged, preserved: stable === surface.view.innerHTML, fonts: current.fonts.length, canonical: current.element.querySelectorAll('[data-node-id]').length };
    controller.dispose(); content.dispose(); return { ...result, released: !surface.source.childNodes.length && !surface.view.childNodes.length };
  } finally { controller.dispose(); content.dispose(); formatter.dispose(); surface.dispose(); container.remove(); }
}
