import { createWorkbook, createWorkbookNumberFormatter, openDocument } from '@dagent-office/core';
import type { Annotation, WorkbookDocumentMetadata } from '@dagent-office/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookWindowController } from '../../../packages/render/src/workbook/window-controller.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { modelWindowSource } from '../../../packages/render/src/workbook/window-source.js';
import { modelAssetSource } from '../../../packages/render/src/workbook/assets.js';
import { modelChartSource } from '../../../packages/render/src/workbook/chart-source.js';
import { createWorkbookAnnotationView } from '../../../packages/render/src/workbook/annotation-view.js';
import type { WorkbookAnnotationView } from '../../../packages/render/src/workbook/annotation-view.js';
import { createWorkbookCommentView } from '../../../packages/render/src/workbook/comment-view.js';
import type { WorkbookCommentView } from '../../../packages/render/src/workbook/comment-view.js';

let dispose = () => {}, inspect = () => ({});
let focus = () => {}, fail = () => {}, clear = () => {};
export function annotationControlsAction(action: 'inspect' | 'focus' | 'fail' | 'clear' | 'dispose') {
  if (action === 'focus') focus(); if (action === 'fail') fail(); if (action === 'clear') clear(); if (action === 'dispose') dispose();
  return inspect();
}
export async function mountAnnotationControls(rtl: boolean) {
  dispose();
  const files = new Map(createWorkbook({ documentId: 'annotation-controls' })), encoder = new TextEncoder();
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.rightToLeft = rtl; metadata.workbook.sheets[0]!.frozen = { rows: 2, columns: 1 };
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  const range = (row: number, column: number, endRow = row, endColumn = column) => ({ from: { row, column }, to: { row: endRow, column: endColumn } });
  const ranges = [range(0, 0, 6, 6), range(9, 6)];
  const annotation = (id: string, regions = ranges, status: Annotation['status'] = 'open'): Annotation => ({ annotationId: id, documentId: 'annotation-controls', baseRevision: 0,
    status, anchor: { kind: 'cells', sheetId: 'sheet-1', ranges: regions }, instruction: '批注正文 ' + id } as Annotation);
  files.set('annotations.json', encoder.encode(JSON.stringify([annotation('first'), annotation('same'), annotation('separate', [range(8, 5)]), annotation('recover', ranges, 'needs_reanchor')])));
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), initial = model.serialize();
  const host = document.createElement('div'); host.style.cssText = 'position:fixed;left:0;top:0;width:600px;height:440px;z-index:2147483647;background:white'; document.body.append(host);
  const surface = await createWorkbookSurface(host, snapshot, { widthPt: 420, heightPt: 300 }), sources = new WorkbookContentSources(surface.document, snapshot);
  const formatter = createWorkbookNumberFormatter({ dateSystem: '1900' }), controller = new WorkbookWindowController(surface, { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model) }, formatter);
  let controls: WorkbookAnnotationView | undefined, panel: WorkbookCommentView | undefined, error = '', failNext = false;
  const calls: string[][] = [];
  try {
    await controller.render('sheet-1', { widthPt: 420, heightPt: 300, scrollLeftPt: 20, scrollTopPt: 20 });
    const window = controller.current!;
    controls = createWorkbookAnnotationView(surface.document, window.annotations, window.frame.layout, rtl, ids => {
      calls.push([...ids]);
      try {
        const next = createWorkbookCommentView(surface, snapshot, sources, ids[0]!, failNext ? { maxNodes: 1 } : {});
        next.element.style.cssText += ';position:absolute;left:180pt;top:40pt;width:220pt;max-height:240pt;overflow:auto';
        controls!.select(ids[0]); window.element.append(next.element); panel?.dispose(); panel = next;
      } catch (reason) { error = (reason as { code: string }).code; }
      finally { failNext = false; }
    });
    window.element.append(controls.element);
    const bounded = createWorkbookAnnotationView(surface.document, window.annotations, window.frame.layout, rtl, () => {}, controls.nodes + 4);
    bounded.select('separate'); const preserved = bounded.element.innerHTML; let budget = '';
    try { bounded.select('first'); budget = 'accepted'; } catch (reason) { budget = (reason as { code: string }).code; }
    const selectionPreserved = bounded.element.innerHTML === preserved; bounded.dispose();
    const doc = surface.document, globals = doc.defaultView as unknown as Record<string, unknown>, violations: string[] = [];
    globals.__workbookUnsafeSource = false; let trustedEvents = 0;
    const violation = (event: SecurityPolicyViolationEvent) => { violations.push(event.effectiveDirective); }; doc.addEventListener('securitypolicyviolation', violation);
    const script = doc.createElement('script'); script.textContent = 'window.__workbookUnsafeSource=true'; doc.body.append(script);
    const unsafe = doc.createElement('button'); unsafe.setAttribute('onclick', 'window.__workbookUnsafeSource=true');
    unsafe.addEventListener('click', () => { trustedEvents++; }); doc.body.append(unsafe); unsafe.click();
    await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    const sourceBlocked = globals.__workbookUnsafeSource === false; delete globals.__workbookUnsafeSource;
    doc.removeEventListener('securitypolicyviolation', violation); script.remove(); unsafe.remove();
    const first = controls.groups[0]!, frame = surface.frame.getBoundingClientRect(), button = first.button.getBoundingClientRect();
    const initialGroups = controls.groups.map(group => [...group.annotationIds]);
    focus = () => first.button.focus(); fail = () => { failNext = true; }; clear = () => { panel?.dispose(); panel = undefined; controls!.select(undefined); };
    inspect = () => {
      const highlights = [...controls!.element.querySelectorAll<HTMLElement>('[data-office-view="workbook-annotation-highlights"] > div')].map(node => ({ x: parseFloat(node.style.left), y: parseFloat(node.style.top), width: parseFloat(node.style.width), height: parseFloat(node.style.height),
        top: node.style.borderTopWidth, left: node.style.borderLeftWidth, right: node.style.borderRightWidth, bottom: node.style.borderBottomWidth }));
      const actual = surface.document.createTreeWalker(controls!.element, 0xffff_ffff); let actualNodes = 1; while (actual.nextNode()) actualNodes++;
      const after = model.serialize();
      return { calls: calls.map(ids => [...ids]), groups: initialGroups, error, panel: panel?.element.textContent, highlights, nodes: controls!.nodes, actualNodes,
        sourceClean: !surface.source.childNodes.length, pressed: controls!.groups.map(group => group.button.getAttribute('aria-pressed')),
        unchanged: [...initial].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index])) };
    };
    const cleanup = () => { panel?.dispose(); controls?.dispose(); sources.dispose(); controller.dispose(); formatter.dispose(); surface.dispose(); host.remove(); };
    dispose = () => {
      const before = calls.length; cleanup(); first.button.click(); const eventsReleased = before === calls.length;
      inspect = () => ({ eventsReleased, removed: !host.isConnected, sourceClean: !surface.source.childNodes.length });
      dispose = () => {}; focus = fail = clear = () => {};
    };
    return { x: frame.x + button.x + button.width / 2, y: frame.y + button.y + button.height / 2, groups: initialGroups, rtl,
      budget, selectionPreserved, sourceBlocked, trustedEvents, violations };
  } catch (error) { panel?.dispose(); controls?.dispose(); sources.dispose(); controller.dispose(); formatter.dispose(); surface.dispose(); host.remove(); throw error; }
}
