import type { WorkbookValue } from "@doffi/core";
import { createWorkbookTextView } from "../../../packages/render/src/workbook/text-view.js";
import { layoutWorkbookTexts } from "../../../packages/render/src/workbook/text-layout.js";

export interface TextLayoutCase {
  text: string; type: WorkbookValue["type"]; family: string; sizePt: number; widthPt: number; heightPt: number;
  horizontal: string; vertical: string; readingOrder: string; indent: number; rightToLeft: boolean;
  wrap?: boolean; shrink?: boolean;
  underline?: 'singleAccounting' | 'doubleAccounting'; inkOnly?: boolean;
  rotation?: number;
  splitOffsets?: readonly number[];
  justifyLastLine?: boolean;
}
export async function textLayoutMeasurement(scope: HTMLElement, cases: readonly TextLayoutCase[], limits: { readonly maxNodes?: number; readonly maxTextUnits?: number } = {}) {
  await document.fonts.ready;
  const sources = document.createElement('div'); sources.style.cssText = 'position:absolute;visibility:hidden;width:0;height:0;overflow:hidden'; scope.append(sources);
  const gallery = document.createElement('div'); gallery.id = 'workbook-text-layout-probe';
  gallery.style.cssText = 'position:absolute;left:0;top:0;display:flex;flex-wrap:wrap;align-items:flex-start;width:850px;gap:4px;background:white;color:black';
  const inputs = cases.map(item => {
    const table = document.createElement('table'), cell = table.insertRow().insertCell(); cell.textContent = item.text; cell.dir = item.rightToLeft ? 'rtl' : 'ltr';
    if (item.splitOffsets) {
      cell.replaceChildren(); let from = 0;
      for (const to of [...item.splitOffsets, item.text.length]) { const span = document.createElement('span'); span.textContent = item.text.slice(from, to); cell.append(span); from = to; }
    }
    if (item.underline) { cell.style.setProperty('--office-font-underline', item.underline); cell.style.color = '#ff00ff'; }
    cell.style.fontFamily = item.family; cell.style.fontSize = `${item.sizePt}pt`; cell.style.fontWeight = 'normal'; cell.style.fontStyle = 'normal'; cell.style.lineHeight = 'normal';
    cell.style.setProperty('--office-cell-horizontal', item.horizontal); cell.style.setProperty('--office-cell-vertical', item.vertical);
    cell.style.setProperty('--office-cell-reading-order', item.readingOrder); cell.style.setProperty('--office-cell-indent', String(item.indent));
    cell.style.setProperty('--office-cell-wrap', String(item.wrap ?? false)); cell.style.setProperty('--office-cell-shrink-to-fit', String(item.shrink ?? false));
    cell.style.setProperty('--office-cell-text-rotation', String(item.rotation ?? 0));
    cell.style.setProperty('--office-cell-justify-last-line', String(item.justifyLastLine ?? false));
    sources.append(table); const source = { element: cell, format: { cssClass: '' } };
    return { source, text: createWorkbookTextView(source), valueType: item.type, widthPt: item.widthPt, heightPt: item.heightPt };
  });
  const original = sources.innerHTML, canvas = document.createElement('canvas'), context = canvas.getContext('2d')!; context.font = '11pt "Liberation Sans"';
  const space = context.measureText(' ').width;
  try {
    const layouts = layoutWorkbookTexts(scope, inputs, space, limits.maxNodes ?? 50_000, limits.maxTextUnits ?? 8 * 1024 * 1024);
    for (const [index, item] of layouts.cells.entries()) { item.element.style.outline = '0.5px solid red'; if (cases[index]!.inkOnly) item.text.element.style.visibility = 'hidden'; gallery.append(item.element); }
    scope.append(gallery);
    const measurements = layouts.cells.map(item => {
      const walker = document.createTreeWalker(item.element, 0xffff_ffff); let actualNodes = 1; while (walker.nextNode()) actualNodes++;
      const cell = item.element.getBoundingClientRect(), range = document.createRange(); range.selectNodeContents(item.text.element); const text = range.getBoundingClientRect();
      const first = item.text.projections[0]?.node, last = item.text.projections.at(-1)?.node;
      const characterX = (node: Text | undefined, end: boolean) => {
        if (!node?.length) return 0;
        const range = document.createRange(), offset = end ? node.length - 1 : 0; range.setStart(node, offset); range.setEnd(node, offset + 1); return (range.getBoundingClientRect().x - cell.x) * .75;
      };
      return { scale: item.scale, requiredHeightPt: item.requiredHeightPt, horizontal: item.text.element.style.textAlign, direction: item.text.element.style.direction,
        lastLine: item.text.element.style.textAlignLast, nodes: item.text.nodes, materializedNodes: item.nodes, actualNodes, codeUnits: item.codeUnits, actualCodeUnits: item.element.textContent?.length ?? 0,
        accounting: [...item.element.querySelectorAll<SVGGElement>('[data-office-view=workbook-accounting-line]')].map(line => ({ style: line.dataset.officeAccountingStyle === 'doubleAccounting' ? 'double' : 'solid', color: getComputedStyle(line.querySelector('line')!).stroke, strokes: line.childElementCount, text: line.textContent })),
        cell: { x: cell.x, y: cell.y, width: cell.width, height: cell.height }, visibleText: item.text.element.textContent,
        lineRects: item.text.projections.flatMap(projection => { const range = document.createRange(); range.selectNode(projection.node); return [...range.getClientRects()].filter(rect => rect.width > .01 && rect.height > 0).map(rect => ({ text: projection.node.data, x: (rect.x - cell.x) * .75, y: (rect.y - cell.y) * .75, width: rect.width * .75, height: rect.height * .75 })); }),
        transform: item.text.element.style.transform, breaks: item.text.element.querySelectorAll('br').length,
        firstX: characterX(first, false), lastX: characterX(last, true),
        gapsPt: [(text.left - cell.left) * .75, (text.top - cell.top) * .75, (cell.right - text.right) * .75, (cell.bottom - text.bottom) * .75],
        widthPt: text.width * .75, heightPt: text.height * .75, text: item.text.projections.map(projection => projection.node.textContent).join(''),
        offsets: item.text.projections.map(projection => [projection.from, projection.to]), font: getComputedStyle(item.text.projections[0]?.node.parentElement ?? item.text.element).fontFamily };
    });
    const failures: string[] = [], source = inputs[0]!.source;
    for (const [name, value] of [['horizontal', 'wrong'], ['vertical', 'wrong'], ['indent', '256'], ['wrap', 'yes'], ['reading-order', 'wrong']]) {
      const property = `--office-cell-${name}`, before = source.element.style.getPropertyValue(property); source.element.style.setProperty(property, value!);
      try { layoutWorkbookTexts(scope, [{ ...inputs[0]!, text: createWorkbookTextView(source) }], space); }
      catch (error) { failures.push((error as { code: string }).code); }
      finally { source.element.style.setProperty(property, before); }
    }
    for (const input of [{ ...inputs[0]!, widthPt: 0 }, { ...inputs[0]!, heightPt: Infinity }]) {
      try { layoutWorkbookTexts(scope, [{ ...input, text: createWorkbookTextView(source) }], space); } catch (error) { failures.push((error as { code: string }).code); }
    }
    try { layoutWorkbookTexts(scope, [{ ...inputs[0]!, text: createWorkbookTextView(source) }], space, 1); } catch (error) { failures.push((error as { code: string }).code); }
    const stackFailures: string[] = [], extra = source.element.cloneNode(false) as HTMLTableCellElement; extra.textContent = 'A'.repeat(1000); sources.append(extra);
    for (const rotation of [181, 254, 255]) {
      extra.style.setProperty('--office-cell-text-rotation', String(rotation)); const candidate = { element: extra, format: { cssClass: '' } }, text = createWorkbookTextView(candidate);
      try { layoutWorkbookTexts(scope, [{ source: candidate, text, valueType: 'text', widthPt: 50, heightPt: 50 }], space, 20); }
      catch (error) { stackFailures.push((error as { code: string }).code); }
      if (text.element.parentNode) throw new Error('Failed stacked text candidate remained mounted');
    }
    extra.remove();
    // A later invalid item must clean up earlier private views without moving
    // the accepted gallery or retaining any source text in measurement DOM.
    const valid = { ...inputs[0]!, text: createWorkbookTextView(source) }, badSource = inputs.at(-1)!.source;
    const old = badSource.element.style.getPropertyValue('--office-cell-horizontal'); badSource.element.style.setProperty('--office-cell-horizontal', 'bad');
    try { layoutWorkbookTexts(scope, [valid, { ...inputs.at(-1)!, text: createWorkbookTextView(badSource) }], space); }
    catch (error) { failures.push((error as { code: string }).code); }
    finally { badSource.element.style.setProperty('--office-cell-horizontal', old); }
    return { measurements, spacePt: space * .75, fonts: layouts.fonts, usedFallbackLineMetrics: layouts.usedFallbackLineMetrics, unchanged: original === sources.innerHTML,
      retained: scope.querySelectorAll('[data-office-view="workbook-text-measurement"],[data-office-view="workbook-direction-measurement"]').length,
      failures, stackFailures, atomic: gallery.isConnected && !valid.text.element.parentNode, frozen: Object.isFrozen(layouts) && Object.isFrozen(layouts.cells) };
  } finally { sources.remove(); }
}
