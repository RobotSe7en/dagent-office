import type { WorkbookSnapshot } from '@dagent-office/core';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookDrawingText } from '../../../packages/render/src/workbook/drawing-text.js';
import { resolvePreviewFontsInScope } from '../../../packages/render/src/fonts.js';

export function measureDrawingText(snapshot: WorkbookSnapshot) {
  const scope = document.createElement('div'), style = document.createElement('style');
  style.textContent = snapshot.styles.map(style => style.css).join('\n'); document.head.append(style); document.body.append(scope);
  const sources = new WorkbookContentSources(document, snapshot), objects = snapshot.metadata.workbook.objects!.filter(object => object.kind === 'geometry' || object.kind === 'text');
  const projection = sources.project(scope, objects.map(object => object.nodeId));
  try {
    const native = objects.map(object => {
      const source = projection.get(object.nodeId)!, before = source.outerHTML, result = createWorkbookDrawingText(source, object);
      scope.append(result.element); const fonts = resolvePreviewFontsInScope(document, result.element, '[data-office-text-style],[data-office-script]');
      const values = { nodeId: object.nodeId, text: result.element.textContent, unchanged: before === source.outerHTML, detachedSource: source.isConnected,
        nodes: result.nodes, actualNodes: result.element.querySelectorAll('*').length + result.projections.length + 1,
        inheritedIds: result.element.querySelectorAll('[data-node-id],[id],[class]').length,
        pieces: result.projections.map(piece => ({ text: piece.node.data, from: piece.from, to: piece.to, sourceNodeId: piece.sourceNodeId, family: getComputedStyle(piece.node.parentElement!).fontFamily, size: getComputedStyle(piece.node.parentElement!).fontSize, role: piece.node.parentElement!.dataset.officeScript })), fonts };
      result.element.remove(); return values;
    });
    const object = objects[0]!, source = projection.get(object.nodeId)!;
    // A source paragraph boundary has no UTF-16 unit. Structural whitespace
    // between paragraphs does occupy source offsets but is not visible text.
    source.innerHTML = '<p data-node-id="drawing-p1" style="font-family:serif;--office-font-ascii:&#34;Liberation Sans&#34;;--office-font-east-asia:&#34;Noto Sans CJK SC&#34;"><span data-node-id="drawing-r1">A中</span><span data-node-id="drawing-r2">󠄀</span><br data-node-id="drawing-br" style="font-size:18pt"><strong data-node-id="drawing-r3">é😀</strong></p>\n <p data-node-id="drawing-p2">Tail</p>';
    const before = source.outerHTML, text = createWorkbookDrawingText(source, object); scope.append(text.element);
    const pieces = text.projections.map(piece => ({ text: piece.node.data, from: piece.from, to: piece.to, sourceNodeId: piece.sourceNodeId, role: piece.node.parentElement!.dataset.officeScript, size: getComputedStyle(piece.node.parentElement!).fontSize, weight: getComputedStyle(piece.node.parentElement!).fontWeight }));
    text.element.remove(); const exact = createWorkbookDrawingText(source, object, text.nodes).nodes;
    const errors: string[] = [];
    for (const budget of [0, 1, text.nodes - 1]) try { createWorkbookDrawingText(source, object, budget); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    source.append(document.createElement('img'));
    try { createWorkbookDrawingText(source, object); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    source.lastElementChild!.remove(); const unchanged = source.outerHTML === before; projection.dispose();
    try { createWorkbookDrawingText(source, object); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    return { native, pieces, exact, nodes: text.nodes, unchanged, errors };
  } finally { sources.dispose(); scope.remove(); style.remove(); }
}
