import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createWorkbookCommentView } from '../../../packages/render/src/workbook/comment-view.js';
import type { WorkbookCommentView } from '../../../packages/render/src/workbook/comment-view.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { openDocument } from '@dagent-office/core';
import type { Annotation, WorkbookDocumentMetadata } from '@dagent-office/core';

export async function mentionPanelMeasurement(entries: TestPackage, checkBoxes = false) {
  const model = openDocument(decodeTestPackage(entries), { kind: 'workbook' }), snapshot = model.snapshot(), before = model.serialize();
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 }), sources = new WorkbookContentSources(surface.document, snapshot);
  const results = [];
  try {
    for (const annotation of snapshot.annotations) {
      const view = createWorkbookCommentView(surface, snapshot, sources, annotation.annotationId);
      try {
        surface.view.append(view.element);
        const native = snapshot.metadata.workbook.comments?.find(comment => comment.annotationId === annotation.annotationId);
        const walker = surface.document.createTreeWalker(view.element, 0xffff_ffff); let actualNodes = 1; while (walker.nextNode()) actualNodes++;
        const errors = [], active = surface.view.innerHTML;
        if (checkBoxes) for (const limits of [{ maxNodes: view.nodes - 1 }, { maxSourceNodes: 1 }, { maxCodeUnits: 1 }]) {
          try { createWorkbookCommentView(surface, snapshot, sources, annotation.annotationId, limits).dispose(); errors.push('accepted'); }
          catch (error) { errors.push((error as { code: string }).code); }
          if (surface.view.innerHTML !== active) throw Error('A failed note box candidate changed the active panel');
        }
        results.push({ status: annotation.status, text: view.element.textContent, nodes: view.nodes, actualNodes, errors,
          box: native?.format === 'classic' ? { source: native.boxSource, selected: native.box, alternate: native.commentPrBox, differences: native.boxDifferences } : undefined,
          boxWarning: view.element.querySelector('[data-office-view="workbook-comment-box-differences"]')?.textContent, labels: [...view.element.querySelectorAll('[data-office-view="workbook-comment-mention"]')].map(node => node.textContent),
          interactive: view.element.querySelectorAll('a,button,input').length,
          exact: view.bodies.every(({ body }) => body.projections.every(part => view.element.contains(part.node) && part.node.data === model.getNode(part.sourceNodeId)!.text.slice(part.from, part.to))) });
      } finally { view.dispose(); }
    }
    const after = model.serialize(); sources.dispose();
    return { results, unchanged: [...before].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index])),
      released: !surface.source.childNodes.length && !surface.view.childNodes.length };
  } finally { sources.dispose(); surface.dispose(); container.remove(); }
}

export async function commentPanelMeasurement(entries: TestPackage, recovered: boolean) {
  const files = decodeTestPackage(entries), decoder = new TextDecoder(), encoder = new TextEncoder();
  const annotations = JSON.parse(decoder.decode(files.get('annotations.json')!)) as Annotation[], metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  const first = annotations[0]!, native = metadata.workbook.comments!.find(comment => comment.annotationId === first.annotationId)!;
  if (recovered) {
    first.status = 'needs_reanchor'; native.recoveryStatus = 'resolved'; native.recoveryReasons = ['mismatched-placeholder'];
    if (native.format !== 'modern') throw Error('Recovery fixture requires a modern thread');
    first.replies!.push({ replyId: 'review-reply', author: { name: 'Second author' }, instruction: '回复中的纯文本', status: 'open' });
    native.replies.push({ replyId: 'review-reply', parentReplyId: first.replies![0]!.replyId });
  }
  annotations.push({ annotationId: 'host-note', documentId: first.documentId, baseRevision: first.baseRevision, anchor: first.anchor, status: 'closed',
    author: { name: '<script>author</script>' }, instruction: '<img src="https://invalid.example/private">仅作为文字' });
  files.set('annotations.json', encoder.encode(JSON.stringify(annotations))); files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), initial = model.serialize();
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 }), sources = new WorkbookContentSources(surface.document, snapshot);
  let active: WorkbookCommentView | undefined;
  try {
    const candidate = createWorkbookCommentView(surface, snapshot, sources, first.annotationId), detached = !candidate.element.isConnected;
    active = candidate; surface.view.append(candidate.element); const stable = surface.view.innerHTML;
    const walker = surface.document.createTreeWalker(candidate.element, 0xffff_ffff); let actualNodes = 1; while (walker.nextNode()) actualNodes++;
    const results = { detached, nodes: candidate.nodes, actualNodes, text: candidate.element.textContent,
      warning: candidate.element.querySelector('[data-office-view="workbook-comment-recovery"]')?.textContent,
      bodies: candidate.bodies.map(item => ({ replyId: item.replyId, text: item.body.element.textContent, projections: item.body.projections.length })),
      replyParents: [...candidate.element.querySelectorAll<HTMLElement>('[data-parent-reply-id]')].map(node => [node.dataset.replyId, node.dataset.parentReplyId]),
      fonts: candidate.fonts.length, editable: candidate.element.querySelectorAll('input,textarea,select,[contenteditable="true"]').length };
    const errors = [];
    for (const limits of [{ maxNodes: candidate.nodes - 1 }, { maxSourceNodes: 1 }, { maxCodeUnits: 1 }, ...(recovered ? [{ maxReplies: 1 }] : [])]) {
      try { createWorkbookCommentView(surface, snapshot, sources, first.annotationId, limits).dispose(); errors.push('accepted'); }
      catch (error) { errors.push((error as { code: string }).code); }
      if (stable !== surface.view.innerHTML) throw Error('A failed comment candidate changed the published panel');
    }
    for (const target of [{ snapshot: { ...snapshot, manifest: { ...snapshot.manifest, revision: snapshot.manifest.revision + 1 } }, id: first.annotationId }, { snapshot, id: 'missing-comment' }]) {
      try { createWorkbookCommentView(surface, target.snapshot, sources, target.id).dispose(); errors.push('accepted'); }
      catch (error) { errors.push((error as { code: string }).code); }
    }
    const next = createWorkbookCommentView(surface, snapshot, sources, 'host-note'); surface.view.replaceChildren(next.element); active.dispose(); active = next;
    const plain = { text: next.element.textContent, elements: next.element.querySelectorAll('img,script').length, bodies: next.bodies.length };
    const released = !candidate.element.childNodes.length && !candidate.element.isConnected;
    active.dispose(); active.dispose(); const after = model.serialize();
    return { ...results, errors, plain, released, empty: !surface.view.childNodes.length && !surface.source.childNodes.length,
      unchanged: [...initial].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index])) };
  } finally { active?.dispose(); sources.dispose(); surface.dispose(); container.remove(); }
}
