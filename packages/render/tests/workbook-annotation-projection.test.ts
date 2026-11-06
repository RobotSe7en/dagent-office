import { describe, expect, it, vi } from "vitest";
import { createWorkbook, openDocument } from "@doffi/core";
import type { Annotation, WorkbookComment, WorkbookDocumentMetadata, WorkbookRange } from "@doffi/core";
import { AxisDimensions } from "../src/workbook/dimensions.js";
import { projectWorkbookAnnotations } from "../src/workbook/annotation-projection.js";
import { workbookViewport } from "../src/workbook/viewport.js";
import type { WorkbookWindowFrame } from "../src/workbook/window-loader.js";

const range = (row: number, column: number, endRow = row, endColumn = column): WorkbookRange => ({ from: { row, column }, to: { row: endRow, column: endColumn } });
const annotation = (id: string, ranges: WorkbookRange[], status: Annotation['status'] = 'open', sheetId = 'sheet-1'): Annotation =>
  ({ annotationId: id, documentId: 'annotations', baseRevision: 0, status, anchor: { kind: 'cells', sheetId, ranges }, instruction: '原始批注 ' + id } as Annotation);
function snapshot(annotations: Annotation[], merges: WorkbookRange[] = [], comments: WorkbookComment[] = []) {
  const files = new Map(createWorkbook({ documentId: 'annotations' })), encoder = new TextEncoder();
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.merges = merges; metadata.workbook.comments = comments;
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata))); files.set('annotations.json', encoder.encode(JSON.stringify(annotations)));
  return openDocument(files, { kind: 'workbook' }).snapshot();
}
const dimensions = { rows: new AxisDimensions(1_048_576, 10, []), columns: new AxisDimensions(16_384, 20, []) };
function frame(input = dimensions, frozen = { rows: 2, columns: 1 }, scroll = { left: 40, top: 30 }, size = { width: 100, height: 80 }): WorkbookWindowFrame {
  return { documentId: 'annotations', revision: 0, sheetId: 'sheet-1',
    layout: workbookViewport(input, frozen, { widthPt: size.width, heightPt: size.height, scrollLeftPt: scroll.left, scrollTopPt: scroll.top }, 100),
    projections: [], sourceCoordinates: [], completeRows: [], skewColumns: [], storedCellCount: 0, getCell: () => undefined };
}

describe('workbook annotation windows', () => {
  it('projects a covered merged-cell note in four panes without rewriting the original cell anchor', () => {
    const source = snapshot([annotation('note', [range(3, 2)])], [range(0, 0, 6, 6)], [{ annotationId: 'note', format: 'classic' }]);
    const before = JSON.stringify(source), result = projectWorkbookAnnotations(source, frame(), dimensions), note = result.visible[0]!;
    expect(note.annotation).toBe(source.annotations[0]); expect(note.native?.format).toBe('classic'); expect(note.regions.map(region => region.pane)).toEqual(['corner', 'frozenRows', 'frozenColumns', 'body']);
    const anchor = source.annotations[0]!.anchor; if (anchor.kind !== 'cells') throw Error('Expected a cell anchor');
    expect(note.regions.every(region => region.range === anchor.ranges[0])).toBe(true);
    expect(note.regions[3]).toMatchObject({ rangeIndex: 0, range: range(3, 2), rect: { x: -40, y: -30, width: 140, height: 70 }, clip: { x: 20, y: 20, width: 80, height: 20 } });
    expect(JSON.stringify(source)).toBe(before); expect(Object.isFrozen(note.regions[0]!.clip)).toBe(true);
  });
  it('keeps disjoint regions, overlapping annotations and resolved/closed state separate', () => {
    const source = snapshot([annotation('thread', [range(0, 0), range(9, 6), range(90, 90)], 'resolved'), annotation('other', [range(9, 6)], 'closed')], [],
      [{ annotationId: 'thread', format: 'modern', replies: [] }]);
    const result = projectWorkbookAnnotations(source, frame(), dimensions);
    expect(result.visible.map(item => item.annotation.status)).toEqual(['resolved', 'closed']);
    expect(result.visible[0]!.regions.map(region => [region.rangeIndex, region.pane])).toEqual([[0, 'corner'], [1, 'body']]);
    expect(result.visible[0]!.native?.format).toBe('modern'); expect(result.visible[1]!.native).toBeUndefined();
    expect(result.visible[1]!.regions[0]!.clip).toEqual(result.visible[0]!.regions[1]!.clip);
  });
  it('retains recovered threads including missing sheets without assigning trustworthy grid geometry', () => {
    const source = snapshot([annotation('unlocated', [], 'needs_reanchor'), annotation('mismatch', [range(0, 0)], 'needs_reanchor'),
      annotation('deleted-sheet', [range(0, 0)], 'needs_reanchor', 'missing-sheet')], [],
      [{ annotationId: 'unlocated', format: 'modern', replies: [], recoveryStatus: 'resolved', recoveryReasons: ['missing-location'] }]);
    const result = projectWorkbookAnnotations(source, frame(), dimensions);
    expect(result.visible).toEqual([]); expect(result.recovery.map(item => item.annotation.annotationId)).toEqual(['unlocated', 'mismatch', 'deleted-sheet']);
    expect(result.recovery[0]!.native).toMatchObject({ recoveryStatus: 'resolved', recoveryReasons: ['missing-location'] });
    expect(result.recovery[0]!.annotation.instruction).toBe('原始批注 unlocated');
  });
  it('projects a whole-sheet range over a million hidden rows with only boundary queries', () => {
    const source = snapshot([annotation('all', [range(0, 0, 1_048_575, 16_383)]), annotation('hidden', [range(0, 0)]), annotation('last', [range(1_048_575, 1)])]);
    const sparse = { rows: new AxisDimensions(1_048_576, 10, [{ from: 0, to: 1_048_574, sizePt: 0 }]), columns: dimensions.columns };
    const window = frame(sparse, { rows: 0, columns: 0 }, { left: 0, top: 0 }, { width: 40, height: 10 });
    const rows = vi.spyOn(sparse.rows, 'position'), columns = vi.spyOn(sparse.columns, 'position');
    const result = projectWorkbookAnnotations(source, window, sparse);
    expect(result.visible.map(item => item.annotation.annotationId)).toEqual(['all', 'last']);
    expect(result.visible[0]!.regions[0]!.clip).toEqual({ x: 0, y: 0, width: 40, height: 10 });
    expect(rows).toHaveBeenCalledTimes(6); expect(columns).toHaveBeenCalledTimes(6); rows.mockRestore(); columns.mockRestore();
  });
  it('rejects mismatched frames and cumulative source/fragment limits without partial results', () => {
    const source = snapshot([annotation('a', [range(0, 0, 6, 6)]), annotation('b', [range(9, 6), range(10, 6)])]);
    for (const limits of [{ maxAnnotations: 1 }, { maxRanges: 2 }, { maxFragments: 3 }])
      expect(() => projectWorkbookAnnotations(source, frame(), dimensions, limits)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    expect(() => projectWorkbookAnnotations(source, { ...frame(), revision: 1 }, dimensions)).toThrowError(expect.objectContaining({ code: 'RevisionConflict' }));
    expect(() => projectWorkbookAnnotations(source, frame(), dimensions, { maxRanges: 0 })).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
  });
});
