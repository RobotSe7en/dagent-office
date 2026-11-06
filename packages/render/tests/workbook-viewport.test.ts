import { describe, expect, it } from "vitest";
import { AxisDimensions } from "../src/workbook/dimensions.js";
import { viewportWindows, workbookViewport } from "../src/workbook/viewport.js";

const dimensions = { rows: new AxisDimensions(100, 10, []), columns: new AxisDimensions(100, 20, []) };
const viewport = { widthPt: 100, heightPt: 80, scrollLeftPt: 40, scrollTopPt: 30 };
describe("workbook frozen viewport partition", () => {
  it("maps four panes to disjoint source cells and one shared element budget", () => {
    const layout = workbookViewport(dimensions, { rows: 2, columns: 1 }, viewport, 40);
    expect(layout.cellCount).toBe(40); expect(layout.panes.map(pane => pane.kind)).toEqual(['corner', 'frozenRows', 'frozenColumns', 'body']);
    expect(layout.panes[3]).toMatchObject({ inlineStartPt: 20, blockStartPt: 20, widthPt: 80, heightPt: 60, sourceLeftPt: 60, sourceTopPt: 50 });
    expect(viewportWindows('sheet', layout)).toEqual([
      { sheetId: 'sheet', row: 0, column: 0, rows: 2, columns: 1 }, { sheetId: 'sheet', row: 0, column: 3, rows: 2, columns: 4 },
      { sheetId: 'sheet', row: 5, column: 0, rows: 6, columns: 1 }, { sheetId: 'sheet', row: 5, column: 3, rows: 6, columns: 4 },
    ]);
    const cells = layout.panes.flatMap(pane => pane.rows.flatMap(row => pane.columns.map(column => `${row.index}:${column.index}`)));
    expect(new Set(cells).size).toBe(40);
    expect(() => workbookViewport(dimensions, { rows: 2, columns: 1 }, viewport, 39)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
  it("queries distant visible rows separately instead of expanding hidden source coordinates", () => {
    const sparse = { rows: new AxisDimensions(1_048_576, 15, [{ from: 1, to: 1_048_574, sizePt: 0 }]), columns: new AxisDimensions(16_384, 20, []) };
    const layout = workbookViewport(sparse, { rows: 0, columns: 0 }, { widthPt: 40, heightPt: 30, scrollLeftPt: 0, scrollTopPt: 0 }, 4);
    expect(layout.cellCount).toBe(4);
    expect(viewportWindows('sheet', layout)).toEqual([{ sheetId: 'sheet', row: 0, column: 0, rows: 1, columns: 2 }, { sheetId: 'sheet', row: 1_048_575, column: 0, rows: 1, columns: 2 }]);
  });
  it("clamps the logical scroll extent and handles frozen panes larger than the viewport", () => {
    const end = workbookViewport(dimensions, { rows: 2, columns: 1 }, { ...viewport, scrollLeftPt: 10000, scrollTopPt: 10000 }, 40);
    expect(end.scrollLeftPt).toBe(1900); expect(end.scrollTopPt).toBe(920);
    expect(end.panes[3]?.rows.at(-1)?.index).toBe(99); expect(end.panes[3]?.columns.at(-1)?.index).toBe(99);
    const frozen = workbookViewport(dimensions, { rows: 99, columns: 0 }, viewport, 40);
    expect(frozen.panes.map(pane => pane.kind)).toEqual(['frozenRows']); expect(frozen.maxScrollTopPt).toBe(0);
    expect(workbookViewport(dimensions, { rows: 0, columns: 0 }, { ...viewport, widthPt: 0 }, 40).panes).toEqual([]);
  });
  it("does not materialize invisible axes and rejects malformed geometry", () => {
    const hidden = { ...dimensions, rows: new AxisDimensions(100, 0, []) };
    expect(workbookViewport(hidden, { rows: 2, columns: 1 }, viewport, 1).cellCount).toBe(0);
    expect(workbookViewport({ ...dimensions, columns: new AxisDimensions(100, 0, []) }, { rows: 2, columns: 1 }, viewport, 1).cellCount).toBe(0);
    expect(() => workbookViewport(dimensions, { rows: 100, columns: 1 }, viewport, 5000)).toThrow();
    expect(() => workbookViewport(dimensions, { rows: 0, columns: 0 }, { ...viewport, scrollTopPt: NaN }, 5000)).toThrow();
  });
});
