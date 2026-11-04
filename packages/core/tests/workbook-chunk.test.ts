import { describe, expect, it } from "vitest";
import { indexWorkbookChunk } from "../src/workbook/chunk-index.js";

const encoder = new TextEncoder();
const position = { sheetId: "sheet-1", row: 256, column: 32 };
function source(content: string): Uint8Array {
  return encoder.encode('<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="256" data-column="32"><tbody>' + content + '</tbody></table></body></html>');
}
const content = '<tr data-row="256"><td data-column="32" data-value-type="text"><span class="east-asian">中😀禰󠄀</span><br>字</td><td data-column="34" data-value-type="number" data-value="3.5" data-formula="=AG257+1" data-cache-state="imported">3.50</td></tr><tr data-row="511"><td data-column="63" data-value-type="boolean" data-value="false">FALSE</td></tr>';

describe("private workbook chunk storage baseline", () => {
  it("indexes sparse coordinates, preserves UTF-16 source and isolates bytes", () => {
    const bytes = source(content), copy = bytes.slice(), index = indexWorkbookChunk(bytes, position);
    expect(index.cellCount).toBe(3); expect(index.formulaCount).toBe(1); expect(index.indexBytes).toBe(30);
    bytes.fill(0); index.serialize().fill(0);
    expect(index.serialize()).toEqual(copy);
    const window = index.readWindow({ row: 256, column: 32, rows: 1, columns: 2 });
    expect(window).toEqual([{ row: 256, column: 32, html: '<td data-column="32" data-value-type="text"><span class="east-asian">中😀禰󠄀</span><br>字</td>' }]);
    expect(Object.isFrozen(window[0])).toBe(true);
    expect(index.readWindow({ row: 511, column: 63, rows: 1, columns: 1 })[0]?.html).toContain('data-value="false"');
    expect(index.readWindow({ row: 0, column: 0, rows: 1, columns: 1 })).toEqual([]);
  });
  it.each([
    content.replace('data-column="34"', 'data-column="32"'),
    content.replace('data-column="34"', 'data-column="31"'),
    content.replace('data-row="511"', 'data-row="512"'),
    content.replace('data-row="511"', 'data-row="256"'),
    content.replace('data-value="3.5"', 'data-value="Infinity"'),
    content.replace('data-value="false"', 'data-value="0"'),
    content.replace('data-cache-state="imported"', ''),
    content.replace('<br>', '<br onclick="alert(1)">'),
    content.replace('<br>', '<script>alert(1)</script>'),
    content.replace('<br>', '<img src="https://example.com/a.png">'),
    content.replace('<br>', '<!-- unmodeled -->'),
    content.replace('class="east-asian"', 'class="east-asian" style="color:red"'),
    content.replace('</td>', ''),
    content.replace('<span', '<table'),
  ])("rejects malformed, unordered or unmodeled source before indexing", input => {
    expect(() => indexWorkbookChunk(source(input), position)).toThrow();
  });
  it("enforces source budgets and bounds before keeping data", () => {
    expect(() => indexWorkbookChunk(source(content), position, 10)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(() => indexWorkbookChunk(new Uint8Array([0xff]), position)).toThrow();
    expect(() => indexWorkbookChunk(source(content), { ...position, row: 257 })).toThrow();
    expect(() => indexWorkbookChunk(source(content), { ...position, sheetId: "another" })).toThrow();
    const index = indexWorkbookChunk(source(content), position);
    for (const window of [
      { row: -1, column: 0, rows: 1, columns: 1 },
      { row: 0, column: 0, rows: 0, columns: 1 },
      { row: 0, column: 0, rows: 1_048_577, columns: 1 },
      { row: 0, column: 16_384, rows: 1, columns: 1 },
    ]) expect(() => index.readWindow(window)).toThrow();
  });
});
