import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import { importXlsx } from '@dagent-office/office';

describe('independent native hidden-column row-height sources', () => {
  it.each(['', '-producer'])('preserves%s hidden widths, typography, manual and merged rows on package reopen', async suffix => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-hidden-row-heights' + suffix + '.xlsx');
    if (!suffix) expect(execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('hidden','scripts/check-xlsx-hidden-row-heights.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"])).toEqual(bytes);
    const model = openDocument(await importXlsx(bytes, { documentId: 'hidden-column-height' }), { kind: 'workbook' }), snapshot = model.snapshot();
    expect(snapshot.metadata.workbook.sheets).toHaveLength(12);
    for (const [index, sheet] of snapshot.metadata.workbook.sheets.entries()) {
      const column = sheet.columnDimensions.find(run => run.from <= 1 && run.to >= 1)!;
      expect(column?.hidden ?? false).toBe(index >= 6 || index % 2 === 1);
      if (!suffix && index < 6) expect(column.width).toBe([8.43, 4, 16][Math.floor(index / 2)]);
      if (!suffix && index === 6) expect(column.width).toBeUndefined();
    }
    expect(snapshot.metadata.workbook.sheets[7]!.columnDimensions.some(run => run.cssClass && run.hidden)).toBe(true);
    expect(snapshot.metadata.workbook.sheets[11]!.merges).toEqual([{ from: { row: 1, column: 1 }, to: { row: 1, column: 2 } }]);
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
});
