import { expect, it } from 'vitest';
import { resolveWorkbookCommentMention } from '../src/index.js';

it('resolves exact UTF-16 mentions without splitting surrogate pairs or variation sequences', () => {
  const text = 'A😀禰󠄀B';
  for (const [startIndex, length, expected] of [[1, 5, { from: 1, to: 6 }], [3, 3, { from: 3, to: 6 }], [1, 2, { from: 1, to: 3 }],
    [7, 0, { from: 7, to: 7 }], [4, 2, undefined], [3, 1, undefined], [2, 1, undefined], [8, 0, undefined],
    [-1, 1, undefined], [1, .5, undefined], [4294967295, 4294967295, undefined]] as const) {
    expect(resolveWorkbookCommentMention(text, { startIndex, length })).toEqual(expected);
  }
  expect(resolveWorkbookCommentMention('', { startIndex: 0, length: 0 })).toEqual({ from: 0, to: 0 });
});
