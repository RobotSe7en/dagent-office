import { expect, it } from 'vitest';
import { decodeWorkbookCell } from '../src/workbook/chunk-index.js';
const cases = [
  ['data-value-type="text" class="字型"', '中文 Latin 😀禰󠄀 >', { type: 'text', value: '中文 Latin 😀禰󠄀 >' }],
  ['data-value-type="number" data-value="1e3" data-format-id="general"', '1000.00', { type: 'number', value: 1000 }],
  ['data-value-type="boolean" data-value="false"', 'FALSE', { type: 'boolean', value: false }],
  ['data-value-type="error" data-value="#N/A"', '#N/A', { type: 'error', value: '#N/A' }],
  ['data-value-type="blank"', '', { type: 'blank' }],
  ['data-value-type="text"', '', { type: 'text', value: '' }],
  ['data-value-type="number" data-value="99" data-formula="=SUM(A1:B1)" data-cache-state="imported"', '99', { type: 'number', value: 99 }],
] as const;
it.each(cases)('matches the full parser for canonical cell %s', (attributes, text, value) => {
  const html = `<td data-column="0" ${attributes}>${text}</td>`, plain = decodeWorkbookCell(html);
  // CR in inter-attribute whitespace forces the original parser without
  // changing HTML semantics. The opening path always uses that parser too.
  expect(plain).toEqual(decodeWorkbookCell(html.replace('<td ', '<td\r ')));
  expect(plain.value).toEqual(value); expect(plain.text).toBe(text);
});
it('keeps entities, rich content, character-reference CR and quoted formula characters on the full parser', () => {
  const rich = decodeWorkbookCell('<td data-column="0" data-value-type="text"><span>😀&#x79B0;&#xE0100;</span><br>&amp;&#13;&lt;</td>');
  expect(rich.value).toEqual({ type: 'text', value: '😀禰󠄀\n&\r<' });
  expect(decodeWorkbookCell('<td data-column="0" data-value-type="number" data-value="-0" data-formula="=IF(A1&gt;0,1,2)" data-cache-state="stale">-0</td>').formula)
    .toEqual({ expression: '=IF(A1>0,1,2)', cacheState: 'stale' });
  expect(decodeWorkbookCell("<TD data-column='0' data-value-type='text'>single quoted</TD>").text).toBe('single quoted');
});
