/** Independent deterministic storage load, not an XLSX export or formula evaluator. */
export const baselineRows = 10_000;
export const baselineColumns = 20;
export function generateBlock(row, rows = baselineRows) {
  const parts = ['<!doctype html><html><head></head><body><table data-sheet-id="baseline" data-row="' + row + '" data-column="0"><tbody>'];
  for (let r = row; r < Math.min(row + 256, rows); r++) {
    parts.push('<tr data-row="' + r + '">');
    let total = 0;
    for (let c = 0; c < baselineColumns; c++) {
      const value = (r + 1) * 20 + c;
      if (c >= 2 && c < 18) total += value;
      if (c === 1) parts.push('<td data-column="1" data-value-type="text">客户 ' + r + ' 😀禰󠄀</td>');
      else if (c < 18) parts.push('<td data-column="' + c + '" data-value-type="number" data-value="' + value + '">' + value + '</td>');
      else {
        const formula = c === 18 ? '=SUM(C' + (r + 1) + ':R' + (r + 1) + ')' : '=A' + (r + 1) + '+S' + (r + 1);
        const cached = c === 18 ? total : total + (r + 1) * 20;
        parts.push('<td data-column="' + c + '" data-value-type="number" data-value="' + cached + '" data-formula="' + formula + '" data-cache-state="imported">' + cached + '</td>');
      }
    }
    parts.push('</tr>');
  }
  parts.push('</tbody></table></body></html>');
  return new TextEncoder().encode(parts.join(''));
}
