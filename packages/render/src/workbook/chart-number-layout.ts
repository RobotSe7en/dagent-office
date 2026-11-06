import { DocumentError } from '@doffi/core';
import type { BasicChartScene } from '../chart.js';
import { fitWorkbookFill } from './number-layout.js';

/** Resolve fields after actual local fonts are known. Blank advances contain
 * ordinary spaces, never source placeholder glyphs. Repetition uses the same
 * bounded shaping primitive as cells, measured on its final SVG font runs. */
export function layoutChartNumberFields(scene: BasicChartScene, accountText: (units: number) => void): void {
  for (const label of scene.numberTexts) {
    let lastFill = -1, fixed = 0, fixedUnits = 0;
    for (const [index, part] of label.parts.entries()) {
      if (part.kind === 'fill') { lastFill = index; continue; }
      fixed += label.fields[index]!.getComputedTextLength();
      fixedUnits += part.kind === 'text' ? part.text.length : part.character.length;
    }
    if (!Number.isFinite(fixed)) throw new DocumentError('ResourceLimitExceeded', 'Chart numeric fields exceed finite glyph geometry.');
    for (const [index, part] of label.parts.entries()) {
      if (part.kind === 'text') continue;
      const field = label.fields[index]!;
      if (field.childNodes.length !== 1 || field.firstChild?.nodeType !== 3) throw new DocumentError('InvalidDocument', 'Chart format fields require one shaped Unicode character.');
      const node = field.firstChild as Text;
      if (part.kind === 'space') {
        const width = field.getComputedTextLength(); accountText(1); node.data = ' ';
        const blank = field.getComputedTextLength();
        if (![width, blank].every(value => Number.isFinite(value) && value >= 0)) throw new DocumentError('ResourceLimitExceeded', 'Chart numeric spacing exceeds finite glyph geometry.');
        // Some engines ignore textLength on tspans. Explicit advance plus a
        // preserved blank retains exact placement across the browser matrix.
        field.setAttribute('dx', String(width - blank));
        field.setAttribute('data-sdk-chart-number-space', String(width));
      } else {
        const repeated = index === lastFill ? fitWorkbookFill(part.character, Math.max(0, label.availableWidth - fixed), fixedUnits, sample => {
          node.data = sample; return field.getComputedTextLength();
        }) : { text: '', width: 0 };
        accountText(repeated.text.length); node.data = repeated.text;
        field.setAttribute('data-sdk-chart-number-fill', '');
      }
      field.removeAttribute('visibility');
    }
  }
}
