import { workbookCellBorders } from "../../../packages/render/src/workbook/cell-borders.js";
import { createWorkbookBorderView } from "../../../packages/render/src/workbook/border-view.js";

export function borderSourceMeasurement(scope: HTMLElement) {
  const source = document.createElement('div'); source.style.cssText = 'color:rgb(10,20,30);--office-theme-accent1:#4472c4'; scope.append(source);
  const read = (css: string, rtl = false) => {
    const cell = document.createElement('div'); cell.style.cssText = css; source.append(cell);
    try { const before = cell.style.cssText, borders = workbookCellBorders(document, getComputedStyle(cell), rtl); if (before !== cell.style.cssText) throw Error('Source was changed'); return borders; }
    finally { cell.remove(); }
  };
  try {
    const nativeCss = '--office-cell-border-left-kind:thin;--office-cell-border-left-color:var(--office-theme-accent1);--office-cell-border-left-tint:.4;--office-cell-border-right-kind:double;--office-cell-border-bottom-kind:none;border-bottom:8px solid red';
    const physicalCss = 'border-left:4px double rgba(10,20,30,.5);border-right:2px dashed blue;border-top:1px dotted green;border-bottom:hidden';
    const native = [read(nativeCss), read(nativeCss, true)], css = [read(physicalCss), read(physicalCss, true)];
    const aliases = [read('--office-cell-border-start-kind:medium;--office-cell-border-end-kind:thin'), read('--office-cell-border-left-kind:none;--office-cell-border-start-kind:thin'),
      read('--office-cell-border-left-kind:thin;--office-cell-border-start-kind:thin')];
    const failures: string[] = [];
    for (const value of ['--office-cell-border-left-kind:constructor', '--office-cell-border-left-kind:thin;--office-cell-border-start-kind:double', '--office-cell-border-right-kind:thin;--office-cell-border-end-kind:none', 'border-top:3px groove red', '--office-cell-border-top-kind:thin;--office-cell-border-top-color:not-color']) {
      try { read(value); } catch (error) { failures.push((error as { code: string }).code); }
    }
    const layers = document.createElement('div'); layers.id = 'workbook-css-border-probe'; layers.style.cssText = 'position:absolute;left:0;top:0;width:200px;height:120px;background:white';
    const styles = [css[0]!.before!, css[0]!.after!, css[0]!.top!, native[0]!.after!];
    const svg = createWorkbookBorderView(document, { x: 0, y: 0, width: 150, height: 90 }, styles.map((border, index) => ({ ...border, from: { x: 10, y: 15 + index * 20 }, to: { x: 140, y: 15 + index * 20 } })));
    layers.append(svg); scope.append(layers);
    const strokes = [...svg.querySelectorAll('line')].map(line => ({ width: line.getAttribute('stroke-width'), dash: line.getAttribute('stroke-dasharray'), cap: line.getAttribute('stroke-linecap'), color: line.getAttribute('stroke') }));
    return { native, css, aliases, failures, strokes, unchanged: source.childNodes.length === 0 };
  } finally { source.remove(); }
}
