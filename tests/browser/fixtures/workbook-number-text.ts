import { createWorkbookNumberFormatter } from "@doffi/core";
import type { WorkbookDisplayFormat } from "@doffi/core";
import { layoutWorkbookNumbers } from "../../../packages/render/src/workbook/number-text-layout.js";

export interface NumberTextCase { value: number; format?: WorkbookDisplayFormat; widthPt: number; heightPt?: number; styles?: Readonly<Record<string, string>> }
export async function numberTextMeasurement(scope: HTMLElement, cases: readonly NumberTextCase[], hostLocale = "zh-CN") {
  await document.fonts.ready;
  const sources = document.createElement("div"), gallery = document.createElement("div");
  sources.style.cssText = "position:absolute;visibility:hidden;width:0;height:0;overflow:hidden"; scope.append(sources);
  gallery.id = "workbook-number-text-probe"; gallery.style.cssText = "position:absolute;left:0;top:0;display:flex;flex-wrap:wrap;width:850px;gap:4px;background:white";
  const formatter = createWorkbookNumberFormatter({ dateSystem: "1900", hostLocale });
  const inputs = cases.map(item => {
    const table = document.createElement("table"), cell = table.insertRow().insertCell(); cell.textContent = "STALE IMPORTED DISPLAY";
    cell.dataset.valueType = "number"; cell.dataset.value = String(item.value);
    cell.style.cssText = 'font-family:"Liberation Sans";font-size:11pt;line-height:normal;color:#112233';
    cell.style.setProperty("--office-font-east-asia", '"Noto Sans CJK SC"');
    for (const [property, value] of Object.entries(item.styles ?? {})) cell.style.setProperty(property, value);
    sources.append(table);
    return { source: { element: cell, format: { cssClass: "" } }, value: item.value, ...(item.format ? { format: item.format } : {}), widthPt: item.widthPt, heightPt: item.heightPt ?? 30 };
  });
  const original = sources.innerHTML;
  try {
    const result = layoutWorkbookNumbers(scope, inputs, formatter, 4);
    for (const cell of result.cells) { cell.element.style.outline = "1px solid #ddd"; gallery.append(cell.element); }
    scope.append(gallery);
    const measurements = result.cells.map((cell, index) => {
      const bounds = cell.element.getBoundingClientRect(), content = cell.text.element.getBoundingClientRect();
      return { text: cell.text.element.textContent, widthPt: content.width * .75, heightPt: content.height * .75, rightGapPt: (bounds.right - content.right) * .75,
        color: getComputedStyle(cell.text.element).color, scale: cell.scale, overflow: result.overflow[index] ?? null, locale: result.locales[index],
        projections: cell.text.projections.length, breaks: cell.text.element.querySelectorAll("br").length, transform: cell.text.element.style.transform,
        spaces: [...cell.text.element.querySelectorAll<HTMLElement>("[data-office-number-space]")].map(span => ({ sourceWidth: Number.parseFloat(span.style.width), actualWidth: span.getBoundingClientRect().width })),
        fonts: [...cell.text.element.querySelectorAll<HTMLElement>("[data-office-script]")].map(span => ({ role: span.dataset.officeScript, family: getComputedStyle(span).fontFamily })) };
    });
    const failures: string[] = [];
    for (const [nodes, units] of [[1, 100000], [50000, 1]]) {
      try { layoutWorkbookNumbers(scope, inputs, formatter, 4, nodes, units); }
      catch (error) { failures.push((error as { code: string }).code); }
    }
    try { layoutWorkbookNumbers(scope, [inputs[0]!, { ...inputs[0]!, format: { code: "ss.0000" } }], formatter, 4); }
    catch (error) { failures.push((error as { code: string }).code); }
    try { layoutWorkbookNumbers(scope, [{ ...inputs[0]!, widthPt: Infinity }], formatter, 4); }
    catch (error) { failures.push((error as { code: string }).code); }
    formatter.dispose();
    try { layoutWorkbookNumbers(scope, inputs, formatter, 4); }
    catch (error) { failures.push((error as { code: string }).code); }
    return { measurements, fonts: result.fonts, failures, unchanged: original === sources.innerHTML, atomic: gallery.isConnected,
      retained: scope.querySelectorAll('[data-office-view="workbook-number-measurement"],[data-office-view="workbook-text-measurement"]').length };
  } finally { sources.remove(); formatter.dispose(); }
}
