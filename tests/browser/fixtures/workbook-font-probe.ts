import { FontProbeCache } from '../../../packages/render/src/font-probe-cache.js';
import { resolvePreviewFontsInScope } from '../../../packages/render/src/fonts.js';
export async function fontProbeMeasurement() {
  await document.fonts.ready;
  const scope = document.createElement('div'), text = document.createElement('span'); text.dataset.officeScript = 'east-asia';
  text.style.fontFamily = 'serif'; text.textContent = '客户禰󠄀'; scope.append(text); document.body.append(scope);
  const cache = new FontProbeCache(), prototype = CanvasRenderingContext2D.prototype, original = prototype.getImageData; let reads = 0;
  prototype.getImageData = function(...args: Parameters<typeof original>) { reads++; return original.apply(this, args); };
  try {
    const first = resolvePreviewFontsInScope(document, scope, '[data-office-script]', cache), initial = reads;
    const second = resolvePreviewFontsInScope(document, scope, '[data-office-script]', cache), repeated = reads;
    cache.clear(); const third = resolvePreviewFontsInScope(document, scope, '[data-office-script]', cache);
    return { first, second, third, initial, repeated, cleared: reads };
  } finally { prototype.getImageData = original; cache.clear(); scope.remove(); }
}
