import { evaluatePresentationPreset } from '@dagent-office/core';
import { createDrawingGeometryView } from '../../../packages/render/src/drawing-geometry.js';

export async function measureDrawingGradients() {
  const paint = document.createElement('div');
  paint.style.cssText = '--office-fill-kind:linear;--office-fill-angle:45;--office-fill-stop-count:2;--office-fill-stop-0-position:0;--office-fill-stop-0-color:#ff0000;--office-fill-stop-1-position:1;--office-fill-stop-1-color:#0000ff';
  document.body.append(paint);
  const input = { geometry: evaluatePresentationPreset('rect', 200, 100), preset: 'rect', id: 'pixel-gradient', widthPt: 200, heightPt: 100 };
  const samples = [[0, 0], [199, 0], [0, 99], [199, 99], [99, 49]] as const;
  try {
    const cases = [];
    for (const scaled of [false, true]) {
      paint.style.setProperty('--office-fill-scaled', String(scaled));
      const svg = createDrawingGeometryView(document, getComputedStyle(paint), input);
      svg.setAttribute('width', '200'); svg.setAttribute('height', '100'); svg.style.width = '200px'; svg.style.height = '100px';
      const gradient = svg.querySelector('linearGradient')!;
      const coordinates = ['x1', 'y1', 'x2', 'y2'].map(name => Number(gradient.getAttribute(name)));
      const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(svg)], { type: 'image/svg+xml' })), image = new Image();
      try {
        image.src = url; await image.decode();
        const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100;
        const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
        cases.push({ scaled, coordinates, units: gradient.getAttribute('gradientUnits'), pixels: samples.map(([x, y]) => [...context.getImageData(x, y, 1, 1).data]) });
      } finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
    }
    const errors = [];
    for (const [property, value] of [['--office-fill-angle', '45deg'], ['--office-fill-scaled', 'yes'], ['--office-fill-stop-1-position', '-1'], ['--office-fill-stop-0-position', '2'], ['--office-fill-rotate-with-shape', 'yes']]) {
      const previous = paint.style.getPropertyValue(property!); paint.style.setProperty(property!, value!);
      try { createDrawingGeometryView(document, getComputedStyle(paint), input); errors.push('accepted'); }
      catch (error) { errors.push((error as { code: string }).code); }
      finally { paint.style.setProperty(property!, previous); }
    }
    return { cases, errors };
  } finally { paint.remove(); }
}
