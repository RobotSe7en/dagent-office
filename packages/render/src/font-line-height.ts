/** Scalable font bounds avoid accumulating whole-pixel normal-leading rounding.
 * Undefined means the engine or font shorthand cannot supply these metrics. */
export function measuredFontLineHeight(context: CanvasRenderingContext2D | null, computed: CSSStyleDeclaration): number | undefined {
  const size = Number.parseFloat(computed.fontSize);
  if (size === 0) return 0;
  if (!context) return undefined;
  // Invalid Canvas shorthand is silently ignored: never measure a prior run.
  context.font = "1px monospace";
  const sentinel = context.font;
  context.font = `${computed.fontStyle} ${computed.fontVariant} ${computed.fontWeight} 1024px ${computed.fontFamily}`;
  if (context.font === sentinel) return undefined;
  const metrics = context.measureText("M"), height = (metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent) * size / 1024;
  return Number.isFinite(height) && height > 0 ? height : undefined;
}
