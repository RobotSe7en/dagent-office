/** The browser's own shaper measures simple numeric candidates without adding
 * a DOM probe for each value. Other typography stays on the DOM path. This is
 * a whole-string advance, never a sum or extrapolation of individual digits. */
export function workbookAsciiWidth(context: CanvasRenderingContext2D | null, style: CSSStyleDeclaration | undefined): ((value: string) => number | undefined) | undefined {
  if (!context || !style || !("fontKerning" in context) || style.fontVariant !== "normal" || style.textTransform !== "none"
    || !["normal", "0px", "0"].includes(style.letterSpacing) || !["normal", "0px", "0"].includes(style.wordSpacing)) return undefined;
  const font = `${style.fontStyle} ${style.fontVariant} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
  return value => {
    if (!/^[\x20-\x7e]*$/.test(value)) return undefined;
    context.font = "1px monospace"; const sentinel = context.font; context.font = font;
    if (context.font === sentinel) return undefined;
    context.fontKerning = "none"; context.direction = "ltr";
    const width = context.measureText(value).width;
    return Number.isFinite(width) && width >= 0 ? width : undefined;
  };
}
