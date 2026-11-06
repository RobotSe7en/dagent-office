import { DocumentError } from "@dagent-office/core";
import { markViewNode } from "../view-nodes.js";

interface Box { left: number; top: number; right: number; bottom: number }
interface Line extends Box { node: Text }
const markerTypography = ["font-family", "font-size", "font-weight", "font-style", "font-stretch", "font-variant", "line-height", "letter-spacing", "color", "direction", "unicode-bidi"];

/** Clip the private view, leaving every original UTF-16 text node intact.
 * Ellipses are derived sibling markers with no source offsets. A polygon cuts
 * at complete grapheme bounds on the first/last visible line; earlier rich
 * lines keep their measured wrapping and backgrounds. Inspection is bounded,
 * and runs before the object's compensating text reflection is installed. */
export function applyDrawingTextOverflow(text: HTMLElement, box: Box, horizontal: "overflow" | "clip", vertical: "overflow" | "clip" | "ellipsis",
  account: (nodes: number) => void, maxProbes = 100_000): void {
  if (horizontal === "overflow" && vertical === "overflow") return;
  const document = text.ownerDocument, view = document.defaultView!, origin = text.getBoundingClientRect(), range = document.createRange();
  let probes = 0;
  const probe = (): void => { if (++probes > maxProbes) throw new DocumentError("ResourceLimitExceeded", "Drawing text overflow inspection exceeds its geometry budget."); };
  const local = (rect: DOMRect): Box => ({ left: rect.left - origin.left, top: rect.top - origin.top, right: rect.right - origin.left, bottom: rect.bottom - origin.top });
  const lines: Line[] = [], textNodes: Text[] = [];
  let left = Math.min(0, box.left), top = Math.min(0, box.top), right = Math.max(origin.width, box.right), bottom = Math.max(origin.height, box.bottom);
  const walker = document.createTreeWalker(text, view.NodeFilter.SHOW_ELEMENT | view.NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    probe();
    if (node.nodeType === 3) { range.selectNodeContents(node); textNodes.push(node as Text); }
    const rectangles = node.nodeType === 3 ? range.getClientRects() : (node as Element).getClientRects();
    for (const rect of rectangles) {
      probe(); const item = local(rect);
      left = Math.min(left, item.left); top = Math.min(top, item.top); right = Math.max(right, item.right); bottom = Math.max(bottom, item.bottom);
      if (vertical === "ellipsis" && node.nodeType === 3 && rect.height > 0 && rect.width > 0) lines.push({ ...item, node: node as Text });
    }
  }
  if (![left, top, right, bottom].every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "Drawing text overflow has nonfinite geometry.");
  // One CSS pixel protects antialiasing overhang on the unconstrained axis.
  left--; top--; right++; bottom++;
  if (horizontal === "clip") { left = box.left; right = box.right; }
  if (vertical !== "overflow") { top = box.top; bottom = box.bottom; }
  let upper: { x: number; y: number; rtl: boolean } | undefined, lower: typeof upper;
  let hideContent = false;
  if (vertical === "ellipsis" && lines.length && box.right > box.left && box.bottom > box.top) {
    lines.sort((a, b) => a.top - b.top || a.bottom - b.bottom);
    const grouped: Line[] = [];
    for (const line of lines) {
      const previous = grouped.at(-1);
      // Font bounding boxes can overlap the next line by a fraction of a pixel
      // even when native leading is correct. A tiny overlap is not the same
      // shaped line; mixed sizes on one baseline overlap substantially.
      if (previous && previous.node.parentElement!.closest('p') === line.node.parentElement!.closest('p')
        && Math.min(line.bottom, previous.bottom) - Math.max(line.top, previous.top) > Math.min(line.bottom - line.top, previous.bottom - previous.top) / 2) {
        previous.left = Math.min(previous.left, line.left); previous.right = Math.max(previous.right, line.right);
        previous.top = Math.min(previous.top, line.top); previous.bottom = Math.max(previous.bottom, line.bottom);
      } else grouped.push({ ...line });
    }
    // CSS range font bounds round differently from the measured native leading
    // (notably in Chromium). The same one-pixel tolerance as auto-fit avoids
    // treating a fitted final line's font-box overhang as another hidden line.
    const visible = grouped.filter(line => line.top >= box.top - 1 && line.bottom <= box.bottom + 1);
    const hiddenAbove = grouped[0]!.top < box.top - 1, hiddenBelow = grouped.at(-1)!.bottom > box.bottom + 1;
    const add = (line: Line, leading: boolean): { x: number; y: number; rtl: boolean } => {
      if (typeof Intl.Segmenter !== "function") throw new DocumentError("UnsupportedFeature", "Drawing text ellipses require Unicode grapheme segmentation.");
      account(2);
      const marker = markViewNode(document.createElement("span")), computed = view.getComputedStyle(line.node.parentElement!);
      marker.dataset.officeView = "workbook-drawing-ellipsis"; marker.setAttribute("aria-hidden", "true"); marker.textContent = "\u2026";
      for (const name of markerTypography) marker.style.setProperty(name, computed.getPropertyValue(name));
      Object.assign(marker.style, { position: "absolute", display: "block", whiteSpace: "pre", margin: "0", padding: "0", border: "0", pointerEvents: "none", top: "0", left: "0" });
      text.append(marker);
      const measured = marker.getBoundingClientRect(), rtl = computed.direction === "rtl", fromRight = leading ? !rtl : rtl;
      const target = fromRight ? Math.max(box.left, line.left) + measured.width : Math.min(box.right, line.right) - measured.width;
      // Shape complete graphemes with the browser. Reading whole text runs first
      // avoids inspecting offscreen paragraphs character by character.
      const glyphs: { left: number; right: number }[] = [];
      const segmenter = new Intl.Segmenter("und", { granularity: "grapheme" });
      const groups = new Map<HTMLElement, { node: Text; from: number; to: number }[]>();
      for (const node of textNodes) {
        const paragraph = node.parentElement!.closest('p') ?? text, group = groups.get(paragraph) ?? [], from = group.at(-1)?.to ?? 0;
        group.push({ node, from, to: from + node.length }); groups.set(paragraph, group);
      }
      for (const group of groups.values()) {
        let intersects = false;
        for (const { node } of group) {
          probe(); range.selectNodeContents(node);
          for (const rect of range.getClientRects()) { probe(); if (rect.bottom > origin.top + line.top + .1 && rect.top < origin.top + line.bottom - .1) intersects = true; }
        }
        if (!intersects) continue;
        if (group.at(-1)!.to > maxProbes * 16) throw new DocumentError("ResourceLimitExceeded", "Drawing ellipsis segmentation exceeds its text budget.");
        const content = group.map(item => item.node.data).join('');
        let start = 0, end = 0;
        for (const segment of segmenter.segment(content)) {
          probe(); const to = segment.index + segment.segment.length;
          while (start < group.length - 1 && group[start]!.to <= segment.index) start++;
          end = Math.max(start, end); while (end < group.length - 1 && group[end]!.to < to) end++;
          range.setStart(group[start]!.node, segment.index - group[start]!.from); range.setEnd(group[end]!.node, to - group[end]!.from);
          let glyphLeft = Infinity, glyphRight = -Infinity;
          for (const rect of range.getClientRects()) {
            probe(); const part = local(rect);
            if (Math.min(part.bottom, line.bottom) - Math.max(part.top, line.top) > Math.min(part.bottom - part.top, line.bottom - line.top) / 2 && part.right > part.left) {
              glyphLeft = Math.min(glyphLeft, part.left); glyphRight = Math.max(glyphRight, part.right);
            }
          }
          if (Number.isFinite(glyphLeft) && Number.isFinite(glyphRight)) glyphs.push({ left: glyphLeft, right: glyphRight });
        }
      }
      // Font/script spans and bidi runs can have overlapping shaped bounds.
      // A boundary of one grapheme must not cut through another one.
      glyphs.sort((a, b) => a.left - b.left || a.right - b.right);
      const unions: { left: number; right: number }[] = [];
      for (const glyph of glyphs) {
        const previous = unions.at(-1);
        if (previous && glyph.left < previous.right - 1) previous.right = Math.max(previous.right, glyph.right);
        else unions.push({ ...glyph });
      }
      const edges = unions.flatMap(glyph => fromRight ? glyph.left >= target - .01 ? [glyph.left] : [] : glyph.right <= target + .01 ? [glyph.right] : []);
      const edge = edges.reduce((value, next) => fromRight ? Math.min(value, next) : Math.max(value, next), fromRight ? Infinity : -Infinity);
      const x = fromRight ? edges.length ? Math.min(box.right, edge) : right : edges.length ? Math.max(box.left, edge) : left;
      // Range font bounds and the CSS line box may have different ascenders.
      // Align the marker's own range top to this shaped line instead of using
      // the default paragraph's (possibly zero) font strut.
      range.selectNodeContents(marker); const glyph = range.getBoundingClientRect();
      marker.style.top = line.top - (glyph.top - measured.top) + "px";
      marker.style.left = Math.max(box.left, Math.min(box.right - measured.width, fromRight ? x - measured.width : x)) + "px";
      return { x, y: leading ? line.bottom : line.top, rtl: fromRight };
    };
    if (visible.length) {
      if (hiddenAbove) { upper = add(visible[0]!, true); top = Math.max(top, visible[0]!.top); }
      if (hiddenBelow) { lower = add(visible.at(-1)!, false); bottom = Math.min(bottom, visible.at(-1)!.bottom); }
    } else if (hiddenAbove || hiddenBelow) {
      // No complete line fits. A clipped derived marker still signals hidden
      // text without inventing a visible source character or shrinking fonts.
      lower = add({ ...grouped[0]!, top: box.top, bottom: box.bottom }, false);
      hideContent = true;
    }
  }
  // Markers are positioned siblings outside the original content's clip. Give
  // each original paragraph the same local polygon; this keeps markers visible
  // while preserving source nodes and the exact layout of every rich line.
  const points: [number, number][] = [[left, top]];
  if (upper && !upper.rtl) points.push([upper.x, top], [upper.x, upper.y], [right, upper.y]);
  else points.push([right, top]);
  if (lower && !lower.rtl) points.push([right, lower.y], [lower.x, lower.y], [lower.x, bottom]);
  else points.push([right, bottom]);
  if (lower?.rtl) points.push([lower.x, bottom], [lower.x, lower.y], [left, lower.y]);
  else points.push([left, bottom]);
  if (upper?.rtl) points.push([left, upper.y], [upper.x, upper.y], [upper.x, top]);
  for (const child of text.children) {
    const node = child as HTMLElement;
    if (node.dataset.officeView === "workbook-drawing-ellipsis") continue;
    if (hideContent) node.style.visibility = "hidden";
    const rect = node.getBoundingClientRect(), x = rect.left - origin.left, y = rect.top - origin.top;
    node.style.clipPath = "polygon(" + points.map(point => `${point[0] - x}px ${point[1] - y}px`).join(",") + ")";
  }
  // The marker also obeys both vertical barriers and any horizontal clipping.
  for (const node of text.querySelectorAll<HTMLElement>('[data-office-view="workbook-drawing-ellipsis"]')) {
    const rect = node.getBoundingClientRect(), x = rect.left - origin.left, y = rect.top - origin.top;
    node.style.clipPath = `polygon(${left - x}px ${top - y}px,${right - x}px ${top - y}px,${right - x}px ${bottom - y}px,${left - x}px ${bottom - y}px)`;
  }
}
