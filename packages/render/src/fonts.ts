import type { FontResolution } from "./contracts.js";
export type { FontResolution } from "./contracts.js";
const generic = new Set([
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
]);
function families(value: string): string[] {
  return (value.match(/"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|[^,]+/g) ?? [])
    .map((name) =>
      name
        .trim()
        .replace(/^['"]|['"]$/g, "")
        .replace(/\\(["'\\])/g, "$1"),
    )
    .filter(Boolean);
}
/** Measure against multiple fallback metrics; this neither requests permission nor loads fonts. */
export function resolvePreviewFonts(
  document: Document,
  selector: string = "[data-node-id]",
): readonly FontResolution[] {
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d")!;
  const cache = new Map<string, boolean>();
  const resolutions = new Map<
    string,
    {
      requestedFamily: string;
      selectedFamily: string;
      substituted: boolean;
      missingCharacters: Set<string>;
      glyphAuditComplete: boolean;
    }
  >();
  const selected = new WeakMap<HTMLElement, string>();
  const available = (family: string): boolean => {
    if (generic.has(family.toLowerCase())) return true;
    const known = cache.get(family);
    if (known !== undefined) return known;
    const quoted = JSON.stringify(family);
    const sample = "汉字中文，。AaWim0123";
    let found = false;
    for (const base of ["monospace", "serif", "sans-serif"]) {
      context.font = "72px " + base;
      const fallback = context.measureText(sample).width;
      context.font = "72px " + quoted + "," + base;
      if (Math.abs(context.measureText(sample).width - fallback) > 0.01)
        found = true;
    }
    cache.set(family, found);
    return found;
  };
  const nodes = [...document.querySelectorAll<HTMLElement>(selector)];
  // Resolve parents before children so inherited fallback and native east-Asian overrides remain distinct.
  for (const node of nodes) {
    const requested = families(
      document.defaultView!.getComputedStyle(node).fontFamily,
    );
    if (!requested.length) continue;
    const first = requested[0]!;
    let rendered = requested.find(available);
    if (!rendered) {
      const candidates = /仿宋|FangSong/i.test(first)
        ? ["FandolFang", "Noto Serif CJK SC", "serif"]
        : /楷|KaiTi|SimKai/i.test(first)
          ? ["FandolKai", "LXGW WenKai", "Noto Serif CJK SC", "serif"]
          : /宋|明|SimSun|NSimSun|Ming|Mincho|Serif/i.test(first)
            ? ["Noto Serif CJK SC", "FandolSong", "Noto Sans CJK SC", "serif"]
            : ["Noto Sans CJK SC", "Noto Serif CJK SC", "sans-serif"];
      rendered = candidates.find(available)!;
    }
    if (rendered !== first) node.style.fontFamily = JSON.stringify(rendered);
    selected.set(node, first);
    if (!resolutions.has(first))
      resolutions.set(first, {
        requestedFamily: first,
        selectedFamily: rendered,
        substituted: rendered !== first,
        missingCharacters: new Set(),
        glyphAuditComplete: true,
      });
  }
  // A local raster comparison also catches the generic missing-glyph box. This
  // is a preview diagnostic, not a promise of platform-independent font coverage.
  canvas.width = 128;
  canvas.height = 128;
  const hashes = new Map<string, number>();
  let audited = 0;
  const glyphHash = (font: string, character: string): number => {
    const key = font + "\0" + character;
    const cached = hashes.get(key);
    if (cached !== undefined) return cached;
    context.clearRect(0, 0, 128, 128);
    context.font = "64px " + font;
    context.fillText(character, 16, 90);
    const pixels = context.getImageData(0, 0, 128, 128).data;
    let hash = 2166136261;
    for (const value of pixels) hash = Math.imul(hash ^ value, 16777619);
    hashes.set(key, hash >>> 0);
    return hash >>> 0;
  };
  for (const node of nodes.filter((item) =>
    item.matches('[data-word-script="eastAsia"],[data-office-script="east-asia"],.sdk-number-marker,.sdk-ppt-marker'),
  )) {
    const resolution = resolutions.get(selected.get(node) ?? "");
    if (!resolution) continue;
    const font = document.defaultView!.getComputedStyle(node).fontFamily;
    for (const character of new Set(node.textContent ?? "")) {
      const key = font + "\0" + character;
      if (!hashes.has(key) && audited >= 4096) {
        resolution.glyphAuditComplete = false;
        continue;
      }
      if (!hashes.has(key)) audited++;
      if (
        glyphHash(font, character) === glyphHash(font, "\u{10ffff}") &&
        glyphHash(font, character) === glyphHash(font, "\uffff")
      )
        resolution.missingCharacters.add(character);
    }
  }
  return Object.freeze(
    [...resolutions.values()].map((item) =>
      Object.freeze({
        ...item,
        missingCharacters: Object.freeze([...item.missingCharacters]),
      }),
    ),
  );
}
