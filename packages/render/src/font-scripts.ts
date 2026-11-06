export type OfficeFontRole = "ascii" | "h-ansi" | "east-asia" | "complex-script";
export interface OfficeFontSegment { readonly role: OfficeFontRole; readonly text: string; readonly from: number; readonly to: number }

/** Script selection never normalizes text. Combining marks and variation
 * selectors stay with the preceding base; offsets always count UTF-16 units. */
export function officeFontSegments(text: string, initial: OfficeFontRole = "ascii"): readonly OfficeFontSegment[] {
  const segments: { role: OfficeFontRole; text: string; from: number; to: number }[] = [];
  let previous = initial, offset = 0;
  for (const character of text) {
    const code = character.codePointAt(0)!;
    const role: OfficeFontRole = /\p{Mark}/u.test(character) || code >= 0xfe00 && code <= 0xfe0f || code >= 0xe0100 && code <= 0xe01ef ? previous
      : /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\u3000-\u303f\uff00-\uffef]/u.test(character) ? "east-asia"
        : code < 128 ? "ascii" : /[\p{Script=Arabic}\p{Script=Hebrew}]/u.test(character) ? "complex-script" : "h-ansi";
    const end = offset + character.length, last = segments.at(-1);
    if (last?.role === role) { last.text += character; last.to = end; } else segments.push({ role, text: character, from: offset, to: end });
    previous = role; offset = end;
  }
  return Object.freeze(segments.map(segment => Object.freeze(segment)));
}
