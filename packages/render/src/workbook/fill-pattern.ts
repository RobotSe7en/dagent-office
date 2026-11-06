/** ECMA-376 Part 1 §18.18.55 illustrates binary repeating cell patterns.
 * One logical dot is a 96-dpi device-independent pixel (0.75 document points).
 * Masks describe pattern geometry, never a bitmap resource or source color. */
const patterns = {
  mediumGray: ["10", "01"], darkGray: ["0111", "1101"], lightGray: ["0100", "0001"],
  darkHorizontal: ["1", "1", "0", "0"], darkVertical: ["1100"],
  darkDown: ["1100", "0110", "0011", "1001"], darkUp: ["1100", "1001", "0011", "0110"],
  darkGrid: ["1100", "1100", "0011", "0011"], darkTrellis: ["1100", "1111", "0011", "1111"],
  lightHorizontal: ["1", "0", "0", "0"], lightVertical: ["1000"],
  lightDown: ["1000", "0100", "0010", "0001"], lightUp: ["0100", "1000", "0001", "0010"],
  lightGrid: ["1111", "1000", "1000", "1000"], lightTrellis: ["1010", "0100", "1010", "0001"],
  gray125: ["0000", "0100", "0000", "0001"], gray0625: ["00000000", "00000001", "00000000", "00010000"],
} as const;
for (const rows of Object.values(patterns)) Object.freeze(rows);
Object.freeze(patterns);
export type WorkbookFillPattern = keyof typeof patterns;
export function isWorkbookFillPattern(kind: string): kind is WorkbookFillPattern { return Object.hasOwn(patterns, kind); }
export function workbookFillPattern(kind: WorkbookFillPattern): readonly string[] { return patterns[kind]; }
