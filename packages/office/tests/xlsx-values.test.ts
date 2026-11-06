import { DEFAULT_DOCUMENT_READ_LIMITS } from "@doffi/core";
import { expect, it } from "vitest";
import { address, range, spreadsheetText, textHtml } from "../src/xlsx/values.js";
import { xml } from "../src/ooxml/xml.js";

const node = xml(new TextEncoder().encode('<t xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'), "text.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 });
it("decodes SpreadsheetML escapes once while preserving valid UTF-16 and literal escape text", () => {
  expect(spreadsheetText("_x005F_x0041_ _x0041_ _xD83D__xDE00_ 禰󠄀_x000D__x000A_", node)).toBe("_x0041_ A 😀 禰󠄀\r\n");
  expect(spreadsheetText("_x0041_ _X0042_ _x00Af_ _x005F_x0041_", node)).toBe("A _X0042_ ¯ _x0041_");
  expect(textHtml('a\rb\n<c>&"')).toBe('a&#13;b\n&lt;c&gt;&amp;&quot;');
  for (const source of ["_xD800_", "_xDC00_", "_xD800_a", "_x0000_", "_xFFFE_"]) expect(() => spreadsheetText(source, node)).toThrow();
  expect(() => spreadsheetText("😀", node, 1)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
});
it("maps native cell/range addresses within exact bounds without expanding sparse axes", () => {
  expect(address("$XFD$1048576", node)).toEqual({ row: 1048575, column: 16383 });
  expect(range("B2:D4", node)).toEqual({ from: { row: 1, column: 1 }, to: { row: 3, column: 3 } });
  for (const source of ["A0", "XFE1", "A1048577", "A1:B2", "A01", "[book]A1"]) expect(() => address(source, node)).toThrow();
  expect(() => range("D4:B2", node)).toThrow();
});
