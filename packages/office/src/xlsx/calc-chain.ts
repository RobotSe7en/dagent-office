import { DocumentError } from "@doffi/core";
import type { WorkbookCoordinate } from "@doffi/core";
import type { OpcPackage } from "../ooxml/package.js";
import { check, elements, fail, S } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { address, attributes, boolean, integer, value } from "./values.js";
import type { NativeWorkbook } from "./workbook.js";

const key = (sheet: number, cell: WorkbookCoordinate): string => `${sheet}:${cell.row}:${cell.column}`;
/** Native calculation order is a disposable optimization, not dependency data
 * or proof that cached formula results are current. Validate its references
 * against the streaming source before discarding the old scheduling hints. */
export class NativeCalculationChain {
  readonly part: string | undefined;
  private readonly pending = new Set<string>();
  private readonly root: XmlNode | undefined;
  constructor(opc: OpcPackage, workbook: NativeWorkbook, maxItems: number) {
    this.part = opc.related(workbook.part, "calcChain");
    if (!this.part) return;
    const root = this.root = opc.read(this.part);
    if (root.uri !== S || root.name !== "calcChain") fail(root, "The calculation-chain root is invalid.");
    attributes(root, []); check(root, ["c"], S);
    const entries = elements(root, "c", S), sheets = new Set(workbook.sheets.map(sheet => sheet.nativeId));
    if (!entries.length) fail(root, "A native calculation chain has no cells.");
    if (entries.length > maxItems) throw new DocumentError("ResourceLimitExceeded", "The native calculation chain exceeds its formula budget.");
    let sheetId = 0;
    for (const node of entries) {
      attributes(node, ["r", "i", "s", "l", "t", "a"]); check(node, [], S);
      if (value(node, "i") !== undefined) sheetId = integer(node, "i", 1, 0x7fff_ffff);
      if (!sheets.has(sheetId)) fail(node, "A calculation-chain cell has no existing worksheet identity.");
      for (const name of ["s", "l", "t", "a"]) boolean(node, name);
      this.pending.add(key(sheetId, address(value(node, "r") ?? "", node)));
    }
  }
  formula(sheetId: number, cell: WorkbookCoordinate): void { if (this.pending.size) this.pending.delete(key(sheetId, cell)); }
  finish(): void { if (this.pending.size) fail(this.root!, "A calculation-chain entry does not refer to a formula cell."); }
}
