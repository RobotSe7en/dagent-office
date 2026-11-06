import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
function source(cells: string): Uint8Array {
  const files = new Map(readDocumentZip(readFileSync("examples/xlsx/04-formulas-names.xlsx")));
  const edit = (path: string, from: string, to: string) => files.set(path, encoder.encode(decoder.decode(files.get(path)!).replace(from, to)));
  edit("[Content_Types].xml", '</Types>', '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>');
  edit("xl/_rels/workbook.xml.rels", '</Relationships>', '<Relationship Id="chain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>');
  files.set("xl/calcChain.xml", encoder.encode('<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + cells + '</calcChain>'));
  return writeOpcZip(files);
}
describe("native calculation-chain hints", () => {
  it("checks inherited sheet identities against actual formulas without treating order as dependencies", async () => {
    const result = await readXlsxRecords(source('<c r="B3" i="1"/><c r="B2" s="1" l="1" t="1"/><c r="B3" a="1"/>'));
    // A stale chain array hint cannot override the actual scalar cell formula.
    expect(result.counts.formulas).toBeGreaterThan(2);
  });
  it("rejects unknown sheets, absent cells, non-formula cells, unknown extensions and invalid flags", async () => {
    for (const cells of ['', '<c r="B2"/>', '<c r="B2" i="99"/>', '<c r="XFD1048576" i="1"/>', '<c r="A2" i="1"/>', '<c r="B2" i="1" t="maybe"/>', '<c r="B2" i="1"/><extLst/>', '<c r="B2" i="1" unknown="1"/>']) {
      await expect(readXlsxRecords(source(cells)), cells).rejects.toBeDefined();
    }
  });
  it("counts repeated cache entries against the cumulative formula budget", async () => {
    await expect(readXlsxRecords(source('<c r="B2" i="1"/>' + '<c r="B2"/>'.repeat(20)), { workbookLimits: { maxFormulas: 20 } })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  });
});
