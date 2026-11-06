import { expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { openDocument } from "@doffi/core";
import { importXlsx } from "../src/xlsx-import.js";

it("preserves independently generated native text formats and rich values through public import and reopen", async () => {
  const evidence = JSON.parse(readFileSync("docs/evidence/xlsx-text-formats.json", "utf8")) as { sourceSha256: string; samples: { value: string; code: string; rich?: boolean }[] };
  const bytes = readFileSync("packages/office/tests/fixtures/xlsx-text-formats.xlsx");
  const regenerated = execFileSync("python3", ["-c", "import importlib.util,sys;s=importlib.util.spec_from_file_location('text_formats','scripts/check-xlsx-text-formats.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate())"]);
  expect(regenerated).toEqual(bytes); expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  const model = openDocument(await importXlsx(bytes, { documentId: "native-text-formats" }), { kind: "workbook" });
  const snapshot = model.snapshot(), formats = new Map(snapshot.metadata.workbook.numberFormats.map(format => [format.formatId, format]));
  const reopened = openDocument(model.serialize(), { kind: "workbook" });
  for (const [row, sample] of evidence.samples.entries()) {
    const cell = model.getCell("sheet-1", row, 0)!;
    expect(cell.value).toEqual({ type: "text", value: sample.value }); expect(cell.text).toBe(sample.value);
    expect(formats.get(cell.formatId!)).toMatchObject({ code: sample.code });
    if (sample.rich) expect(cell.html.match(/<span /g)).toHaveLength(2);
    expect(reopened.getCell("sheet-1", row, 0)).toEqual(cell);
  }
});
