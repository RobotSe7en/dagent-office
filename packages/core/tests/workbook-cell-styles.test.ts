import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "../src/index.js";
import type { WorkbookCellStyles, WorkbookDocumentMetadata, WorkbookManifest } from "../src/index.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const flags = { font: false, fill: true, border: false, alignment: true, numberFormat: false, protection: false };
function seeded(edit: (styles: WorkbookCellStyles) => void = () => {}, cellClass = "local") {
  const files = new Map(createWorkbook({ documentId: "named-styles" }));
  const metadata = JSON.parse(decoder.decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.cellStyles = {
    definitions: [{ styleId: "normal", cssClass: "cell-default", numberFormatId: "general", apply: { ...flags } }, { styleId: "accent", cssClass: "accent", numberFormatId: "general", apply: { ...flags } }],
    names: [{ name: "常规", styleId: "normal", builtinId: 0, hidden: false, customBuiltin: true }, { name: "汇总😀", styleId: "accent", builtinId: 1, level: 2, hidden: true, customBuiltin: false }],
    bindings: [{ cssClass: "local", styleId: "accent", apply: { ...flags } }],
  };
  edit(metadata.workbook.cellStyles);
  const sheet = metadata.workbook.sheets[0]!; sheet.rowDimensions = [{ from: 1, to: 1, cssClass: "local" }]; sheet.columnDimensions = [{ from: 1, to: 1, unit: "characters", cssClass: "local" }];
  files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")!) + '\n.accent{font-weight:bold}.local{font-weight:normal}.extra{color:red}'));
  const manifest = JSON.parse(decoder.decode(files.get("manifest.json")!)) as WorkbookManifest, path = "sheets/sheet-1/0-0.html";
  const bytes = encoder.encode(`<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="0" data-column="0"><tbody><tr data-row="0"><td data-column="0" data-value-type="number" data-value="12" class="${cellClass}" data-format-id="general">12</td></tr></tbody></table></body></html>`);
  files.set(path, bytes); manifest.chunks.push({ path, sheetId: "sheet-1", row: 0, column: 0, sha256: createHash("sha256").update(bytes).digest("hex") });
  files.set("manifest.json", encoder.encode(JSON.stringify(manifest))); return files;
}
describe("named workbook cell style sources", () => {
  it("retains definitions, localized names, update flags and existing cell/row/column associations", () => {
    const model = openDocument(seeded(), { kind: "workbook" }), snapshot = model.snapshot(), styles = snapshot.metadata.workbook.cellStyles!;
    expect(styles.names[1]).toEqual({ name: "汇总😀", styleId: "accent", builtinId: 1, level: 2, hidden: true, customBuiltin: false });
    expect(styles.bindings).toEqual([{ cssClass: "local", styleId: "accent", apply: flags }]);
    expect(model.getCell("sheet-1", 0, 0)).toMatchObject({ cssClass: "local", formatId: "general", value: { type: "number", value: 12 } });
    expect(snapshot.metadata.workbook.sheets[0]!.rowDimensions[0]!.cssClass).toBe("local");
    expect(snapshot.metadata.workbook.sheets[0]!.columnDimensions[0]!.cssClass).toBe("local");
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(snapshot);
    // Local appearance and a named definition may differ; propagation flags
    // never retroactively replace the cell's current CSS or numeric format.
    expect(snapshot.styles[0]!.css).toContain(".local{font-weight:normal}");
    expect(() => openDocument(seeded(() => {}, "local local extra"))).not.toThrow();
  });
  it("rejects missing or ambiguous definitions and class associations on every reopen", () => {
    const edits: ((styles: WorkbookCellStyles) => void)[] = [
      styles => { styles.definitions.push(styles.definitions[0]!); },
      styles => { styles.definitions[0]!.cssClass = "missing"; },
      styles => { styles.definitions[0]!.numberFormatId = "missing"; },
      styles => { styles.names[1]!.styleId = "missing"; },
      styles => { styles.names.push({ ...styles.names[1]!, name: "汇总😀" }); },
      styles => { styles.names[1]!.builtinId = 0; },
      styles => { styles.bindings[0]!.cssClass = "accent"; },
      styles => { styles.bindings[0]!.styleId = "missing"; },
      styles => { styles.bindings.push(styles.bindings[0]!); },
    ];
    for (const edit of edits) expect(() => openDocument(seeded(edit))).toThrow();
    expect(() => openDocument(seeded(() => {}, "local accent"))).toThrow();
    expect(() => openDocument(seeded(styles => { styles.names = Array.from({ length: 100 }, (_, i) => ({ ...styles.names[1]!, name: "样式" + i })); }), { limits: { maxNodes: 50 } })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
});
