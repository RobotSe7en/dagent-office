import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { evaluatePresentationPreset, openDocument, readPresentationGeometry } from "@dagent-office/core";
import { importPptx } from "@dagent-office/office";

test("evaluates imported native shape/connector adjustments from independent and Office-resaved PPTX", () => {
  for (const producer of ["", "producer/"]) for (const name of ["04-shapes.pptx", "06-connectors.pptx"]) {
    const model = openDocument(importPptx(readFileSync(new URL("../examples/pptx/" + producer + name, import.meta.url)), { documentId: "geometry" }));
    for (const object of model.snapshot().metadata.presentation!.objects) {
      if (object.kind !== "geometry" && object.kind !== "connector") continue;
      const node = model.getNode(object.nodeId)!, geometry = readPresentationGeometry(node.attributes)!;
      const adjustments = Object.fromEntries(Object.entries(node.attributes).filter(([key]) => key.startsWith("data-office-adjust-")).map(([key, value]) => [key.slice("data-office-adjust-".length), Number(value)]));
      const shape = evaluatePresentationPreset(object.kind === "geometry" ? object.preset : node.attributes["data-office-connector-preset"]!, geometry.width, geometry.height, adjustments);
      expect(shape.paths).not.toHaveLength(0);
      for (const connection of shape.connections) expect(Number.isFinite(connection.x + connection.y)).toBe(true);
    }
  }
});

test("shipped geometry data and the browser renderer retain the complete normative data license", () => {
  const notice = readFileSync(new URL("../packages/core/ECMA-376-GEOMETRY-LICENSE.txt", import.meta.url), "utf8");
  for (const file of ["../packages/core/dist/presentation-geometry-data.js", "../packages/render/dist/browser-script.js"]) {
    const output = readFileSync(new URL(file, import.meta.url), "utf8");
    for (const line of notice.split("\n").filter(Boolean)) expect(output).toContain(line);
  }
});
