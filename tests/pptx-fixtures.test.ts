import { execFileSync } from "node:child_process";
import { expect, test } from "vitest";

test("independent PPTX corpus is reproducible and its native semantics match expectations", () => {
  expect(execFileSync("python3", ["scripts/check-pptx-fixtures.py"], {
    encoding: "utf8",
  })).toContain("Verified 16 deterministic independent PPTX sources");
});
