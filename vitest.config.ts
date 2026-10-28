import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["tests/**/*.test.ts", "packages/*/tests/**/*.test.ts"],
    maxWorkers: 2,
    testTimeout: 120_000,
    hookTimeout: 120_000,
    clearMocks: true,
  },
});
