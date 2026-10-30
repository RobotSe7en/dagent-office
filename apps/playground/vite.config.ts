import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  plugins: [react()],
  publicDir: fileURLToPath(new URL("../../examples/docx/", import.meta.url)),
  build: { target: "es2022" },
});
