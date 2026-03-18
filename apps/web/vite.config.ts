import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const require = createRequire(import.meta.url);
const glideEntry = require.resolve("@glideapps/glide-data-grid", {
  paths: [fileURLToPath(new URL("../../packages/grid", import.meta.url))]
});
const glidePackageRoot = resolve(dirname(glideEntry), "..", "..");

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      "@glideapps/glide-data-grid/index.css": resolve(glidePackageRoot, "dist/index.css"),
      "@bilig/formula/program-arena": fileURLToPath(new URL("../../packages/formula/src/program-arena.ts", import.meta.url)),
      "@bilig/core": fileURLToPath(new URL("../../packages/core/src/index.ts", import.meta.url)),
      "@bilig/crdt": fileURLToPath(new URL("../../packages/crdt/src/index.ts", import.meta.url)),
      "@bilig/formula": fileURLToPath(new URL("../../packages/formula/src/index.ts", import.meta.url)),
      "@bilig/grid": fileURLToPath(new URL("../../packages/grid/src/index.ts", import.meta.url)),
      "@bilig/protocol": fileURLToPath(new URL("../../packages/protocol/src/index.ts", import.meta.url)),
      "@bilig/renderer": fileURLToPath(new URL("../../packages/renderer/src/index.ts", import.meta.url)),
      "@bilig/storage-browser": fileURLToPath(new URL("../../packages/storage-browser/src/index.ts", import.meta.url)),
      "@bilig/wasm-kernel": fileURLToPath(new URL("../../packages/wasm-kernel/src/index.ts", import.meta.url))
    }
  }
});
