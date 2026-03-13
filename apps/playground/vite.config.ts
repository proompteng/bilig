import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      "@bilig/core": fileURLToPath(new URL("../../packages/core/src/index.ts", import.meta.url)),
      "@bilig/crdt": fileURLToPath(new URL("../../packages/crdt/src/index.ts", import.meta.url)),
      "@bilig/formula": fileURLToPath(new URL("../../packages/formula/src/index.ts", import.meta.url)),
      "@bilig/protocol": fileURLToPath(new URL("../../packages/protocol/src/index.ts", import.meta.url)),
      "@bilig/wasm-kernel": fileURLToPath(new URL("../../packages/wasm-kernel/src/index.ts", import.meta.url))
    }
  }
});
