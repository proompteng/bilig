import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

const require = createRequire(import.meta.url);
const glideEntry = require.resolve("@glideapps/glide-data-grid", {
  paths: [fileURLToPath(new URL("../../packages/grid", import.meta.url))],
});
const glidePackageRoot = resolve(dirname(glideEntry), "..", "..");

function includesAny(id: string, patterns: readonly string[]): boolean {
  const normalizedId = id.replaceAll("\\", "/");
  return patterns.some((pattern) => normalizedId.includes(pattern));
}

const codeSplittingGroups = [
  {
    name: "react-vendor",
    priority: 70,
    test(id: string) {
      return includesAny(id, [
        "/node_modules/react/",
        "/node_modules/react-dom/",
        "/node_modules/scheduler/",
      ]);
    },
  },
  {
    name: "sync-vendor",
    priority: 60,
    test(id: string) {
      return includesAny(id, [
        "/node_modules/@rocicorp/zero/",
        "/packages/zero-sync/",
        "/packages/crdt/",
      ]);
    },
  },
  {
    name: "grid-vendor",
    priority: 50,
    test(id: string) {
      return includesAny(id, [
        "/node_modules/@glideapps/glide-data-grid/",
        "/node_modules/marked/",
        "/node_modules/react-number-format/",
        "/node_modules/react-responsive-carousel/",
        "/node_modules/lodash/",
      ]);
    },
  },
  {
    name: "icons-vendor",
    priority: 40,
    test(id: string) {
      return includesAny(id, ["/node_modules/lucide-react/"]);
    },
  },
  {
    name: "formula-vendor",
    priority: 30,
    test(id: string) {
      return includesAny(id, ["/packages/formula/"]);
    },
  },
  {
    name: "engine-vendor",
    priority: 20,
    test(id: string) {
      return includesAny(id, [
        "/packages/binary-protocol/",
        "/packages/protocol/",
        "/packages/core/",
        "/packages/wasm-kernel/",
      ]);
    },
  },
  {
    name: "workbook-vendor",
    priority: 10,
    test(id: string) {
      return includesAny(id, [
        "/packages/grid/",
        "/packages/renderer/",
        "/packages/storage-browser/",
        "/packages/worker-transport/",
        "/packages/workbook-domain/",
        "/apps/web/src/WorkerWorkbookApp.tsx",
        "/apps/web/src/viewport-cache.ts",
        "/apps/web/src/worker-runtime.ts",
        "/apps/web/src/zero/",
      ]);
    },
  },
];

export default defineConfig({
  plugins: [react(), tailwindcss()],
  build: {
    rolldownOptions: {
      output: {
        codeSplitting: {
          groups: codeSplittingGroups,
        },
      },
    },
  },
  resolve: {
    alias: {
      "@glideapps/glide-data-grid/index.css": resolve(glidePackageRoot, "dist/index.css"),
      "@bilig/binary-protocol": fileURLToPath(
        new URL("../../packages/binary-protocol/src/index.ts", import.meta.url),
      ),
      "@bilig/formula/program-arena": fileURLToPath(
        new URL("../../packages/formula/src/program-arena.ts", import.meta.url),
      ),
      "@bilig/core": fileURLToPath(new URL("../../packages/core/src/index.ts", import.meta.url)),
      "@bilig/crdt": fileURLToPath(new URL("../../packages/crdt/src/index.ts", import.meta.url)),
      "@bilig/formula": fileURLToPath(
        new URL("../../packages/formula/src/index.ts", import.meta.url),
      ),
      "@bilig/grid": fileURLToPath(new URL("../../packages/grid/src/index.ts", import.meta.url)),
      "@bilig/protocol": fileURLToPath(
        new URL("../../packages/protocol/src/index.ts", import.meta.url),
      ),
      "@bilig/renderer": fileURLToPath(
        new URL("../../packages/renderer/src/index.ts", import.meta.url),
      ),
      "@bilig/storage-browser": fileURLToPath(
        new URL("../../packages/storage-browser/src/index.ts", import.meta.url),
      ),
      "@bilig/zero-sync": fileURLToPath(
        new URL("../../packages/zero-sync/src/index.ts", import.meta.url),
      ),
      "@bilig/wasm-kernel": fileURLToPath(
        new URL("../../packages/wasm-kernel/src/index.ts", import.meta.url),
      ),
    },
  },
});
