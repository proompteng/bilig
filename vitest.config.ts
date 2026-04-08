import { existsSync, readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const rootDir = fileURLToPath(new URL(".", import.meta.url));
const packagesDir = join(rootDir, "packages");

function createWorkspacePackageAliases() {
  return readdirSync(packagesDir, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .flatMap((entry) => {
      const packageDir = join(packagesDir, entry.name);
      const packageJsonPath = join(packageDir, "package.json");
      const sourceEntryPath = join(packageDir, "src", "index.ts");
      if (!existsSync(packageJsonPath) || !existsSync(sourceEntryPath)) {
        return [];
      }
      const packageJson = JSON.parse(readFileSync(packageJsonPath, "utf8")) as {
        name?: unknown;
      };
      if (typeof packageJson.name !== "string" || !packageJson.name.startsWith("@bilig/")) {
        return [];
      }
      return [{ find: packageJson.name, replacement: sourceEntryPath }];
    })
    .toSorted((left, right) => left.find.localeCompare(right.find));
}

const workspacePackageAliases = createWorkspacePackageAliases();

export default defineConfig({
  resolve: {
    alias: [
      {
        find: "@bilig/formula/program-arena",
        replacement: `${rootDir}packages/formula/src/program-arena.ts`,
      },
      ...workspacePackageAliases,
    ],
  },
  test: {
    environment: "node",
    include: [
      "packages/*/src/**/*.test.ts",
      "packages/*/src/**/*.test.tsx",
      "apps/*/src/**/*.test.ts",
      "apps/*/src/**/*.test.tsx",
    ],
    exclude: ["**/dist/**", "**/build/**"],
    coverage: {
      provider: "v8",
      reporter: ["text", "lcov", "json", "json-summary"],
      include: [
        "packages/core/src/**/*.ts",
        "packages/formula/src/**/*.ts",
        "packages/renderer/src/**/*.ts",
      ],
      exclude: [
        "**/__tests__/**",
        "**/*.d.ts",
        "packages/core/src/index.ts",
        "packages/core/src/snapshot.ts",
        "packages/formula/src/index.ts",
        "packages/formula/src/ast.ts",
        "packages/renderer/src/index.ts",
      ],
      thresholds: {
        lines: 90,
        statements: 90,
        functions: 90,
        branches: 70,
      },
    },
  },
});
