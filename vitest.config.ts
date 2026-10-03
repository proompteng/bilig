import { join } from 'node:path'
import { defineConfig } from 'vitest/config'
import { createVitestAliasEntries, workspaceRootDir } from './scripts/workspace-resolution.js'

const workspacePackageAliases = createVitestAliasEntries()

function parsePositiveInteger(value: string | undefined): number | undefined {
  if (!value) {
    return undefined
  }
  const parsed = Number.parseInt(value, 10)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined
}

function resolveTestTimeoutMs(): number | undefined {
  const explicitTimeout = parsePositiveInteger(process.env['BILIG_VITEST_TEST_TIMEOUT_MS'])
  if (explicitTimeout !== undefined) {
    return explicitTimeout
  }
  if (!process.env['BILIG_FUZZ_PROFILE'] && !process.env['BILIG_FUZZ_REPLAY']) {
    return undefined
  }
  return 120_000
}

export default defineConfig({
  resolve: {
    alias: workspacePackageAliases,
  },
  test: {
    environment: 'node',
    globalSetup: join(workspaceRootDir, 'scripts/vitest-global-setup.ts'),
    setupFiles: [join(workspaceRootDir, 'scripts/vitest-setup.ts')],
    testTimeout: resolveTestTimeoutMs(),
    include: [
      'packages/*/src/**/*.test.ts',
      'packages/*/src/**/*.test.tsx',
      'apps/*/src/**/*.test.ts',
      'apps/*/src/**/*.test.tsx',
      'scripts/**/*.test.ts',
    ],
    exclude: ['**/dist/**', '**/build/**'],
    coverage: {
      provider: 'v8',
      reportsDirectory: process.env['BILIG_COVERAGE_DIR'] ?? './coverage',
      reporter: ['text', 'lcov', 'json', 'json-summary'],
      include: ['packages/core/src/**/*.ts', 'packages/formula/src/**/*.ts'],
      exclude: [
        '**/__tests__/**',
        '**/*.d.ts',
        'packages/core/src/index.ts',
        'packages/core/src/snapshot.ts',
        'packages/formula/src/index.ts',
        'packages/formula/src/ast.ts',
        'packages/formula/src/js-evaluator-types.ts',
        '**/packages/formula/src/js-evaluator-types.ts',
        '**/js-evaluator-types.ts',
      ],
      thresholds: {
        // Package line coverage is enforced by scripts/coverage-contracts.ts after
        // the V8 report is written; keep Vitest responsible for the global shape
        // metrics it can enforce without blocking that package-aware contract.
        functions: 91,
        branches: 70,
      },
    },
  },
})
