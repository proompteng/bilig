import { join } from 'node:path'
import { defineConfig } from 'vitest/config'
import { createVitestAliasEntries, workspaceRootDir } from './scripts/workspace-resolution.js'

const workspacePackageAliases = createVitestAliasEntries([
  {
    find: '@bilig/formula/program-arena',
    replacement: join(workspaceRootDir, 'packages/formula/src/program-arena.ts'),
  },
])

export default defineConfig({
  resolve: {
    alias: workspacePackageAliases,
  },
  test: {
    environment: 'node',
    globalSetup: join(workspaceRootDir, 'scripts/vitest-global-setup.ts'),
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
      reporter: ['text', 'lcov', 'json', 'json-summary'],
      include: ['packages/core/src/**/*.ts', 'packages/formula/src/**/*.ts', 'packages/renderer/src/**/*.ts'],
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
        'packages/renderer/src/index.ts',
      ],
      thresholds: {
        lines: 91,
        statements: 91,
        functions: 91,
        branches: 70,
      },
    },
  },
})
