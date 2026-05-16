import { existsSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('headless package workflow', () => {
  it('runs current headless test files instead of stale paths', () => {
    const source = readFileSync(resolve(repoRoot, '.github/workflows/headless-package.yml'), 'utf8')
    const testPaths = [...source.matchAll(/packages\/headless\/src\/__tests__\/[A-Za-z0-9.-]+\.test\.ts/g)].map((match) => match[0])
    const scriptTestPaths = [...source.matchAll(/scripts\/__tests__\/[A-Za-z0-9.-]+\.test\.ts/g)].map((match) => match[0])

    expect(testPaths).toContain('packages/headless/src/__tests__/work-paper.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/work-paper-runtime.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/work-paper-parity.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/workpaper-formula-regressions.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/workpaper-version-surface.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/sum-formula-members.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/sumifs-wildcard-arithmetic.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/xlookup-decimal-exact-match.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/persistence.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/persistence.fuzz.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/hyperformula-surface-parity.test.ts')
    expect(scriptTestPaths).toContain('scripts/__tests__/runtime-package-publish-validation.test.ts')

    for (const testPath of [...testPaths, ...scriptTestPaths]) {
      expect(existsSync(resolve(repoRoot, testPath)), `${testPath} should exist`).toBe(true)
    }
  })

  it('keeps publish, benchmark, and clean consumer smoke gates in the package workflow', () => {
    const source = readFileSync(resolve(repoRoot, '.github/workflows/headless-package.yml'), 'utf8')

    expect(source).toMatch(/['"]packages\/excel-import\/\*\*['"]/)
    expect(source).toContain('pnpm --filter @bilig/excel-import build')
    expect(source).toContain('packages/excel-import/package.json')
    expect(source).toContain('packages/excel-import')
    expect(source).toContain('allow_new_packages')
    expect(source).toContain('ALLOW_NEW_NPM_PACKAGES')
    expect(source).toContain('bun scripts/sync-runtime-release-metadata.ts')
    expect(source).toContain('runner needs runtime d.ts outputs rebuilt after the version sync')
    expect(source).toContain("group: runtime-packages-${{ github.event_name == 'workflow_dispatch' && github.run_id || github.sha }}")
    expect(source).toContain('cancel-in-progress: false')
    expect(source).toContain('reason=stale workflow SHA; newer main exists')
    expect(source).toContain('git push origin HEAD:main')
    expect(source).toContain('git push github HEAD:main')
    expect(source).toContain('GitHub main reached release metadata SHA during push')
    expect(source).toContain('git tag -a "${TAG_NAME}" -m "Libraries v${TARGET_VERSION}" HEAD')
    expect(source).toContain('pnpm publish:runtime:check')
    expect(source).toContain('pnpm workpaper:bench:competitive:check')
    expect(source).toContain('pnpm workpaper:smoke:external')
  })
})
