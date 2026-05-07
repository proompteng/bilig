import { existsSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('headless package workflow', () => {
  it('runs current headless test files instead of stale paths', () => {
    const source = readFileSync(resolve(repoRoot, '.github/workflows/headless-package.yml'), 'utf8')
    const testPaths = [...source.matchAll(/packages\/headless\/src\/__tests__\/[A-Za-z0-9.-]+\.test\.ts/g)].map((match) => match[0])

    expect(testPaths).toContain('packages/headless/src/__tests__/work-paper.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/work-paper-runtime.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/work-paper-parity.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/github-issues.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/github-issue-124-sumifs-wildcard-arithmetic.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/github-issue-125-xlookup-decimal.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/persistence.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/persistence.fuzz.test.ts')
    expect(testPaths).toContain('packages/headless/src/__tests__/hyperformula-surface-parity.test.ts')

    for (const testPath of testPaths) {
      expect(existsSync(resolve(repoRoot, testPath)), `${testPath} should exist`).toBe(true)
    }
  })

  it('keeps publish, benchmark, and clean consumer smoke gates in the package workflow', () => {
    const source = readFileSync(resolve(repoRoot, '.github/workflows/headless-package.yml'), 'utf8')

    expect(source).toContain('"packages/excel-import/**"')
    expect(source).toContain('pnpm --filter @bilig/excel-import build')
    expect(source).toContain('packages/excel-import/package.json')
    expect(source).toContain('packages/excel-import')
    expect(source).toContain('allow_new_packages')
    expect(source).toContain('ALLOW_NEW_NPM_PACKAGES')
    expect(source).toContain('pnpm publish:runtime:check')
    expect(source).toContain('pnpm workpaper:bench:competitive:check')
    expect(source).toContain('pnpm workpaper:smoke:external')
  })
})
