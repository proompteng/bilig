import { readdirSync, readFileSync, statSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

const criticalFuzzRoots = [
  'apps/bilig/src/codex-app',
  'apps/web/src',
  'packages/agent-api/src',
  'packages/contracts/src',
  'packages/core/src',
  'packages/excel-import/src',
  'packages/formula/src',
  'packages/grid/src',
  'packages/headless/src',
  'packages/protocol/src',
  'packages/renderer/src',
  'packages/worker-transport/src',
  'packages/zero-sync/src',
] as const

const forbiddenMarkers = [
  '@fuzz-browser',
  'BILIG_FUZZ_SKIP_BROWSER',
  'BILIG_BROWSER_INCLUDE_FUZZ',
  'BILIG_FUZZ_BROWSER',
  'test:fuzz:main',
  'test:fuzz:nightly',
  'test:fuzz:replay',
  'test:fuzz:promote',
] as const

describe('fuzz inventory guardrails', () => {
  it('keeps exactly one package fuzz entrypoint', () => {
    const packageJson = JSON.parse(readFileSync(join(repoRoot, 'package.json'), 'utf8')) as unknown
    const scripts = isRecord(packageJson) && isRecord(packageJson.scripts) ? packageJson.scripts : {}

    expect(Object.keys(scripts).filter((scriptName) => scriptName === 'test:fuzz' || scriptName.startsWith('test:fuzz:'))).toEqual([
      'test:fuzz',
    ])
  })

  it('keeps browser fuzz and fuzz variants out of the repo wiring', () => {
    const text = [
      readFileSync(join(repoRoot, 'package.json'), 'utf8'),
      ...listTextFiles(join(repoRoot, 'scripts')),
      ...listTextFiles(join(repoRoot, 'e2e/tests')),
    ].join('\n')

    for (const marker of forbiddenMarkers) {
      expect(text).not.toContain(marker)
    }
  })

  it('keeps every critical correctness surface under direct fuzz coverage', () => {
    const missing = criticalFuzzRoots.filter((root) => listFiles(join(repoRoot, root)).every((file) => !file.endsWith('.fuzz.test.ts')))

    expect(missing).toEqual([])
  })
})

// Helpers

function listTextFiles(root: string): string[] {
  return listFiles(root)
    .filter((file) => /\.(?:ts|tsx|js|mjs|json)$/u.test(file) && !file.endsWith('scripts/__tests__/fuzz-inventory.test.ts'))
    .map((file) => readFileSync(file, 'utf8'))
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function listFiles(root: string): string[] {
  const files: string[] = []
  for (const entry of readdirSync(root)) {
    const path = join(root, entry)
    const stat = statSync(path)
    if (stat.isDirectory()) {
      files.push(...listFiles(path))
    } else if (stat.isFile()) {
      files.push(path)
    }
  }
  return files
}
