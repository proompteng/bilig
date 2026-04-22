import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

function isStringRecord(value: unknown): value is Record<string, string> {
  return typeof value === 'object' && value !== null && Object.values(value).every((entry) => typeof entry === 'string')
}

function readWebPackageScripts(): Record<string, string> {
  const parsed = JSON.parse(readFileSync(resolve(repoRoot, 'apps/web/package.json'), 'utf8')) as unknown
  if (typeof parsed !== 'object' || parsed === null || !('scripts' in parsed) || !isStringRecord(parsed.scripts)) {
    throw new Error('apps/web/package.json must define string scripts.')
  }
  return parsed.scripts
}

describe('web preview build gate', () => {
  it('builds the web app through project references for fresh CI checkouts', () => {
    const scripts = readWebPackageScripts()

    expect(scripts['build']).toMatch(/^tsc -b tsconfig\.json\b/)
    expect(scripts['build']).not.toContain('tsc -p')
  })

  it('ensures the wasm kernel artifact before the preview web-server build', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/run-dev-web-local.ts'), 'utf8')

    expect(source).toContain("import { ensureWasmKernelArtifact } from './ensure-wasm-kernel.js'")
    expect(source).toContain('ensureWasmKernelArtifact()')
  })
})
