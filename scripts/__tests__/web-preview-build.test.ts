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

  it('builds app runtime workspace dependencies before starting the app in run mode', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/run-dev-web-local.ts'), 'utf8')

    expect(source).toContain("['pnpm', '--filter', '@bilig/app^...', 'run', 'build']")
    expect(source).toContain("if (appServerMode === 'run')")
    expect(source).toContain('buildAppRuntimeDependencies()')
  })

  it('allows CI to reuse prebuilt app runtime dependencies for browser smoke', () => {
    const devSource = readFileSync(resolve(repoRoot, 'scripts/run-dev-web-local.ts'), 'utf8')
    const ciSource = readFileSync(resolve(repoRoot, 'scripts/run-ci.ts'), 'utf8')

    expect(devSource).toContain('resolveDevAppRuntimeBuildEnabled')
    expect(ciSource).toContain("pnpm('app runtime dependency build', '--filter', '@bilig/app^...', 'run', 'build')")
    expect(ciSource).toContain("BILIG_DEV_APP_RUNTIME_BUILD: '0'")
  })

  it('bounds lsof probes so local browser stack startup cannot hang indefinitely', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/run-dev-web-local.ts'), 'utf8')

    expect(source).toContain("import { spawnSync as nodeSpawnSync } from 'node:child_process'")
    expect(source).toContain("nodeSpawnSync('lsof'")
    expect(source).toContain('const localProcessProbeTimeoutMs = 1_000')
    expect(source).toContain('timeout: localProcessProbeTimeoutMs')
  })

  it('reuses the fresh CI-built preview bundle during browser tests', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/run-ci.ts'), 'utf8')
    const devSource = readFileSync(resolve(repoRoot, 'scripts/run-dev-web-local.ts'), 'utf8')
    const viteSource = readFileSync(resolve(repoRoot, 'apps/web/vite.config.ts'), 'utf8')

    expect(source).toContain("pnpm('browser web bundle build', '--filter', '@bilig/web', 'build:bundle')")
    expect(source).toContain("BILIG_DEV_WEB_PREVIEW_BUILD: '0'")
    expect(devSource).toContain('resolveDevWebPreviewBuildEnabled')
    expect(devSource).toContain('assertReusableWebPreviewBundleMatchesEnv')
    expect(viteSource).toContain('biligBuildMetadataPlugin')
  })

  it('declares a checked-in favicon so Browser QA starts without missing-asset noise', () => {
    const html = readFileSync(resolve(repoRoot, 'apps/web/index.html'), 'utf8')
    const favicon = readFileSync(resolve(repoRoot, 'apps/web/public/favicon.svg'), 'utf8')

    expect(html).toContain('<link rel="icon" href="/favicon.svg" type="image/svg+xml" />')
    expect(favicon).toContain('<svg')
  })
})
