import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { resolveCiProfile, resolveCiSkipBrowserGates } from '../run-ci-config.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('run-ci', () => {
  it('defaults to the fast CI profile and accepts explicit profiles', () => {
    expect(resolveCiProfile({})).toBe('fast')
    expect(resolveCiProfile({ BILIG_CI_PROFILE: 'fast' })).toBe('fast')
    expect(resolveCiProfile({ BILIG_CI_PROFILE: 'full' })).toBe('full')
  })

  it('rejects malformed CI profiles instead of silently downgrading gates', () => {
    expect(() => resolveCiProfile({ BILIG_CI_PROFILE: 'ful' })).toThrow('BILIG_CI_PROFILE must be "fast" or "full", got ful')
  })

  it('resolves the browser gate skip flag strictly', () => {
    expect(resolveCiSkipBrowserGates({})).toBe(false)
    expect(resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: '1' })).toBe(true)
    expect(resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: 'true' })).toBe(true)
    expect(resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: '0' })).toBe(false)
    expect(resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: 'false' })).toBe(false)
    expect(() => resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: 'yes' })).toThrow(
      'BILIG_CI_SKIP_BROWSER must be "1", "true", "0", or "false" when set, got yes',
    )
  })

  it('serializes generated checks and avoids pnpm for direct preflight gates', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/run-ci.ts'), 'utf8')

    expect(source).toContain('const generatedSourceChecks: readonly CiTask[] = [')
    expect(source).toContain("await runSequential('generated-source checks', generatedSourceChecks)")
    expect(source).toContain("const semanticFastGate = pnpm('semantic correctness fast gate', 'test:semantic:fast')")
    expect(source).toContain("await runSequential('semantic correctness checks', [semanticFastGate])")
    expect(source).toContain("await runSequential('static package build prerequisites'")
    expect(source).toContain("bunScript('protocol check', 'scripts/gen-protocol.ts', '--check')")
    expect(source).toContain(
      "direct('protocol package build for generated-source imports', workspaceBin('tsc'), '-p', 'packages/protocol/tsconfig.json')",
    )
    expect(source).toContain('const wasmBuildTask: CiTask = {')
    expect(source).toContain("directPackageScript('correctness public workbook corpus', 'test:correctness:corpus')")
    expect(source).toContain(
      "withEnv(directPackageScript('correctness core', 'test:correctness:core'), { BILIG_VITEST_FILE_CHUNK_SIZE: '10' })",
    )
    expect(source).toContain(
      "withEnv(directPackageScript('correctness formula', 'test:correctness:formula'), { BILIG_VITEST_FILE_CHUNK_SIZE: '3' })",
    )
    expect(source).toContain("BILIG_VITEST_FILE_CHUNK_SIZE: '10'")
    expect(source).toContain(
      "directPackageScript('financial public workbook corpus resume check', 'public-workbook-corpus:resume-financial:check')",
    )
    expect(source).toContain(
      "directPackageScript('public workbook corpus synthetic memory gate', 'public-workbook-corpus:memory-gate:synthetic')",
    )
    expect(source).toContain("bunScript('agent discovery docs check', 'scripts/sync-agent-discovery-docs.ts', '--check')")
    expect(source).toContain("await runSequential('static direct checks'")
    expect(source).not.toContain("pnpm('protocol check'")
    expect(source).not.toContain("pnpm('wasm build'")
    expect(source).not.toContain("pnpm('correctness public workbook corpus'")
    expect(source).not.toContain("await runStage('generated-source checks'")
    expect(source).not.toContain("await runStage('static package build prerequisites'")
    expect(source).not.toContain("await runStage('static direct checks'")
  })

  it('runs the CI orchestrator through tsx instead of bun', () => {
    const packageJson = readFileSync(resolve(repoRoot, 'package.json'), 'utf8')

    expect(packageJson).toContain('"ci": "BILIG_CI_PROFILE=fast BILIG_CI_SKIP_BROWSER=1 tsx scripts/run-ci.ts"')
    expect(packageJson).toContain('"ci:core": "BILIG_CI_PROFILE=fast BILIG_CI_SKIP_BROWSER=1 tsx scripts/run-ci.ts"')
    expect(packageJson).toContain('"ci:full": "BILIG_CI_PROFILE=full tsx scripts/run-ci.ts"')
    expect(packageJson).toContain(
      '"public-workbook-corpus:memory-gate": "bun scripts/public-workbook-corpus-memory-gate.ts --require-public"',
    )
    expect(packageJson).toContain(
      '"public-workbook-corpus:memory-gate:synthetic": "bun scripts/public-workbook-corpus-memory-gate.ts --synthetic-only"',
    )
  })

  it('guards broad pre-push lint through the same resource gate', () => {
    const packageJson = readFileSync(resolve(repoRoot, 'package.json'), 'utf8')
    const prePushSource = readFileSync(resolve(repoRoot, 'scripts/run-pre-push.ts'), 'utf8')

    expect(packageJson).toContain('"hooks:pre-push": "tsx scripts/run-pre-push.ts"')
    expect(prePushSource).toContain("assertLocalCiResourceGuardAllowsRun(rootDir, process.env, { runLabel: 'pre-push lint' })")
    expect(prePushSource).toContain("await run('pnpm', ['lint'])")
  })
})
