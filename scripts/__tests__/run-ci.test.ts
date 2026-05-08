import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('run-ci', () => {
  it('serializes generated checks and avoids pnpm for direct preflight gates', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/run-ci.ts'), 'utf8')

    expect(source).toContain('const generatedSourceChecks: readonly CiTask[] = [')
    expect(source).toContain("await runSequential('generated-source checks', generatedSourceChecks)")
    expect(source).toContain("await runSequential('static package build prerequisites'")
    expect(source).toContain("bunScript('protocol check', 'scripts/gen-protocol.ts', '--check')")
    expect(source).toContain(
      "direct('protocol package build for generated-source imports', workspaceBin('tsc'), '-p', 'packages/protocol/tsconfig.json')",
    )
    expect(source).toContain('const wasmBuildTask: CiTask = {')
    expect(source).toContain("directPackageScript('correctness public workbook corpus', 'test:correctness:corpus')")
    expect(source).not.toContain("pnpm('protocol check'")
    expect(source).not.toContain("pnpm('wasm build'")
    expect(source).not.toContain("pnpm('correctness public workbook corpus'")
    expect(source).not.toContain("await runStage('generated-source checks'")
    expect(source).not.toContain("await runStage('static package build prerequisites'")
  })
})
