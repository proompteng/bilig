import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildVitestArgBatches, buildVitestArgs, readVitestBatchCooldownMs, resolveVitestBin } from '../run-vitest.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readPackageScripts(packageJsonPath: string): Record<string, string> {
  const manifest: unknown = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  if (!isRecord(manifest) || !isRecord(manifest.scripts)) {
    throw new Error(`${packageJsonPath} must define package scripts`)
  }
  for (const [name, script] of Object.entries(manifest.scripts)) {
    if (typeof script !== 'string') {
      throw new Error(`${packageJsonPath} script ${name} must be a string`)
    }
  }
  return manifest.scripts
}

describe('run-vitest wrapper arguments', () => {
  it('bounds Vitest run mode outside CI so pnpm test uses the stable runner path', () => {
    expect(buildVitestArgs(['--run'], {})).toEqual([
      '--run',
      '--maxWorkers',
      '1',
      '--pool',
      'forks',
      '--configLoader',
      'runner',
      '--reporter',
      'verbose',
    ])
  })

  it('leaves watch mode raw outside CI', () => {
    expect(buildVitestArgs([], {})).toEqual([])
  })

  it('bounds Vitest workers in CI by default', () => {
    expect(buildVitestArgs(['--run', 'sample.test.ts'], { BILIG_CI_PROFILE: 'fast' })).toEqual([
      '--run',
      'sample.test.ts',
      '--maxWorkers',
      '1',
      '--pool',
      'forks',
      '--configLoader',
      'runner',
      '--reporter',
      'verbose',
    ])
  })

  it('bounds focused local Vitest run invocations without splitting them', () => {
    expect(buildVitestArgBatches(['--run', 'sample.test.ts'], {})).toEqual([
      ['--run', 'sample.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
    ])
  })

  it('preserves an explicit maxWorkers flag', () => {
    expect(buildVitestArgs(['--run', '--maxWorkers=1'], {})).toEqual([
      '--run',
      '--maxWorkers=1',
      '--pool',
      'forks',
      '--configLoader',
      'runner',
      '--reporter',
      'verbose',
    ])
  })

  it('allows CI worker limit overrides', () => {
    expect(
      buildVitestArgs(['--run'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_MAX_WORKERS: '3',
      }),
    ).toEqual(['--run', '--maxWorkers', '3', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'])
  })

  it('ignores malformed CI worker limit overrides instead of forwarding them', () => {
    expect(
      buildVitestArgs(['--run'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_MAX_WORKERS: '3abc',
      }),
    ).toEqual(['--run', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'])

    expect(
      buildVitestArgs(['--run'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_MAX_WORKERS: '0',
      }),
    ).toEqual(['--run', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'])
  })

  it('uses the verbose reporter in CI unless a reporter is already explicit', () => {
    expect(buildVitestArgs(['--run', 'sample.test.ts', '--reporter=dot'], { BILIG_CI_PROFILE: 'fast' })).toEqual([
      '--run',
      'sample.test.ts',
      '--reporter=dot',
      '--maxWorkers',
      '1',
      '--pool',
      'forks',
      '--configLoader',
      'runner',
    ])
  })

  it('splits large CI run file lists into serial batches', () => {
    const files = Array.from({ length: 7 }, (_, index) => `test-${index + 1}.test.ts`)

    expect(
      buildVitestArgBatches(['--run', ...files], {
        BILIG_CI_PROFILE: 'fast',
      }),
    ).toEqual([
      ['--run', ...files.slice(0, 3), '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
      ['--run', ...files.slice(3, 6), '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
      ['--run', files[6], '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
    ])
  })

  it('allows CI file chunk size overrides', () => {
    expect(
      buildVitestArgBatches(['--run', 'a.test.ts', 'b.test.ts', 'c.test.ts'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '2',
      }),
    ).toEqual([
      ['--run', 'a.test.ts', 'b.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
      ['--run', 'c.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
    ])
  })

  it('splits default all-file runs into bounded batches', () => {
    const files = Array.from({ length: 18 }, (_, index) => `packages/core/src/__tests__/all-${index + 1}.test.ts`)

    expect(buildVitestArgBatches(['--run'], {}, files)).toEqual([
      ['--run', ...files.slice(0, 16), '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
      ['--run', ...files.slice(16), '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
    ])
  })

  it('allows all-file run chunk size overrides', () => {
    const files = ['a.test.ts', 'b.test.ts', 'c.test.ts', 'd.test.ts', 'e.test.ts']

    expect(buildVitestArgBatches(['--run'], { BILIG_VITEST_FILE_CHUNK_SIZE: '2' }, files)).toEqual([
      ['--run', 'a.test.ts', 'b.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
      ['--run', 'c.test.ts', 'd.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
      ['--run', 'e.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
    ])
  })

  it('does not mix jsdom and node environment files in one CI batch', () => {
    expect(
      buildVitestArgBatches(
        [
          '--run',
          'apps/web/src/__tests__/workbook-editor-conflict.test.tsx',
          'apps/web/src/__tests__/worker-runtime-state.test.ts',
          'apps/web/src/__tests__/worker-workbook-app-model.test.ts',
        ],
        { BILIG_CI_PROFILE: 'fast' },
      ),
    ).toEqual([
      [
        '--run',
        'apps/web/src/__tests__/workbook-editor-conflict.test.tsx',
        '--maxWorkers',
        '1',
        '--pool',
        'forks',
        '--configLoader',
        'runner',
        '--reporter',
        'verbose',
      ],
      [
        '--run',
        'apps/web/src/__tests__/worker-runtime-state.test.ts',
        'apps/web/src/__tests__/worker-workbook-app-model.test.ts',
        '--maxWorkers',
        '1',
        '--pool',
        'forks',
        '--configLoader',
        'runner',
        '--reporter',
        'verbose',
      ],
    ])
  })

  it('ignores malformed CI file chunk size overrides instead of truncating them', () => {
    const files = ['a.test.ts', 'b.test.ts', 'c.test.ts', 'd.test.ts']

    expect(
      buildVitestArgBatches(['--run', ...files], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '2abc',
      }),
    ).toEqual([
      [
        '--run',
        'a.test.ts',
        'b.test.ts',
        'c.test.ts',
        '--maxWorkers',
        '1',
        '--pool',
        'forks',
        '--configLoader',
        'runner',
        '--reporter',
        'verbose',
      ],
      ['--run', 'd.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
    ])
  })

  it('does not split run arguments that include flags', () => {
    expect(
      buildVitestArgBatches(['--run', 'sample.test.ts', '--reporter=dot'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '1',
      }),
    ).toEqual([['--run', 'sample.test.ts', '--reporter=dot', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner']])
  })

  it('adds a short CI-only cooldown between split batches', () => {
    expect(readVitestBatchCooldownMs({})).toBe(0)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast' })).toBe(1000)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '0' })).toBe(0)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '2500' })).toBe(2500)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '2500ms' })).toBe(1000)
  })

  it('resolves the Vitest binary from the workspace root', () => {
    expect(resolveVitestBin(repoRoot, 'darwin')).toBe(resolve(repoRoot, 'node_modules/.bin/vitest'))
    expect(resolveVitestBin(repoRoot, 'win32')).toBe(resolve(repoRoot, 'node_modules/.bin/vitest.cmd'))
  })

  it('runs package Vitest wrappers through tsx instead of bun', () => {
    const packageJson = readFileSync(resolve(repoRoot, 'package.json'), 'utf8')
    const headlessPackageScripts = readPackageScripts(resolve(repoRoot, 'packages/workpaper/package.json'))
    const xlsxFormulaRecalcPackageScripts = readPackageScripts(resolve(repoRoot, 'packages/xlsx-formula-recalc/package.json'))
    const runVitestSource = readFileSync(resolve(repoRoot, 'scripts/run-vitest.ts'), 'utf8')

    expect(packageJson).toContain('"test": "tsx scripts/run-vitest.ts --run"')
    expect(packageJson).toContain('"coverage": "tsx scripts/run-vitest.ts --run --coverage')
    expect(packageJson).toContain('"test:watch": "tsx scripts/run-vitest.ts"')
    expect(packageJson).not.toContain('bun scripts/run-vitest.ts')
    expect(headlessPackageScripts['test:excel-oracle']).toContain('tsx ../../scripts/run-vitest.ts --root ../.. --run')
    expect(headlessPackageScripts['test:excel-oracle']).not.toContain('vitest run --root ../..')
    expect(xlsxFormulaRecalcPackageScripts['test:excel-oracle']).toContain('tsx ../../scripts/run-vitest.ts --root ../.. --run')
    expect(xlsxFormulaRecalcPackageScripts['test:excel-oracle']).not.toContain('vitest run --root ../..')
    expect(runVitestSource).toContain('process.stderr.write(`${error instanceof Error ? error.message : String(error)}\\n`)')
  })
})
