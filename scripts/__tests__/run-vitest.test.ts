import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildVitestArgBatches, buildVitestArgs, isBroadCorpusVitestRun, readVitestBatchCooldownMs } from '../run-vitest.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('run-vitest wrapper arguments', () => {
  it('bounds Vitest workers in CI by default', () => {
    expect(buildVitestArgs(['--run', 'sample.test.ts'], { BILIG_CI_PROFILE: 'fast' })).toEqual([
      '--run',
      'sample.test.ts',
      '--maxWorkers',
      '1',
    ])
  })

  it('preserves an explicit maxWorkers flag', () => {
    expect(buildVitestArgs(['--run', '--maxWorkers=1'], { BILIG_CI_PROFILE: 'fast' })).toEqual(['--run', '--maxWorkers=1'])
  })

  it('allows CI worker limit overrides', () => {
    expect(
      buildVitestArgs(['--run'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_MAX_WORKERS: '3',
      }),
    ).toEqual(['--run', '--maxWorkers', '3'])
  })

  it('splits large CI run file lists into serial batches', () => {
    const files = Array.from({ length: 7 }, (_, index) => `test-${index + 1}.test.ts`)

    expect(
      buildVitestArgBatches(['--run', ...files], {
        BILIG_CI_PROFILE: 'fast',
      }),
    ).toEqual([
      ['--run', ...files.slice(0, 3), '--maxWorkers', '1'],
      ['--run', ...files.slice(3, 6), '--maxWorkers', '1'],
      ['--run', files[6], '--maxWorkers', '1'],
    ])
  })

  it('keeps broad corpus CI runs in one worker-bounded batch', () => {
    const files = [
      'scripts/__tests__/public-workbook-corpus.test.ts',
      'scripts/__tests__/public-workbook-corpus-cli.test.ts',
      'scripts/__tests__/public-workbook-corpus-evidence-refresh.test.ts',
      'scripts/__tests__/public-workbook-corpus-completion-audit.test.ts',
      'scripts/__tests__/public-workbook-corpus-completion-audit-roundtrip.test.ts',
      'scripts/__tests__/public-workbook-corpus-feature-witness-plan.test.ts',
      'scripts/__tests__/public-workbook-corpus-financial-plan.test.ts',
      'scripts/__tests__/public-workbook-corpus-links.test.ts',
      'scripts/__tests__/public-workbook-corpus-resource-limit-plan.test.ts',
      'scripts/__tests__/public-workbook-corpus-missing.test.ts',
      'scripts/__tests__/public-workbook-corpus-verify-checkpoint.test.ts',
      'scripts/__tests__/public-workbook-corpus-workbook.test.ts',
      'packages/excel-import/src/__tests__/xlsx-formula-cache-roundtrip.test.ts',
      'packages/excel-import/src/__tests__/xlsx-table-sort-state-roundtrip.test.ts',
    ]

    expect(
      buildVitestArgBatches(['--run', ...files], {
        BILIG_CI_PROFILE: 'fast',
      }),
    ).toEqual([['--run', ...files, '--maxWorkers', '1']])
  })

  it('allows CI file chunk size overrides', () => {
    expect(
      buildVitestArgBatches(['--run', 'a.test.ts', 'b.test.ts', 'c.test.ts'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '2',
      }),
    ).toEqual([
      ['--run', 'a.test.ts', 'b.test.ts', '--maxWorkers', '1'],
      ['--run', 'c.test.ts', '--maxWorkers', '1'],
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
      ['--run', 'a.test.ts', 'b.test.ts', 'c.test.ts', '--maxWorkers', '1'],
      ['--run', 'd.test.ts', '--maxWorkers', '1'],
    ])
  })

  it('does not split run arguments that include flags', () => {
    expect(
      buildVitestArgBatches(['--run', 'sample.test.ts', '--reporter=dot'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '1',
      }),
    ).toEqual([['--run', 'sample.test.ts', '--reporter=dot', '--maxWorkers', '1']])
  })

  it('adds a short CI-only cooldown between split batches', () => {
    expect(readVitestBatchCooldownMs({})).toBe(0)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast' })).toBe(1000)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '0' })).toBe(0)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '2500' })).toBe(2500)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '2500ms' })).toBe(1000)
  })

  it('classifies the public workbook corpus correctness lane as broad', () => {
    expect(
      isBroadCorpusVitestRun([
        '--run',
        'scripts/__tests__/public-workbook-corpus.test.ts',
        'scripts/__tests__/public-workbook-corpus-cli.test.ts',
        'scripts/__tests__/public-workbook-corpus-completion-audit.test.ts',
        'scripts/__tests__/public-workbook-corpus-links.test.ts',
      ]),
    ).toBe(true)
  })

  it('allows focused public workbook corpus Vitest checks', () => {
    expect(isBroadCorpusVitestRun(['--run', 'scripts/__tests__/public-workbook-corpus-links.test.ts'])).toBe(false)
  })

  it('classifies mixed public-corpus and xlsx import correctness as broad', () => {
    expect(
      isBroadCorpusVitestRun([
        '--run',
        'scripts/__tests__/public-workbook-corpus.test.ts',
        'packages/excel-import/src/__tests__/xlsx-formula-cache-roundtrip.test.ts',
      ]),
    ).toBe(true)
  })

  it('runs package Vitest wrappers through tsx instead of bun', () => {
    const packageJson = readFileSync(resolve(repoRoot, 'package.json'), 'utf8')
    const runVitestSource = readFileSync(resolve(repoRoot, 'scripts/run-vitest.ts'), 'utf8')

    expect(packageJson).toContain('"test": "tsx scripts/run-vitest.ts --run"')
    expect(packageJson).toContain('"coverage": "tsx scripts/run-vitest.ts --run --coverage')
    expect(packageJson).toContain('"test:watch": "tsx scripts/run-vitest.ts"')
    expect(packageJson).not.toContain('bun scripts/run-vitest.ts')
    expect(runVitestSource).toContain('process.stderr.write(`${error instanceof Error ? error.message : String(error)}\\n`)')
  })
})
