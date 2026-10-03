import { afterEach, describe, expect, it, vi } from 'vitest'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { resolveCoverageFilePath, runCoverageContracts } from '../coverage-contracts.ts'

const originalCoverageDir = process.env['BILIG_COVERAGE_DIR']
const originalCoverageFile = process.env['BILIG_COVERAGE_FILE']
const temporaryDirectories: string[] = []

afterEach(async () => {
  if (originalCoverageDir === undefined) {
    delete process.env['BILIG_COVERAGE_DIR']
  } else {
    process.env['BILIG_COVERAGE_DIR'] = originalCoverageDir
  }

  if (originalCoverageFile === undefined) {
    delete process.env['BILIG_COVERAGE_FILE']
  } else {
    process.env['BILIG_COVERAGE_FILE'] = originalCoverageFile
  }

  await Promise.all(temporaryDirectories.splice(0).map((directory) => rm(directory, { force: true, recursive: true })))

  vi.restoreAllMocks()
})

function fullyCoveredFile(path: string) {
  return {
    [path]: {
      s: { '0': 1 },
      statementMap: {
        '0': {
          start: { line: 1 },
          end: { line: 100 },
        },
      },
    },
  }
}

describe('coverage contracts path resolution', () => {
  it('reads coverage-final from the configured coverage reports directory', () => {
    delete process.env['BILIG_COVERAGE_FILE']
    process.env['BILIG_COVERAGE_DIR'] = 'coverage/ci-123'

    expect(resolveCoverageFilePath()).toBe(resolve('coverage/ci-123/coverage-final.json'))
  })

  it('allows an explicit coverage file path to override the reports directory', () => {
    process.env['BILIG_COVERAGE_DIR'] = 'coverage/ci-123'
    process.env['BILIG_COVERAGE_FILE'] = 'tmp/custom-coverage.json'

    expect(resolveCoverageFilePath()).toBe(resolve('tmp/custom-coverage.json'))
  })

  it('validates a coverage-final file without requiring Bun globals', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'bilig-coverage-contracts-'))
    temporaryDirectories.push(directory)
    const coverageFile = join(directory, 'coverage-final.json')
    await writeFile(
      coverageFile,
      JSON.stringify({
        ...fullyCoveredFile('/repo/packages/core/src/engine.ts'),
        ...fullyCoveredFile('/repo/packages/formula/src/builtins.ts'),
      }),
    )

    const log = vi.spyOn(console, 'log').mockImplementation(() => {})

    await expect(runCoverageContracts(coverageFile)).resolves.toBeUndefined()
    expect(log).toHaveBeenCalledWith(expect.stringContaining('"label": "packages/core/src"'))
  })
})
