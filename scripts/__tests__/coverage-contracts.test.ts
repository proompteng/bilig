import { afterEach, describe, expect, it } from 'vitest'
import { resolve } from 'node:path'
import { resolveCoverageFilePath } from '../coverage-contracts.ts'

const originalCoverageDir = process.env['BILIG_COVERAGE_DIR']
const originalCoverageFile = process.env['BILIG_COVERAGE_FILE']

afterEach(() => {
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
})

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
})
