import { readFileSync } from 'node:fs'

import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { readWorkPaperPackageVersion } from '../work-paper-version.js'

function readHeadlessPackageVersion(): string {
  const packageJson = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8'))
  return readWorkPaperPackageVersion(packageJson)
}

describe('WorkPaper.version', () => {
  it('reports the headless package manifest version', () => {
    expect(WorkPaper.version).toBe(readHeadlessPackageVersion())
  })

  it('derives the public version from the runtime package manifest', () => {
    expect(readWorkPaperPackageVersion({ version: '9.9.9' })).toBe('9.9.9')
  })

  it('rejects missing or empty runtime package versions', () => {
    expect(() => readWorkPaperPackageVersion({})).toThrow('Expected @bilig/headless package.json')
    expect(() => readWorkPaperPackageVersion({ version: '' })).toThrow('Expected @bilig/headless package.json')
  })
})
