import { describe, expect, it } from 'vitest'
import { parseHeadlessPackageVersion } from '../check-docs-discovery-context.ts'

describe('docs discovery context loading', () => {
  it('parses the headless package version from package json', () => {
    expect(parseHeadlessPackageVersion(JSON.stringify({ name: '@bilig/workpaper', version: '1.2.3' }))).toBe('1.2.3')
  })

  it('rejects package json without a string version', () => {
    expect(() => parseHeadlessPackageVersion(JSON.stringify({ name: '@bilig/workpaper', version: 123 }))).toThrow(
      'packages/workpaper/package.json is missing a string version',
    )
  })
})
