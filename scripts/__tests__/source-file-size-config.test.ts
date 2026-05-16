import { describe, expect, it } from 'vitest'

import { parseSourceMaxLines } from '../source-file-size-config.js'

describe('source file size config', () => {
  it('defaults to the repository source size limit', () => {
    expect(parseSourceMaxLines(undefined)).toBe(1000)
  })

  it('accepts explicit positive safe integer limits', () => {
    expect(parseSourceMaxLines('1')).toBe(1)
    expect(parseSourceMaxLines('1500')).toBe(1500)
  })

  it.each(['', '0', '-1', '1000.5', '1000abc'])('rejects malformed BILIG_SOURCE_MAX_LINES=%s', (value) => {
    expect(() => parseSourceMaxLines(value)).toThrow(`BILIG_SOURCE_MAX_LINES must be a positive integer, got ${value}`)
  })

  it('rejects unsafe integer limits', () => {
    expect(() => parseSourceMaxLines('9007199254740992')).toThrow('BILIG_SOURCE_MAX_LINES must be a safe integer, got 9007199254740992')
  })
})
