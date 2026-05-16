import { describe, expect, it } from 'vitest'

import { parseAfterRevisionQuery } from './sync-server-document-routes.js'

describe('sync server document route inputs', () => {
  it('defaults missing afterRevision values to zero', () => {
    expect(parseAfterRevisionQuery(undefined)).toBe(0)
    expect(parseAfterRevisionQuery(' 0 ')).toBe(0)
  })

  it('accepts safe non-negative integer revisions', () => {
    expect(parseAfterRevisionQuery('7')).toBe(7)
    expect(parseAfterRevisionQuery(String(Number.MAX_SAFE_INTEGER))).toBe(Number.MAX_SAFE_INTEGER)
  })

  it.each(['', ' ', '-1', '+1', '1.5', '4abc', '01', String(Number.MAX_SAFE_INTEGER + 1)])(
    'rejects malformed afterRevision value %s',
    (value) => {
      expect(parseAfterRevisionQuery(value)).toBeNull()
    },
  )
})
