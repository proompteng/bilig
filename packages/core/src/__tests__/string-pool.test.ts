import { describe, expect, it } from 'vitest'
import { StringPool } from '../string-pool.js'

describe('StringPool', () => {
  it('exports stable lengths and UTF-16 layout for interned strings', () => {
    const pool = new StringPool()

    expect(pool.intern('alpha')).toBe(1)
    expect(pool.intern('beta')).toBe(2)
    expect(pool.intern('alpha')).toBe(1)
    expect(pool.size).toBe(3)
    expect(pool.get(2)).toBe('beta')
    expect(pool.get(99)).toBe('')
    expect([...pool.exportLengths()]).toEqual([0, 5, 4])

    const layout = pool.exportLayout()

    expect([...layout.offsets]).toEqual([0, 0, 5])
    expect([...layout.lengths]).toEqual([0, 5, 4])
    expect(String.fromCharCode(...layout.data)).toBe('alphabeta')
  })
})
