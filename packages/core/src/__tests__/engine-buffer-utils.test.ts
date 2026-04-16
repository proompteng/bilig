import { describe, expect, it } from 'vitest'
import { appendPackedCellIndex, growUint32 } from '../engine-buffer-utils.js'

describe('engine buffer utils', () => {
  it('grows uint32 buffers while preserving existing values', () => {
    const source = new Uint32Array([3, 5, 8, 13])
    const grown = growUint32(source, 7)

    expect(grown).not.toBe(source)
    expect(Array.from(grown.slice(0, source.length))).toEqual([3, 5, 8, 13])
    expect(grown.length).toBe(8)
  })

  it('avoids duplicate packed indices and appends new ones', () => {
    const source = new Uint32Array([4, 9])

    expect(appendPackedCellIndex(source, 9)).toBe(source)
    expect(Array.from(appendPackedCellIndex(source, 12))).toEqual([4, 9, 12])
  })
})
