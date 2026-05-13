import { gzipSync } from 'node:zlib'
import { describe, expect, it, vi } from 'vitest'

import { measureGzipBytes, releaseCheckGzipOptions } from '../release-check-helpers.ts'

describe('release check gzip measurement', () => {
  it('uses explicit deterministic gzip compression options', () => {
    const bytes = new TextEncoder().encode('release-check bundle payload')
    const gzip = vi.fn((input: Uint8Array, options: typeof releaseCheckGzipOptions) => gzipSync(input, options))

    const measuredBytes = measureGzipBytes(bytes, gzip)

    expect(gzip).toHaveBeenCalledExactlyOnceWith(bytes, releaseCheckGzipOptions)
    expect(releaseCheckGzipOptions).toEqual({ level: 9 })
    expect(measuredBytes).toBe(gzipSync(bytes, releaseCheckGzipOptions).byteLength)
  })
})
