import { describe, expect, it } from 'vitest'
import { createKernel } from '../index.js'

describe('wasm kernel direct aggregate batch', () => {
  it('evaluates dense numeric row aggregate batches', async () => {
    const kernel = await createKernel()
    const outNumbers = new Float64Array(3)

    kernel.evalDenseNumericRowAggregateBatch(1, Float64Array.from([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]), 3, 4, 1, 2, 10, outNumbers)

    expect([...outNumbers]).toEqual([15, 23, 31])
  })
})
