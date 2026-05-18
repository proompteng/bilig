import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

describe('wasm kernel direct scalar batch', () => {
  it('evaluates value batches with dependencies on earlier batch outputs', async () => {
    const kernel = await createKernel()
    const outTags = new Uint8Array(3)
    const outNumbers = new Float64Array(3)
    const outErrors = new Uint16Array(3)

    kernel.evalDirectScalarValueBatch(
      Uint8Array.from([1, 3, 1]),
      Uint32Array.from([0xffffffff, 0, 0xffffffff]),
      Uint8Array.from([ValueTag.Number, ValueTag.Empty, ValueTag.String]),
      Float64Array.from([2, 0, 0]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      Uint32Array.from([0xffffffff, 0xffffffff, 0xffffffff]),
      Uint8Array.from([ValueTag.Number, ValueTag.Number, ValueTag.Number]),
      Float64Array.from([3, 2, 1]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      Float64Array.from([0, 5, 0]),
      outTags,
      outNumbers,
      outErrors,
    )

    expect(outTags[0]).toBe(ValueTag.Number)
    expect(outNumbers[0]).toBe(5)
    expect(outTags[1]).toBe(ValueTag.Number)
    expect(outNumbers[1]).toBe(15)
    expect(outTags[2]).toBe(ValueTag.Error)
    expect(outErrors[2]).toBe(ErrorCode.Value)
  })
})
