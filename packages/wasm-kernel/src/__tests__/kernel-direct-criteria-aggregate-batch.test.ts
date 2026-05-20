import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

const DIRECT_AGGREGATE_OP_SUM = 1
const DIRECT_AGGREGATE_OP_AVERAGE = 2
const DIRECT_AGGREGATE_OP_COUNT = 3
const DIRECT_AGGREGATE_OP_MIN = 4
const DIRECT_AGGREGATE_OP_MAX = 5
const CRITERIA_KIND_NUMBER = 0
const CRITERIA_KIND_STRING_ID = 1

describe('wasm kernel direct criteria aggregate batch', () => {
  it('reduces matched row offsets with criteria aggregate semantics', async () => {
    const kernel = await createKernel()
    const outTags = new Uint8Array(7)
    const outNumbers = new Float64Array(7)
    const outErrors = new Uint16Array(7)

    kernel.evalDirectCriteriaMatchedAggregateBatch(
      Uint8Array.from([
        DIRECT_AGGREGATE_OP_SUM,
        DIRECT_AGGREGATE_OP_AVERAGE,
        DIRECT_AGGREGATE_OP_MIN,
        DIRECT_AGGREGATE_OP_MAX,
        DIRECT_AGGREGATE_OP_COUNT,
        DIRECT_AGGREGATE_OP_SUM,
        DIRECT_AGGREGATE_OP_AVERAGE,
      ]),
      Uint32Array.from([0, 0, 0, 0, 0, 4, 3]),
      Uint32Array.from([4, 4, 4, 4, 4, 3, 1]),
      Uint32Array.from([0, 1, 2, 3, 4, 5, 0, 4]),
      Uint8Array.from([ValueTag.Number, ValueTag.Boolean, ValueTag.Empty, ValueTag.String, ValueTag.Error, ValueTag.Number]),
      Float64Array.from([5, 1, 0, 0, 0, 9]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.Ref, ErrorCode.None]),
      outTags,
      outNumbers,
      outErrors,
    )

    expect([...outTags]).toEqual([
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Error,
      ValueTag.Error,
    ])
    expect(outNumbers[0]).toBe(6)
    expect(outNumbers[1]).toBe(2)
    expect(outNumbers[2]).toBe(5)
    expect(outNumbers[3]).toBe(5)
    expect(outNumbers[4]).toBe(4)
    expect(outErrors[5]).toBe(ErrorCode.Ref)
    expect(outErrors[6]).toBe(ErrorCode.Div0)
  })

  it('scans numeric criteria predicates and aggregates in one native pass', async () => {
    const kernel = await createKernel()
    const outTags = new Uint8Array(1)
    const outNumbers = new Float64Array(1)
    const outErrors = new Uint16Array(1)

    kernel.evalDirectCriteriaPredicateAggregateBatch(
      DIRECT_AGGREGATE_OP_SUM,
      6,
      Uint8Array.from([3, 0]),
      Uint8Array.from([CRITERIA_KIND_NUMBER, CRITERIA_KIND_NUMBER]),
      Float64Array.from([3, 1]),
      Uint32Array.from([0, 0]),
      Uint8Array.from([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Boolean,
        ValueTag.Boolean,
        ValueTag.Boolean,
        ValueTag.Boolean,
        ValueTag.Boolean,
        ValueTag.Boolean,
      ]),
      Float64Array.from([1, 2, 3, 4, 5, 6, 1, 0, 1, 1, 0, 1]),
      Uint32Array.from([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      Uint8Array.from([ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number]),
      Float64Array.from([10, 20, 30, 40, 50, 60]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      outTags,
      outNumbers,
      outErrors,
    )

    expect(outTags[0]).toBe(ValueTag.Number)
    expect(outNumbers[0]).toBe(130)
    expect(outErrors[0]).toBe(ErrorCode.None)
  })

  it('matches string id criteria with numeric predicates in one native pass', async () => {
    const kernel = await createKernel()
    const outTags = new Uint8Array(1)
    const outNumbers = new Float64Array(1)
    const outErrors = new Uint16Array(1)
    const targetStringId = 11

    kernel.evalDirectCriteriaPredicateAggregateBatch(
      DIRECT_AGGREGATE_OP_SUM,
      6,
      Uint8Array.from([0, 3, 0]),
      Uint8Array.from([CRITERIA_KIND_STRING_ID, CRITERIA_KIND_NUMBER, CRITERIA_KIND_STRING_ID]),
      Float64Array.from([0, 3, 0]),
      Uint32Array.from([targetStringId, 0, 21]),
      Uint8Array.from([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
      ]),
      Float64Array.from([0, 0, 0, 0, 0, 0, 1, 2, 3, 4, 5, 6, 0, 0, 0, 0, 0, 0]),
      Uint32Array.from([targetStringId, 12, targetStringId, targetStringId, 12, targetStringId, 0, 0, 0, 0, 0, 0, 21, 21, 22, 21, 21, 22]),
      Uint8Array.from([ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number]),
      Float64Array.from([10, 20, 30, 40, 50, 60]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      outTags,
      outNumbers,
      outErrors,
    )

    expect(outTags[0]).toBe(ValueTag.Number)
    expect(outNumbers[0]).toBe(40)
    expect(outErrors[0]).toBe(ErrorCode.None)
  })
})
