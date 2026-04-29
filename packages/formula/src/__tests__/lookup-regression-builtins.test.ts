import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getLookupBuiltin, type RangeBuiltinArgument } from '../builtins/lookup.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

describe('lookup regression builtins', () => {
  it('supports paired regression helpers and exact linear forecasts', () => {
    const CORREL = getLookupBuiltin('CORREL')!
    const PEARSON = getLookupBuiltin('PEARSON')!
    const INTERCEPT = getLookupBuiltin('INTERCEPT')!
    const SLOPE = getLookupBuiltin('SLOPE')!
    const RSQ = getLookupBuiltin('RSQ')!
    const STEYX = getLookupBuiltin('STEYX')!
    const FORECAST = getLookupBuiltin('FORECAST')!
    const FORECAST_LINEAR = getLookupBuiltin('FORECAST.LINEAR')!
    const TREND = getLookupBuiltin('TREND')!
    const GROWTH = getLookupBuiltin('GROWTH')!
    const LINEST = getLookupBuiltin('LINEST')!
    const LOGEST = getLookupBuiltin('LOGEST')!

    const knownY = cellRange([num(5), num(8), num(11)], 3, 1)
    const knownX = cellRange([num(1), num(2), num(3)], 3, 1)
    const newX = cellRange([num(4), num(5)], 2, 1)
    const simpleY = cellRange([num(2), num(4), num(6)], 3, 1)
    const simpleX = cellRange([num(1), num(2), num(3)], 3, 1)

    expect(CORREL(knownY, knownX)).toEqual(num(1))
    expect(PEARSON(knownY, knownX)).toEqual(num(1))
    expect(INTERCEPT(knownY, knownX)).toEqual(num(2))
    expect(SLOPE(knownY, knownX)).toEqual(num(3))
    expect(RSQ(knownY, knownX)).toEqual(num(1))
    expect(STEYX(knownY, knownX)).toEqual(num(0))
    expect(FORECAST(num(4), knownY, knownX)).toEqual(num(14))
    expect(FORECAST_LINEAR(num(4), knownY, knownX)).toEqual(num(14))
    expect(TREND(knownY, knownX, newX)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(14), num(17)],
    })
    expect(TREND(knownY, knownX, num(4))).toEqual(num(14))
    expect(LINEST(knownY)).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [num(3), num(2)],
    })
    expect(LINEST(knownY, knownX)).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [num(3), num(2)],
    })
    expect(LINEST(simpleY, simpleX, bool(false))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [num(2), num(0)],
    })

    const growthY = cellRange([num(2), num(4), num(8)], 3, 1)
    const growth = GROWTH(growthY, knownX, newX)
    expect(growth).toMatchObject({ kind: 'array', rows: 2, cols: 1 })
    if (!('values' in growth)) {
      throw new Error('GROWTH should spill an array')
    }
    expect(growth.values[0]).toMatchObject({ tag: ValueTag.Number, value: expect.closeTo(16, 12) })
    expect(growth.values[1]).toMatchObject({ tag: ValueTag.Number, value: expect.closeTo(32, 12) })
    expect(GROWTH(growthY, knownX, num(4))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(16, 12),
    })
    const growthNoIntercept = GROWTH(growthY, knownX, newX, bool(false))
    expect(growthNoIntercept).toMatchObject({ kind: 'array', rows: 2, cols: 1 })
    if (!('values' in growthNoIntercept)) {
      throw new Error('GROWTH without intercept should spill an array')
    }
    expect(growthNoIntercept.values[0]).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(16, 12),
    })
    expect(growthNoIntercept.values[1]).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(32, 12),
    })
    expect(LOGEST(growthY, knownX)).toMatchObject({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: expect.closeTo(2, 12) },
        { tag: ValueTag.Number, value: expect.closeTo(1, 12) },
      ],
    })
    expect(LOGEST(growthY)).toMatchObject({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: expect.closeTo(2, 12) },
        { tag: ValueTag.Number, value: expect.closeTo(1, 12) },
      ],
    })
    expect(LOGEST(growthY, knownX, bool(false))).toMatchObject({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: expect.closeTo(2, 12) },
        { tag: ValueTag.Number, value: expect.closeTo(1, 12) },
      ],
    })

    expect(SLOPE(knownY, cellRange([num(2), num(2), num(2)], 3, 1))).toEqual(err(ErrorCode.Div0))
    expect(STEYX(cellRange([num(5), num(8)], 2, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Div0))
    expect(FORECAST(text('oops'), knownY, knownX)).toEqual(err(ErrorCode.Value))
    expect(TREND(knownY, cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(TREND(knownY, knownX, newX, text('bad'))).toEqual(err(ErrorCode.Value))
    expect(GROWTH(cellRange([num(2), num(0), num(8)], 3, 1), knownX, newX)).toEqual(err(ErrorCode.Value))
    expect(LINEST(knownY, knownX, bool(true), bool(true))).toMatchObject({
      kind: 'array',
      rows: 5,
      cols: 2,
    })
    expect(LOGEST(growthY, knownX, bool(true), bool(true))).toMatchObject({
      kind: 'array',
      rows: 5,
      cols: 2,
    })
    const minimalLineStats = LINEST(cellRange([num(1), num(2)], 2, 1), cellRange([num(1), num(2)], 2, 1), bool(true), bool(true))
    expect(minimalLineStats).toMatchObject({ kind: 'array', rows: 5, cols: 2 })
    if (!('values' in minimalLineStats)) {
      throw new Error('LINEST stats should spill a matrix')
    }
    expect(minimalLineStats.values[2]).toEqual(err(ErrorCode.Div0))
    expect(minimalLineStats.values[3]).toEqual(err(ErrorCode.Div0))
    expect(minimalLineStats.values[5]).toEqual(err(ErrorCode.Div0))
    expect(minimalLineStats.values[6]).toEqual(err(ErrorCode.Div0))
    expect(LINEST(knownY, knownX, text('bad'))).toEqual(err(ErrorCode.Value))
    expect(LOGEST(growthY, knownX, text('bad'))).toEqual(err(ErrorCode.Value))
    expect(LINEST(knownY, knownX, bool(true), text('bad'))).toEqual(err(ErrorCode.Value))
    expect(LOGEST(growthY, knownX, bool(true), text('bad'))).toEqual(err(ErrorCode.Value))
    expect(LINEST(knownY, cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(LOGEST(cellRange([num(2), num(0), num(8)], 3, 1), knownX)).toEqual(err(ErrorCode.Value))
  })

  it('returns value errors for missing required regression arguments instead of throwing', () => {
    const CORREL = getLookupBuiltin('CORREL')!
    const FORECAST = getLookupBuiltin('FORECAST')!
    const LINEST = getLookupBuiltin('LINEST')!

    const knownY = cellRange([num(5), num(8), num(11)], 3, 1)

    expect(Reflect.apply(CORREL, undefined, [knownY])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(FORECAST, undefined, [num(4), knownY])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(LINEST, undefined, [])).toEqual(err(ErrorCode.Value))
  })

  it('covers covariance, scalar matrix coercion, and regression validation branches', () => {
    const COVAR = getLookupBuiltin('COVAR')!
    const COVARIANCE_P = getLookupBuiltin('COVARIANCE.P')!
    const COVARIANCE_S = getLookupBuiltin('COVARIANCE.S')!
    const CORREL = getLookupBuiltin('CORREL')!
    const FORECAST = getLookupBuiltin('FORECAST')!
    const TREND = getLookupBuiltin('TREND')!
    const GROWTH = getLookupBuiltin('GROWTH')!
    const LINEST = getLookupBuiltin('LINEST')!
    const LOGEST = getLookupBuiltin('LOGEST')!

    const y = cellRange([num(2), num(4), num(6), num(8)], 4, 1)
    const x = cellRange([num(1), num(2), num(3), num(4)], 4, 1)

    expect(COVAR(y, x)).toEqual(num(2.5))
    expect(COVARIANCE_P(y, x)).toEqual(num(2.5))
    expect(COVARIANCE_S(y, x)).toEqual(num(10 / 3))
    expect(COVARIANCE_S(cellRange([num(1)], 1, 1), cellRange([num(2)], 1, 1))).toEqual(err(ErrorCode.Div0))
    expect(CORREL(cellRange([num(1)], 1, 1), cellRange([num(2)], 1, 1))).toEqual(err(ErrorCode.Div0))
    expect(CORREL(cellRange([num(1), num(1)], 2, 1), cellRange([num(2), num(3)], 2, 1))).toEqual(err(ErrorCode.Div0))
    expect(CORREL(cellRange([err(ErrorCode.Ref), num(1)], 2, 1), cellRange([num(2), num(3)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(CORREL({ kind: 'range', refKind: 'rows', rows: 1, cols: 2, values: [num(1), num(2)] }, x)).toEqual(err(ErrorCode.Value))

    expect(FORECAST(err(ErrorCode.NA), y, x)).toEqual(err(ErrorCode.NA))
    expect(FORECAST(cellRange([num(4)], 1, 1), y, x)).toEqual(err(ErrorCode.Value))
    expect(FORECAST(num(4), y, cellRange([num(1), num(1), num(1), num(1)], 4, 1))).toEqual(err(ErrorCode.Div0))

    expect(TREND(y)).toEqual({
      kind: 'array',
      rows: 4,
      cols: 1,
      values: [num(2), num(4), num(6), num(8)],
    })
    expect(TREND(y, x, cellRange([num(5), num(6)], 1, 2), bool(false))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [num(10), num(12)],
    })
    expect(TREND(y, x, { kind: 'range', refKind: 'columns', rows: 1, cols: 1, values: [num(5)] })).toEqual(err(ErrorCode.Value))
    expect(TREND(y, x, num(Number.POSITIVE_INFINITY))).toEqual(err(ErrorCode.Value))
    expect(TREND(y, x, num(5), cellRange([bool(true)], 1, 1))).toEqual(err(ErrorCode.Value))

    expect(GROWTH(cellRange([num(2), num(4), num(8), num(16)], 4, 1))).toMatchObject({
      kind: 'array',
      rows: 4,
      cols: 1,
    })
    expect(GROWTH(cellRange([num(2), num(4), num(8), num(16)], 4, 1), x, num(Number.POSITIVE_INFINITY))).toEqual(err(ErrorCode.Value))

    expect(LINEST(y, x, cellRange([bool(true)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(LINEST(cellRange([], 0, 0), x)).toEqual(err(ErrorCode.Value))
    expect(LOGEST(cellRange([err(ErrorCode.Ref)], 1, 1), x)).toEqual(err(ErrorCode.Ref))
    expect(LOGEST(y, cellRange([num(Number.NaN), num(2), num(3), num(4)], 4, 1))).toEqual(err(ErrorCode.Value))
  })
})
