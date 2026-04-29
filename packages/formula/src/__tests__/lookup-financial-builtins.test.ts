import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getLookupBuiltin, type RangeBuiltinArgument } from '../builtins/lookup.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

describe('lookup financial builtins', () => {
  it('supports native cash-flow rate helpers on numeric ranges', () => {
    const IRR = getLookupBuiltin('IRR')!
    const MIRR = getLookupBuiltin('MIRR')!
    const XNPV = getLookupBuiltin('XNPV')!
    const XIRR = getLookupBuiltin('XIRR')!

    const irrValues = cellRange([num(-70000), num(12000), num(15000), num(18000), num(21000), num(26000)], 6, 1)
    const mirrValues = cellRange([num(-120000), num(39000), num(30000), num(21000), num(37000), num(46000)], 6, 1)
    const xValues = cellRange([num(-10000), num(2750), num(4250), num(3250), num(2750)], 5, 1)
    const xDates = cellRange([num(39448), num(39508), num(39751), num(39859), num(39904)], 5, 1)

    const irr = IRR(irrValues)
    if (irr.tag !== ValueTag.Number) throw new Error(`Expected number result, received ${irr.tag}`)
    expect(irr.value).toBeCloseTo(0.08663094803653162, 12)

    const mirr = MIRR(mirrValues, num(0.1), num(0.12))
    if (mirr.tag !== ValueTag.Number) throw new Error(`Expected number result, received ${mirr.tag}`)
    expect(mirr.value).toBeCloseTo(0.1260941303659051, 12)

    const xnpv = XNPV(num(0.09), xValues, xDates)
    if (xnpv.tag !== ValueTag.Number) throw new Error(`Expected number result, received ${xnpv.tag}`)
    expect(xnpv.value).toBeCloseTo(2086.647602031535, 9)

    const xirr = XIRR(xValues, xDates)
    if (xirr.tag !== ValueTag.Number) throw new Error(`Expected number result, received ${xirr.tag}`)
    expect(xirr.value).toBeCloseTo(0.37336253351883136, 12)

    expect(IRR(cellRange([num(5), num(7)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(MIRR(cellRange([num(5), num(7)], 2, 1), num(0.1), num(0.12))).toEqual(err(ErrorCode.Div0))
    expect(XNPV(num(0.09), xValues, cellRange([num(39448), num(39508)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(XNPV(num(0.09), xValues, cellRange([num(39448), num(39508), num(39400), num(39859), num(39904)], 5, 1))).toEqual(
      err(ErrorCode.Value),
    )
    expect(XIRR(xValues, cellRange([num(39448), num(39508), num(39400), num(39859), num(39904)], 5, 1))).toEqual(err(ErrorCode.Value))
    expect(XIRR(xValues, xDates, text('bad'))).toEqual(err(ErrorCode.Value))
  })

  it('covers remaining cash-flow validation branches and missing required args', () => {
    const IRR = getLookupBuiltin('IRR')!
    const MIRR = getLookupBuiltin('MIRR')!
    const XNPV = getLookupBuiltin('XNPV')!
    const XIRR = getLookupBuiltin('XIRR')!

    const xValues = cellRange([num(-10000), num(2750), num(4250), num(3250), num(2750)], 5, 1)
    const xDates = cellRange([num(39448), num(39508), num(39751), num(39859), num(39904)], 5, 1)
    const mirrValues = cellRange([num(-120000), num(39000), num(30000), num(21000), num(37000), num(46000)], 6, 1)

    expect(IRR(cellRange([err(ErrorCode.Name), num(1)], 2, 1))).toEqual(err(ErrorCode.Name))
    expect(MIRR({ kind: 'range', refKind: 'rows', rows: 1, cols: 2, values: [num(-1), num(2)] }, num(0.1), num(0.12))).toEqual(
      err(ErrorCode.Value),
    )
    expect(MIRR(mirrValues, num(-1), num(0.12))).toEqual(err(ErrorCode.Div0))
    expect(XNPV(err(ErrorCode.Name), xValues, xDates)).toEqual(err(ErrorCode.Name))
    expect(
      XNPV(
        num(0.09),
        { kind: 'range', refKind: 'rows', rows: 1, cols: 2, values: [num(-1), num(2)] },
        cellRange([num(39448), num(39508)], 2, 1),
      ),
    ).toEqual(err(ErrorCode.Value))
    expect(XNPV(num(0.09), cellRange([err(ErrorCode.Ref), num(2)], 2, 1), cellRange([num(39448), num(39508)], 2, 1))).toEqual(
      err(ErrorCode.Ref),
    )
    expect(XNPV(text('bad'), xValues, xDates)).toEqual(err(ErrorCode.Value))
    expect(XNPV(num(-1), xValues, xDates)).toEqual(err(ErrorCode.Value))
    expect(XNPV(num(0.09), cellRange([num(1), num(2)], 2, 1), cellRange([num(39448), num(39508)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(
      XNPV(
        num(0.09),
        cellRange([num(-1), num(2)], 2, 1),
        cellRange([num(39448), { tag: ValueTag.Number, value: Number.POSITIVE_INFINITY }], 2, 1),
      ),
    ).toEqual(err(ErrorCode.Value))
    expect(XIRR(cellRange([num(1), num(2)], 2, 1), cellRange([num(39448), num(39508)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(XIRR(xValues, cellRange([num(39448), num(39400), num(39508), num(39859), num(39904)], 5, 1))).toEqual(err(ErrorCode.Value))
    expect(XIRR(xValues, xDates, cellRange([num(0.1)], 1, 1))).toEqual(err(ErrorCode.Value))

    expect(IRR(xValues, err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(IRR(xValues, cellRange([num(0.1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(MIRR(mirrValues, err(ErrorCode.Ref), num(0.12))).toEqual(err(ErrorCode.Ref))
    expect(MIRR(mirrValues, num(0.1), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(MIRR(mirrValues, text('bad'), num(0.12))).toEqual(err(ErrorCode.Value))
    expect(MIRR(mirrValues, num(0.1), text('bad'))).toEqual(err(ErrorCode.Value))
    expect(MIRR(mirrValues, num(0.1), num(-1))).toEqual(err(ErrorCode.Div0))

    expect(Reflect.apply(IRR, undefined, [])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(MIRR, undefined, [mirrValues, num(0.1)])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(XNPV, undefined, [num(0.09), xValues])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(XIRR, undefined, [xValues])).toEqual(err(ErrorCode.Value))
  })

  it('covers zero-rate roots, solver fallback guesses, and date serial validation', () => {
    const IRR = getLookupBuiltin('IRR')!
    const MIRR = getLookupBuiltin('MIRR')!
    const XNPV = getLookupBuiltin('XNPV')!
    const XIRR = getLookupBuiltin('XIRR')!

    const zeroRateValues = cellRange([num(-100), num(100)], 2, 1)
    const zeroRateDates = cellRange([num(45_000), num(45_365)], 2, 1)

    expect(IRR(zeroRateValues)).toEqual(num(0))
    expect(IRR(zeroRateValues, num(Number.NaN))).toEqual(num(0))
    expect(IRR(cellRange([num(-100), num(20), num(30)], 3, 1), num(-5))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.any(Number),
    })

    expect(MIRR(cellRange([num(-100)], 1, 1), num(0.1), num(0.1))).toEqual(err(ErrorCode.Div0))
    expect(MIRR(cellRange([num(-100), num(0), num(100)], 3, 1), num(0.1), num(0.1))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.any(Number),
    })
    expect(MIRR(cellRange([num(-100), num(100)], 2, 1), num(Number.POSITIVE_INFINITY), num(0.1))).toEqual(err(ErrorCode.Div0))
    expect(MIRR(cellRange([num(-100), num(100)], 2, 1), num(0.1), num(Number.NaN))).toEqual(err(ErrorCode.Div0))

    expect(XNPV(num(0), zeroRateValues, zeroRateDates)).toEqual(num(0))
    expect(XNPV(num(Number.NaN), zeroRateValues, zeroRateDates)).toEqual(err(ErrorCode.Value))
    expect(XNPV(num(0.1), zeroRateValues, cellRange([num(45_000), num(Number.NaN)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(XNPV(num(0.1), zeroRateValues, cellRange([num(45_000), num(0)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(XNPV(num(0.1), cellRange([], 0, 0), cellRange([], 0, 0))).toEqual(err(ErrorCode.Value))

    expect(XIRR(zeroRateValues, zeroRateDates)).toEqual(num(0))
    expect(XIRR(zeroRateValues, zeroRateDates, num(-5))).toEqual(num(0))
    expect(XIRR(zeroRateValues, zeroRateDates, num(Number.POSITIVE_INFINITY))).toEqual(num(0))
    expect(XIRR(cellRange([num(-100), num(20), num(30)], 3, 1), cellRange([num(45_000), num(45_365), num(45_730)], 3, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.any(Number),
    })
  })
})
