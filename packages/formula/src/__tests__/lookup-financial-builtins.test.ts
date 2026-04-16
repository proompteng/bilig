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
    expect(
      XNPV(
        num(0.09),
        cellRange([num(-1), num(2)], 2, 1),
        cellRange([num(39448), { tag: ValueTag.Number, value: Number.POSITIVE_INFINITY }], 2, 1),
      ),
    ).toEqual(err(ErrorCode.Value))
    expect(XIRR(xValues, xDates, cellRange([num(0.1)], 1, 1))).toEqual(err(ErrorCode.Value))

    expect(Reflect.apply(IRR, undefined, [])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(MIRR, undefined, [mirrValues, num(0.1)])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(XNPV, undefined, [num(0.09), xValues])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(XIRR, undefined, [xValues])).toEqual(err(ErrorCode.Value))
  })
})
