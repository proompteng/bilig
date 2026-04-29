import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getLookupBuiltin, type RangeBuiltinArgument } from '../builtins/lookup.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

describe('lookup order-statistics builtins', () => {
  it('supports PROB and TRIMMEAN with validation-heavy edge cases', () => {
    const PROB = getLookupBuiltin('PROB')!
    const TRIMMEAN = getLookupBuiltin('TRIMMEAN')!

    const xValues = cellRange([num(1), num(2), num(3), num(4)], 4, 1)
    const probabilities = cellRange([num(0.1), num(0.2), num(0.3), num(0.4)], 4, 1)
    const mixedTrimRange = cellRange(
      [num(1), text('ignored'), num(4), num(7), { tag: ValueTag.Boolean, value: true }, num(9), num(10), num(12)],
      8,
      1,
    )

    expect(PROB(xValues, probabilities, num(2))).toEqual(num(0.2))
    expect(PROB(xValues, probabilities, num(2), num(3))).toEqual(num(0.5))
    expect(PROB(xValues, probabilities, num(5), num(6))).toEqual(num(0))
    expect(TRIMMEAN(cellRange([num(1), num(2), num(4), num(7), num(8), num(9), num(10), num(12)], 8, 1), num(0.25))).toEqual(num(40 / 6))
    expect(TRIMMEAN(mixedTrimRange, num(0.4))).toEqual(num(7.5))

    expect(PROB(cellRange([num(1), num(2)], 2, 1), cellRange([num(0.4)], 1, 1), num(1))).toEqual(err(ErrorCode.Value))
    expect(PROB(xValues, cellRange([num(0.1), num(0.2), num(0.3), num(0.5)], 4, 1), num(1))).toEqual(err(ErrorCode.Value))
    expect(PROB(cellRange([err(ErrorCode.Ref), num(2), num(3), num(4)], 4, 1), probabilities, num(1))).toEqual(err(ErrorCode.Ref))
    expect(PROB(xValues, cellRange([num(0.1), err(ErrorCode.NA), num(0.3), num(0.6)], 4, 1), num(1))).toEqual(err(ErrorCode.NA))
    expect(PROB(xValues, cellRange([num(0.1), text('bad'), num(0.3), num(0.6)], 4, 1), num(1))).toEqual(err(ErrorCode.Value))
    expect(PROB(xValues, probabilities, text('bad'))).toEqual(err(ErrorCode.Value))
    expect(PROB(xValues, probabilities, num(3), num(2))).toEqual(err(ErrorCode.Value))
    expect(PROB(xValues, probabilities, num(3), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(PROB(xValues, probabilities, cellRange([num(2)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(TRIMMEAN(cellRange([text('bad')], 1, 1), num(0.25))).toEqual(err(ErrorCode.Value))
    expect(TRIMMEAN(cellRange([num(1), num(2)], 2, 1), num(1))).toEqual(err(ErrorCode.Value))
    expect(TRIMMEAN(cellRange([num(1), num(2)], 2, 1), text('bad'))).toEqual(err(ErrorCode.Value))
    expect(TRIMMEAN(err(ErrorCode.Ref), num(0.25))).toEqual(err(ErrorCode.Ref))
    expect(TRIMMEAN(cellRange([num(1), num(2)], 2, 1), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(TRIMMEAN({ kind: 'range', refKind: 'rows', values: [num(1), num(2)], rows: 2, cols: 1 }, num(0.25))).toEqual(
      err(ErrorCode.Value),
    )
    expect(TRIMMEAN(num(8), num(0))).toEqual(num(8))
  })

  it('covers order-statistics validation branches and rank variants', () => {
    const MEDIAN = getLookupBuiltin('MEDIAN')!
    const SMALL = getLookupBuiltin('SMALL')!
    const LARGE = getLookupBuiltin('LARGE')!
    const PERCENTILE = getLookupBuiltin('PERCENTILE')!
    const PERCENTILE_INC = getLookupBuiltin('PERCENTILE.INC')!
    const PERCENTILE_EXC = getLookupBuiltin('PERCENTILE.EXC')!
    const PERCENTRANK = getLookupBuiltin('PERCENTRANK')!
    const PERCENTRANK_INC = getLookupBuiltin('PERCENTRANK.INC')!
    const PERCENTRANK_EXC = getLookupBuiltin('PERCENTRANK.EXC')!
    const QUARTILE = getLookupBuiltin('QUARTILE')!
    const QUARTILE_INC = getLookupBuiltin('QUARTILE.INC')!
    const QUARTILE_EXC = getLookupBuiltin('QUARTILE.EXC')!
    const MODE_MULT = getLookupBuiltin('MODE.MULT')!
    const FREQUENCY = getLookupBuiltin('FREQUENCY')!
    const RANK = getLookupBuiltin('RANK')!
    const RANKEQ = getLookupBuiltin('RANK.EQ')!
    const RANKAVG = getLookupBuiltin('RANK.AVG')!

    const sample = cellRange([num(1), num(4), num(2), num(4), num(3)], 5, 1)
    const ordered = cellRange([num(1), num(2), num(4), num(7), num(8), num(9), num(10), num(12)], 8, 1)

    expect(MEDIAN(sample)).toEqual(num(3))
    expect(MEDIAN(num(7))).toEqual(num(7))
    expect(MEDIAN(cellRange([num(1), num(2), num(3), num(4)], 2, 2))).toEqual(num(2.5))
    expect(MEDIAN(cellRange([text('bad'), num(1)], 1, 2))).toEqual(err(ErrorCode.Value))

    expect(SMALL(sample, num(1))).toEqual(num(1))
    expect(SMALL(sample, num(4))).toEqual(num(4))
    expect(LARGE(sample, num(1))).toEqual(num(4))
    expect(LARGE(sample, num(4))).toEqual(num(2))
    expect(SMALL(sample, num(0))).toEqual(err(ErrorCode.Value))
    expect(LARGE(sample, num(99))).toEqual(err(ErrorCode.Value))
    expect(PERCENTILE(ordered, num(0.25))).toEqual(num(3.5))
    expect(PERCENTILE_INC(ordered, num(0.25))).toEqual(num(3.5))
    expect(PERCENTILE_EXC(ordered, num(0.25))).toEqual(num(2.5))
    expect(PERCENTILE(ordered, cellRange([num(0.5)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(PERCENTILE(cellRange([], 0, 0), num(0.5))).toEqual(err(ErrorCode.Value))
    expect(PERCENTILE_EXC(ordered, num(0))).toEqual(err(ErrorCode.Value))
    expect(PERCENTRANK(ordered, num(8))).toEqual(num(0.571))
    expect(PERCENTRANK_INC(ordered, num(8))).toEqual(num(0.571))
    expect(PERCENTRANK_EXC(ordered, num(8))).toEqual(num(0.555))
    expect(PERCENTRANK(ordered, num(5))).toEqual(num(0.333))
    expect(PERCENTRANK(ordered, num(8), num(1))).toEqual(num(0.5))
    expect(PERCENTRANK_EXC(ordered, num(5))).toEqual(num(0.37))
    expect(PERCENTRANK(cellRange([num(1)], 1, 1), num(1))).toEqual(err(ErrorCode.Value))
    expect(PERCENTRANK(ordered, num(8), num(0))).toEqual(err(ErrorCode.Value))
    expect(PERCENTRANK(ordered, num(100))).toEqual(err(ErrorCode.NA))
    expect(QUARTILE(ordered, num(0))).toEqual(num(1))
    expect(QUARTILE(ordered, num(1))).toEqual(num(3.5))
    expect(QUARTILE_INC(ordered, num(1))).toEqual(num(3.5))
    expect(QUARTILE_INC(ordered, num(4))).toEqual(num(12))
    expect(QUARTILE_EXC(ordered, num(1))).toEqual(num(2.5))
    expect(QUARTILE_EXC(ordered, num(0))).toEqual(err(ErrorCode.Value))
    expect(QUARTILE_EXC(ordered, num(4))).toEqual(err(ErrorCode.Value))
    expect(MODE_MULT(cellRange([num(1), num(2), num(2), num(3), num(3), num(4)], 6, 1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(2), num(3)],
    })
    expect(MODE_MULT(num(7), num(7), num(5))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 1,
      values: [num(7)],
    })
    expect(MODE_MULT(cellRange([num(1), num(2), num(3)], 3, 1))).toEqual(err(ErrorCode.NA))
    expect(MODE_MULT(text('bad'))).toEqual(err(ErrorCode.Value))
    expect(
      FREQUENCY(cellRange([num(79), num(85), num(78), num(85), num(50), num(81)], 6, 1), cellRange([num(60), num(80), num(90)], 3, 1)),
    ).toEqual({
      kind: 'array',
      rows: 4,
      cols: 1,
      values: [num(1), num(2), num(3), num(0)],
    })
    expect(FREQUENCY(num(5), num(4))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(0), num(1)],
    })

    expect(RANK(num(4), sample)).toEqual(num(1))
    expect(RANK(num(1), sample)).toEqual(num(5))
    expect(RANK(num(3), sample, num(1))).toEqual(num(3))
    expect(RANKEQ(num(4), sample)).toEqual(num(1))
    expect(RANKAVG(num(4), sample)).toEqual(num(1.5))
    expect(RANKEQ(num(8), sample)).toEqual(err(ErrorCode.NA))
  })

  it('returns value errors for missing required order-statistic arguments instead of throwing', () => {
    const LARGE = getLookupBuiltin('LARGE')!
    const MODE_MULT = getLookupBuiltin('MODE.MULT')!
    const PERCENTILE = getLookupBuiltin('PERCENTILE')!
    const PERCENTILE_EXC = getLookupBuiltin('PERCENTILE.EXC')!
    const PERCENTRANK = getLookupBuiltin('PERCENTRANK')!
    const PERCENTRANK_EXC = getLookupBuiltin('PERCENTRANK.EXC')!
    const PROB = getLookupBuiltin('PROB')!
    const QUARTILE = getLookupBuiltin('QUARTILE')!
    const QUARTILE_EXC = getLookupBuiltin('QUARTILE.EXC')!
    const RANK = getLookupBuiltin('RANK')!
    const RANKAVG = getLookupBuiltin('RANK.AVG')!
    const SMALL = getLookupBuiltin('SMALL')!
    const TRIMMEAN = getLookupBuiltin('TRIMMEAN')!
    const FREQUENCY = getLookupBuiltin('FREQUENCY')!

    const ordered = cellRange([num(1), num(2), num(4), num(7)], 4, 1)
    const probabilities = cellRange([num(0.1), num(0.2), num(0.3), num(0.4)], 4, 1)

    expect(Reflect.apply(SMALL, undefined, [ordered])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(LARGE, undefined, [])).toEqual(err(ErrorCode.Value))
    expect(SMALL(ordered, cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(LARGE(ordered, err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(SMALL(cellRange([], 0, 0), num(1))).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(PERCENTILE, undefined, [ordered])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(PERCENTILE_EXC, undefined, [])).toEqual(err(ErrorCode.Value))
    expect(PERCENTILE(err(ErrorCode.Ref), num(0.5))).toEqual(err(ErrorCode.Ref))
    expect(PERCENTILE(ordered, err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(PERCENTILE(ordered, text('bad'))).toEqual(err(ErrorCode.Value))
    expect(PERCENTILE(ordered, num(-0.01))).toEqual(err(ErrorCode.Value))
    expect(PERCENTILE(ordered, num(1.01))).toEqual(err(ErrorCode.Value))
    expect(PERCENTILE_EXC(ordered, num(0.01))).toEqual(err(ErrorCode.Value))
    expect(PERCENTILE_EXC(ordered, num(0.99))).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(PERCENTRANK, undefined, [ordered])).toEqual(err(ErrorCode.Value))
    expect(PERCENTRANK(err(ErrorCode.Ref), num(1))).toEqual(err(ErrorCode.Ref))
    expect(PERCENTRANK(ordered, err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(PERCENTRANK(ordered, num(1), err(ErrorCode.Div0))).toEqual(err(ErrorCode.Div0))
    expect(PERCENTRANK(ordered, cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(PERCENTRANK(ordered, num(1), cellRange([num(3)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(PERCENTRANK(ordered, text('bad'))).toEqual(err(ErrorCode.Value))
    expect(PERCENTRANK_EXC(ordered, num(0))).toEqual(err(ErrorCode.NA))
    expect(PERCENTRANK_EXC(ordered, num(1.5))).toEqual(num(0.3))
    expect(Reflect.apply(QUARTILE, undefined, [ordered])).toEqual(err(ErrorCode.Value))
    expect(QUARTILE(err(ErrorCode.Ref), num(1))).toEqual(err(ErrorCode.Ref))
    expect(QUARTILE(ordered, err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(QUARTILE(ordered, cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(QUARTILE(ordered, text('bad'))).toEqual(err(ErrorCode.Value))
    expect(QUARTILE(ordered, num(-1))).toEqual(err(ErrorCode.Value))
    expect(QUARTILE(ordered, num(5))).toEqual(err(ErrorCode.Value))
    expect(QUARTILE(cellRange([], 0, 0), num(1))).toEqual(err(ErrorCode.Value))
    expect(QUARTILE_EXC(ordered, num(2))).toEqual(num(3))
    expect(Reflect.apply(PROB, undefined, [ordered, probabilities])).toEqual(err(ErrorCode.Value))
    expect(PROB(err(ErrorCode.Ref), probabilities, num(1))).toEqual(err(ErrorCode.Value))
    expect(PROB(ordered, err(ErrorCode.NA), num(1))).toEqual(err(ErrorCode.Value))
    expect(PROB(ordered, cellRange([num(0.2), num(0.2), num(-0.1), num(0.7)], 4, 1), num(1))).toEqual(err(ErrorCode.Value))
    expect(PROB(ordered, cellRange([num(0.2), num(0.2), num(1.1), num(-0.5)], 4, 1), num(1))).toEqual(err(ErrorCode.Value))
    expect(PROB(ordered, probabilities, err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(Reflect.apply(RANK, undefined, [num(1)])).toEqual(err(ErrorCode.Value))
    expect(RANK(cellRange([num(1)], 1, 1), ordered)).toEqual(err(ErrorCode.Value))
    expect(RANK(num(1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(RANK(num(1), ordered, err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(RANK(text('bad'), ordered)).toEqual(err(ErrorCode.Value))
    expect(RANK(num(1), ordered, text('bad'))).toEqual(err(ErrorCode.Value))
    expect(RANK(num(1), ordered, num(2))).toEqual(err(ErrorCode.Value))
    expect(RANK(num(1), cellRange([], 0, 0))).toEqual(err(ErrorCode.NA))
    expect(RANKAVG(num(2), cellRange([num(2), num(2), num(3), num(4)], 4, 1))).toEqual(num(3.5))
    expect(Reflect.apply(TRIMMEAN, undefined, [ordered])).toEqual(err(ErrorCode.Value))
    expect(TRIMMEAN(ordered, cellRange([num(0.2)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(TRIMMEAN(ordered, num(-0.1))).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(MODE_MULT, undefined, [])).toEqual(err(ErrorCode.NA))
    expect(MODE_MULT(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(FREQUENCY(err(ErrorCode.Ref), ordered)).toEqual(err(ErrorCode.Ref))
    expect(FREQUENCY(ordered, err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
  })
})
