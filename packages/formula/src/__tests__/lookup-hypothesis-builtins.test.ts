import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getLookupBuiltin, type RangeBuiltinArgument } from '../builtins/lookup.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

describe('lookup hypothesis builtins', () => {
  it('supports legacy and modern hypothesis test wrappers', () => {
    const CHITEST = getLookupBuiltin('CHITEST')!
    const FTEST = getLookupBuiltin('FTEST')!
    const ZTEST = getLookupBuiltin('ZTEST')!
    const TTEST = getLookupBuiltin('TTEST')!

    expect(
      CHITEST(cellRange([num(10), num(20), num(20), num(40)], 2, 2), cellRange([num(15), num(15), num(15), num(45)], 2, 2)),
    ).toMatchObject({ tag: ValueTag.Number, value: expect.any(Number) })
    expect(FTEST(cellRange([num(1), num(2), num(3)], 3, 1), cellRange([num(2), num(4), num(6)], 3, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.any(Number),
    })
    expect(ZTEST(cellRange([num(1), num(2), num(3)], 3, 1), num(2))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.any(Number),
    })
    expect(TTEST(cellRange([num(1), num(2), num(3)], 3, 1), cellRange([num(2), num(4), num(6)], 3, 1), num(2), num(2))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.any(Number),
    })
  })

  it('returns value errors for missing required hypothesis test args', () => {
    const CHITEST = getLookupBuiltin('CHITEST')!
    const FTEST = getLookupBuiltin('FTEST')!
    const ZTEST = getLookupBuiltin('ZTEST')!
    const TTEST = getLookupBuiltin('TTEST')!

    expect(Reflect.apply(CHITEST, undefined, [cellRange([num(1)], 1, 1)])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(FTEST, undefined, [cellRange([num(1)], 1, 1)])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(ZTEST, undefined, [cellRange([num(1)], 1, 1)])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(TTEST, undefined, [cellRange([num(1)], 1, 1), cellRange([num(2)], 1, 1)])).toEqual(err(ErrorCode.Value))
  })

  it('validates chi-square matrix shape and expected frequencies', () => {
    const CHISQ_TEST = getLookupBuiltin('CHISQ.TEST')!
    const LEGACY_CHITEST = getLookupBuiltin('LEGACY.CHITEST')!

    expect(CHISQ_TEST(cellRange([num(1), num(2)], 1, 2), cellRange([num(1), num(2), num(3), num(4)], 2, 2))).toEqual(err(ErrorCode.NA))
    expect(CHISQ_TEST(cellRange([num(1)], 1, 1), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.NA))
    expect(CHISQ_TEST(cellRange([num(-1), num(2), num(3), num(4)], 2, 2), cellRange([num(1), num(2), num(3), num(4)], 2, 2))).toEqual(
      err(ErrorCode.Value),
    )
    expect(CHISQ_TEST(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(1), num(0), num(3), num(4)], 2, 2))).toEqual(
      err(ErrorCode.Div0),
    )
    expect(CHISQ_TEST(cellRange([err(ErrorCode.Ref), num(2)], 1, 2), cellRange([num(1), num(2)], 1, 2))).toEqual(err(ErrorCode.Value))
    expect(CHISQ_TEST(num(1), cellRange([num(1), num(2)], 1, 2))).toEqual(err(ErrorCode.NA))
    expect(LEGACY_CHITEST(cellRange([num(5), num(7)], 1, 2), cellRange([num(6), num(6)], 1, 2))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.any(Number),
    })
  })

  it('validates F.TEST and Z.TEST sample coercion branches', () => {
    const F_TEST = getLookupBuiltin('F.TEST')!
    const Z_TEST = getLookupBuiltin('Z.TEST')!

    expect(F_TEST(num(1), num(2))).toEqual(err(ErrorCode.Div0))
    expect(F_TEST(cellRange([num(1), num(1), num(1)], 3, 1), cellRange([num(2), num(3), num(4)], 3, 1))).toEqual(err(ErrorCode.Div0))
    expect(F_TEST(cellRange([err(ErrorCode.NA), num(2)], 2, 1), cellRange([num(2), num(3)], 2, 1))).toEqual(err(ErrorCode.NA))
    expect(F_TEST(text('bad'), cellRange([num(2), num(3)], 2, 1))).toEqual(err(ErrorCode.Value))

    expect(Z_TEST(cellRange([num(1), num(2), num(3)], 3, 1), num(2), num(1.5))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.any(Number),
    })
    expect(Z_TEST(cellRange([num(1), num(2), num(3)], 3, 1), cellRange([num(2)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(Z_TEST(cellRange([], 0, 0), num(2))).toEqual(err(ErrorCode.Value))
    expect(Z_TEST(cellRange([num(1), num(1), num(1)], 3, 1), num(1))).toEqual(err(ErrorCode.Div0))
    expect(Z_TEST(cellRange([num(1), num(2), num(3)], 3, 1), num(2), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Div0))
    expect(Z_TEST(cellRange([num(1), num(2), num(3)], 3, 1), num(2), num(0))).toEqual(err(ErrorCode.Div0))
    expect(Z_TEST(cellRange([err(ErrorCode.Ref), num(2)], 2, 1), num(2))).toEqual(err(ErrorCode.Ref))
    expect(Z_TEST(text('bad'), num(2))).toEqual(err(ErrorCode.Value))
  })

  it('validates T.TEST paired, pooled, and Welch branches', () => {
    const T_TEST = getLookupBuiltin('T.TEST')!
    const first = cellRange([num(1), num(2), num(5), num(8)], 4, 1)
    const second = cellRange([num(2), num(4), num(6), num(12)], 4, 1)

    expect(T_TEST(first, second, num(1), num(1))).toMatchObject({ tag: ValueTag.Number, value: expect.any(Number) })
    expect(T_TEST(first, second, num(2), num(2))).toMatchObject({ tag: ValueTag.Number, value: expect.any(Number) })
    expect(T_TEST(first, second, num(2), num(3))).toMatchObject({ tag: ValueTag.Number, value: expect.any(Number) })
    expect(T_TEST(cellRange([num(1), num(2)], 2, 1), cellRange([num(1)], 1, 1), num(1), num(1))).toEqual(err(ErrorCode.NA))
    expect(T_TEST(cellRange([num(1)], 1, 1), cellRange([num(2)], 1, 1), num(1), num(1))).toEqual(err(ErrorCode.Div0))
    expect(T_TEST(cellRange([num(1), num(2)], 2, 1), cellRange([num(1), num(2)], 2, 1), num(1), num(1))).toEqual(err(ErrorCode.Div0))
    expect(T_TEST(cellRange([num(1), num(1)], 2, 1), cellRange([num(2), num(3)], 2, 1), num(2), num(2))).toEqual(err(ErrorCode.Div0))
    expect(T_TEST(first, second, num(3), num(2))).toEqual(err(ErrorCode.Value))
    expect(T_TEST(first, second, num(2), num(4))).toEqual(err(ErrorCode.Value))
    expect(T_TEST(cellRange([err(ErrorCode.Name), num(2)], 2, 1), second, num(2), num(2))).toEqual(err(ErrorCode.Name))
    expect(T_TEST(text('bad'), second, num(2), num(2))).toEqual(err(ErrorCode.Value))
  })
})
