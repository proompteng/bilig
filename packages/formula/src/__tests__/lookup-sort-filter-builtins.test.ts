import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getLookupBuiltin, type RangeBuiltinArgument } from '../builtins/lookup.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

describe('lookup sort/filter builtins', () => {
  it('supports sorting, filtering, and uniqueness helpers', () => {
    const SORT = getLookupBuiltin('SORT')!
    const SORTBY = getLookupBuiltin('SORTBY')!
    const FILTER = getLookupBuiltin('FILTER')!
    const UNIQUE = getLookupBuiltin('UNIQUE')!

    expect(SORT(cellRange([num(3), num(1), num(2)], 3, 1))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [num(1), num(2), num(3)],
    })
    expect(SORTBY(cellRange([text('b'), text('a'), text('c')], 3, 1), cellRange([num(2), num(1), num(3)], 3, 1))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [text('a'), text('b'), text('c')],
    })
    expect(
      FILTER(
        cellRange([text('north'), text('south'), text('east'), text('west')], 4, 1),
        cellRange([bool(true), bool(false), bool(true), bool(false)], 4, 1),
      ),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [text('north'), text('east')],
    })
    expect(UNIQUE(cellRange([text('A'), text('a'), text('B'), text('A')], 4, 1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [text('A'), text('B')],
    })
  })

  it('preserves explicit errors and missing-arg validation in sort helpers', () => {
    const SORT = getLookupBuiltin('SORT')!
    const SORTBY = getLookupBuiltin('SORTBY')!

    expect(Reflect.apply(SORT, undefined, [])).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(SORTBY, undefined, [undefined, cellRange([num(1)], 1, 1)])).toEqual(err(ErrorCode.Value))
    expect(SORT(cellRange([num(3), num(1), num(2)], 3, 1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(SORTBY(cellRange([num(3), num(1), num(2)], 3, 1), err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
  })

  it('covers matrix sort modes and invalid sort criteria', () => {
    const SORT = getLookupBuiltin('SORT')!
    const SORTBY = getLookupBuiltin('SORTBY')!

    const matrix = cellRange([num(3), text('c'), num(1), text('a'), num(2), text('b')], 3, 2)
    expect(SORT(matrix, num(1), num(1), bool(false))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 2,
      values: [num(1), text('a'), num(2), text('b'), num(3), text('c')],
    })
    expect(SORT(cellRange([num(3), num(1), num(2), num(4)], 2, 2), num(1), num(-1), bool(true))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(3), num(1), num(2), num(4)],
    })

    expect(SORT(matrix, cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(SORT(matrix, num(9))).toEqual(err(ErrorCode.Value))
    expect(SORT(matrix, num(1), num(0))).toEqual(err(ErrorCode.Value))
    expect(SORT(matrix, num(1), num(1), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(SORT(cellRange([num(1), err(ErrorCode.Ref)], 2, 1))).toEqual(err(ErrorCode.Value))

    expect(SORTBY(matrix, cellRange([num(1), num(2), num(3)], 3, 1))).toEqual(err(ErrorCode.Value))
    expect(SORTBY(cellRange([text('b'), text('a'), text('c')], 3, 1), cellRange([num(2), num(1), num(1)], 3, 1), num(-1))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [text('b'), text('a'), text('c')],
    })
    expect(SORTBY(cellRange([text('b'), text('a')], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(SORTBY(cellRange([text('b'), text('a')], 2, 1), undefined)).toEqual(err(ErrorCode.Value))
    expect(SORTBY(cellRange([text('b'), text('a')], 2, 1), cellRange([num(1), num(2)], 2, 1), text('bad'))).toEqual(err(ErrorCode.Value))
    expect(SORTBY(cellRange([text('b'), text('a')], 2, 1), cellRange([err(ErrorCode.Ref), num(2)], 2, 1))).toEqual(err(ErrorCode.Value))
  })
})
