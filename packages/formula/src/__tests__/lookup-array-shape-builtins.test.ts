import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getLookupBuiltin, type RangeBuiltinArgument } from '../builtins/lookup.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const empty = (): CellValue => ({ tag: ValueTag.Empty })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

describe('lookup array-shape builtins', () => {
  it('formats arrays, indexes vectors, and offsets windows', () => {
    const AREAS = getLookupBuiltin('AREAS')!
    const ARRAYTOTEXT = getLookupBuiltin('ARRAYTOTEXT')!
    const COLUMNS = getLookupBuiltin('COLUMNS')!
    const ROWS = getLookupBuiltin('ROWS')!
    const INDEX = getLookupBuiltin('INDEX')!
    const OFFSET = getLookupBuiltin('OFFSET')!

    const matrix = cellRange([num(1), text('a'), bool(true), empty(), num(5), text('b')], 2, 3)

    expect(AREAS(matrix)).toEqual(num(1))
    expect(COLUMNS(matrix)).toEqual(num(3))
    expect(ROWS(matrix)).toEqual(num(2))
    expect(ARRAYTOTEXT(matrix)).toEqual(text('1\ta\tTRUE;\t5\tb'))
    expect(ARRAYTOTEXT(matrix, num(1))).toEqual(text('{1, "a", TRUE;, 5, "b"}'))
    expect(ARRAYTOTEXT(cellRange([err(ErrorCode.Ref)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(ARRAYTOTEXT(matrix, cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))

    expect(INDEX(matrix, num(2), num(3))).toEqual(text('b'))
    expect(INDEX(cellRange([num(1), num(2), num(3)], 1, 3), num(2))).toEqual(num(2))
    expect(INDEX(matrix, err(ErrorCode.NA), num(1))).toEqual(err(ErrorCode.NA))
    expect(INDEX(matrix, num(0), num(1))).toEqual(err(ErrorCode.Ref))
    expect(INDEX({ kind: 'range', refKind: 'rows', values: [num(1)], rows: 1, cols: 1 }, num(1), num(1))).toEqual(err(ErrorCode.Value))

    expect(OFFSET(matrix, num(0), num(1), num(2), num(2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('a'), bool(true), num(5), text('b')],
    })
    expect(OFFSET(matrix, num(1), num(2), num(1), num(1))).toEqual(text('b'))
    expect(OFFSET(matrix, err(ErrorCode.Ref), num(0))).toEqual(err(ErrorCode.Ref))
    expect(OFFSET(matrix, num(0), num(0), num(0), num(1))).toEqual(err(ErrorCode.Value))
    expect(OFFSET(matrix, num(2), num(0))).toEqual(err(ErrorCode.Ref))
  })

  it('slices, chooses, stacks, flattens, and wraps arrays', () => {
    const TAKE = getLookupBuiltin('TAKE')!
    const DROP = getLookupBuiltin('DROP')!
    const CHOOSECOLS = getLookupBuiltin('CHOOSECOLS')!
    const CHOOSEROWS = getLookupBuiltin('CHOOSEROWS')!
    const TRANSPOSE = getLookupBuiltin('TRANSPOSE')!
    const HSTACK = getLookupBuiltin('HSTACK')!
    const VSTACK = getLookupBuiltin('VSTACK')!
    const TOCOL = getLookupBuiltin('TOCOL')!
    const TOROW = getLookupBuiltin('TOROW')!
    const WRAPROWS = getLookupBuiltin('WRAPROWS')!
    const WRAPCOLS = getLookupBuiltin('WRAPCOLS')!

    const matrix = cellRange([num(1), num(2), num(3), num(4), empty(), num(6)], 2, 3)

    expect(TAKE(matrix, num(1), num(-2))).toEqual({ kind: 'array', rows: 1, cols: 2, values: [num(2), num(3)] })
    expect(TAKE(matrix, num(0), num(1))).toEqual(err(ErrorCode.Value))
    expect(TAKE(matrix, cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(DROP(matrix, num(1), num(-1))).toEqual({ kind: 'array', rows: 1, cols: 2, values: [num(4), empty()] })
    expect(DROP(matrix, num(2))).toEqual(err(ErrorCode.Value))
    expect(DROP(matrix, text('bad'))).toEqual(err(ErrorCode.Value))

    expect(CHOOSECOLS(matrix, num(3), num(1))).toEqual({ kind: 'array', rows: 2, cols: 2, values: [num(3), num(1), num(6), num(4)] })
    expect(CHOOSEROWS(matrix, num(2), num(1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [num(4), empty(), num(6), num(1), num(2), num(3)],
    })
    expect(CHOOSECOLS(matrix)).toEqual(err(ErrorCode.Value))
    expect(CHOOSEROWS(matrix, err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(CHOOSEROWS(matrix, num(3))).toEqual(err(ErrorCode.Value))

    expect(TRANSPOSE(matrix)).toEqual({ kind: 'array', rows: 3, cols: 2, values: [num(1), num(4), num(2), empty(), num(3), num(6)] })
    expect(TRANSPOSE(cellRange([num(9)], 1, 1))).toEqual(num(9))
    expect(HSTACK(cellRange([num(1), num(2)], 1, 2), cellRange([num(3), num(4), num(5), num(6)], 2, 2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 4,
      values: [num(1), num(2), num(3), num(4), num(1), num(2), num(5), num(6)],
    })
    expect(VSTACK(cellRange([num(1), num(2)], 2, 1), cellRange([num(3), num(4), num(5), num(6)], 2, 2))).toEqual({
      kind: 'array',
      rows: 4,
      cols: 2,
      values: [num(1), num(1), num(2), num(2), num(3), num(4), num(5), num(6)],
    })
    expect(HSTACK()).toEqual(err(ErrorCode.Value))
    expect(VSTACK(cellRange([num(1), num(2)], 1, 2), cellRange([num(3), num(4), num(5)], 1, 3))).toEqual(err(ErrorCode.Value))

    expect(TOCOL(matrix, num(1), bool(false))).toEqual({
      kind: 'array',
      rows: 5,
      cols: 1,
      values: [num(1), num(2), num(3), num(4), num(6)],
    })
    expect(TOROW(matrix, num(0), bool(true))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 6,
      values: [num(1), num(4), num(2), empty(), num(3), num(6)],
    })
    expect(TOCOL(matrix, num(2))).toEqual(err(ErrorCode.Value))
    expect(TOROW(matrix, num(0), text('bad'))).toEqual(err(ErrorCode.Value))

    expect(WRAPROWS(cellRange([num(1), num(2), num(3)], 3, 1), num(2), text('pad'))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(1), num(2), num(3), text('pad')],
    })
    expect(WRAPCOLS(cellRange([num(1), num(2), num(3)], 3, 1), num(2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(1), num(3), num(2), err(ErrorCode.NA)],
    })
    expect(WRAPROWS(matrix, num(0))).toEqual(err(ErrorCode.Value))
    expect(WRAPCOLS(matrix, num(2), text('pad'), text('bad'))).toEqual(err(ErrorCode.Value))
  })
})
