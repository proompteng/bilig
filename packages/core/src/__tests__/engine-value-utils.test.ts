import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import {
  areCellValuesEqual,
  cellValueDisplayText,
  emptyValue,
  errorValue,
  literalToValue,
  pivotItemMatches,
  writeLiteralToCellStore,
} from '../engine-value-utils.js'
import { CellStore } from '../cell-store.js'
import { StringPool } from '../string-pool.js'

describe('engine value utils', () => {
  it('materializes literal inputs into engine cell values', () => {
    const strings = new StringPool()

    expect(literalToValue(null, strings)).toEqual(emptyValue())
    expect(literalToValue(42, strings)).toEqual({ tag: ValueTag.Number, value: 42 })
    expect(literalToValue(true, strings)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(literalToValue('hello', strings)).toEqual({
      tag: ValueTag.String,
      value: 'hello',
      stringId: strings.intern('hello'),
    })
  })

  it('writes literal inputs directly into the cell store', () => {
    const strings = new StringPool()
    const store = new CellStore()
    const index = store.allocate(1, 0, 0)

    writeLiteralToCellStore(store, index, 42, strings)
    expect(store.getValue(index, (id) => strings.get(id))).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })

    writeLiteralToCellStore(store, index, true, strings)
    expect(store.getValue(index, (id) => strings.get(id))).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })

    writeLiteralToCellStore(store, index, 'hello', strings)
    expect(store.getValue(index, (id) => strings.get(id))).toEqual({
      tag: ValueTag.String,
      value: 'hello',
      stringId: strings.intern('hello'),
    })

    writeLiteralToCellStore(store, index, null, strings)
    expect(store.getValue(index, (id) => strings.get(id))).toEqual(emptyValue())
  })

  it('compares cell values with numeric identity semantics', () => {
    expect(areCellValuesEqual({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 })).toBe(true)
    expect(areCellValuesEqual({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: -0 })).toBe(false)
    expect(areCellValuesEqual({ tag: ValueTag.String, value: 'x', stringId: 1 }, { tag: ValueTag.String, value: 'x', stringId: 99 })).toBe(
      true,
    )
    expect(areCellValuesEqual({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Error, code: ErrorCode.Value })).toBe(false)
  })

  it('formats display text and matches pivot lookups case-insensitively', () => {
    expect(cellValueDisplayText({ tag: ValueTag.Number, value: -0 })).toBe('-0')
    expect(cellValueDisplayText({ tag: ValueTag.Boolean, value: true })).toBe('TRUE')
    expect(cellValueDisplayText(errorValue(ErrorCode.Name))).toBe('#Name!')

    expect(
      pivotItemMatches({ tag: ValueTag.String, value: 'Revenue', stringId: 1 }, { tag: ValueTag.String, value: ' revenue ', stringId: 2 }),
    ).toBe(true)
    expect(pivotItemMatches({ tag: ValueTag.Boolean, value: true }, { tag: ValueTag.String, value: ' true ', stringId: 3 })).toBe(true)
  })

  it('formats empty cells and rejects mismatched boolean and pivot values', () => {
    expect(cellValueDisplayText(emptyValue())).toBe('')
    expect(areCellValuesEqual({ tag: ValueTag.Boolean, value: true }, { tag: ValueTag.Boolean, value: false })).toBe(false)
    expect(pivotItemMatches({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Boolean, value: true })).toBe(false)
  })
})
