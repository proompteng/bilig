import { describe, expect, it } from 'vitest'
import {
  ROW_PAIR_LEFT_DIV_RIGHT,
  ROW_PAIR_LEFT_MINUS_RIGHT,
  ROW_PAIR_LEFT_PLUS_RIGHT,
  ROW_PAIR_LEFT_TIMES_RIGHT,
  ROW_PAIR_RIGHT_DIV_LEFT,
  ROW_PAIR_RIGHT_MINUS_LEFT,
  directScalarLiteralNumericValue,
  directScalarCellNumber,
  directScalarDeltaFromNumbers,
  directScalarDeltaFromValues,
  directScalarValueNumber,
  evaluateDirectScalarNumber,
  evaluateDirectScalarWithReplacementNumbers,
  evaluateRowPairDirectScalarCode,
  rowPairDirectScalarCode,
  rowPairDirectScalarCodeNeedsZeroGuard,
  singleInputAffineDirectScalar,
} from '../engine/services/direct-scalar-helpers.js'
import { ErrorCode, ValueTag } from '@bilig/protocol'

const readCellNumberForDirectScalarTests = (cellIndex: number) => (cellIndex === 2 ? 10 : undefined)

describe('direct scalar helpers', () => {
  it('normalizes literal numeric operands for direct scalar paths', () => {
    expect(directScalarLiteralNumericValue(null)).toBe(0)
    expect(directScalarLiteralNumericValue(-0)).toBe(0)
    expect(directScalarLiteralNumericValue(true)).toBe(1)
    expect(directScalarLiteralNumericValue(false)).toBe(0)
    expect(directScalarLiteralNumericValue('1')).toBeUndefined()
    expect(directScalarLiteralNumericValue(undefined)).toBeUndefined()
  })

  it('recognizes affine single-input scalar formulas', () => {
    expect(singleInputAffineDirectScalar({ kind: 'abs', operand: { kind: 'cell', cellIndex: 1 } }, 1)).toBeNull()
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '+', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'literal-number', value: 2 } },
        1,
      ),
    ).toEqual({ scale: 1, offset: 2 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '-', left: { kind: 'literal-number', value: 2 }, right: { kind: 'cell', cellIndex: 1 } },
        1,
      ),
    ).toEqual({ scale: -1, offset: 2 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '*', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'literal-number', value: 2 } },
        1,
      ),
    ).toEqual({ scale: 2, offset: 0 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '/', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'literal-number', value: 0 } },
        1,
      ),
    ).toBeNull()
  })

  it('encodes and evaluates row-pair scalar formulas', () => {
    expect(
      rowPairDirectScalarCode(
        { kind: 'binary', operator: '+', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'cell', cellIndex: 2 } },
        1,
        2,
      ),
    ).toBe(ROW_PAIR_LEFT_PLUS_RIGHT)
    expect(
      rowPairDirectScalarCode(
        { kind: 'binary', operator: '-', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'cell', cellIndex: 2 } },
        1,
        2,
      ),
    ).toBe(ROW_PAIR_LEFT_MINUS_RIGHT)
    expect(
      rowPairDirectScalarCode(
        { kind: 'binary', operator: '-', left: { kind: 'cell', cellIndex: 2 }, right: { kind: 'cell', cellIndex: 1 } },
        1,
        2,
      ),
    ).toBe(ROW_PAIR_RIGHT_MINUS_LEFT)
    expect(
      rowPairDirectScalarCode(
        { kind: 'binary', operator: '*', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'cell', cellIndex: 2 } },
        1,
        2,
      ),
    ).toBe(ROW_PAIR_LEFT_TIMES_RIGHT)
    expect(
      rowPairDirectScalarCode(
        { kind: 'binary', operator: '/', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'cell', cellIndex: 2 } },
        1,
        2,
      ),
    ).toBe(ROW_PAIR_LEFT_DIV_RIGHT)
    expect(
      rowPairDirectScalarCode(
        { kind: 'binary', operator: '/', left: { kind: 'cell', cellIndex: 2 }, right: { kind: 'cell', cellIndex: 1 } },
        1,
        2,
      ),
    ).toBe(ROW_PAIR_RIGHT_DIV_LEFT)

    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_PLUS_RIGHT, 8, 2)).toBe(10)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_MINUS_RIGHT, 8, 2)).toBe(6)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_RIGHT_MINUS_LEFT, 8, 2)).toBe(-6)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_TIMES_RIGHT, 8, 2)).toBe(16)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_DIV_RIGHT, 8, 2)).toBe(4)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_RIGHT_DIV_LEFT, 8, 2)).toBe(0.25)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_DIV_RIGHT, 8, 0)).toBeUndefined()
    expect(evaluateRowPairDirectScalarCode(0, 8, 2)).toBeUndefined()

    expect(rowPairDirectScalarCodeNeedsZeroGuard(ROW_PAIR_LEFT_DIV_RIGHT)).toBe(true)
    expect(rowPairDirectScalarCodeNeedsZeroGuard(ROW_PAIR_RIGHT_DIV_LEFT)).toBe(true)
    expect(rowPairDirectScalarCodeNeedsZeroGuard(ROW_PAIR_LEFT_PLUS_RIGHT)).toBe(false)
  })

  it('normalizes direct scalar cell values and store slots', () => {
    expect(directScalarValueNumber({ tag: ValueTag.Number, value: -0 })).toBe(0)
    expect(directScalarValueNumber({ tag: ValueTag.Boolean, value: true })).toBe(1)
    expect(directScalarValueNumber({ tag: ValueTag.Boolean, value: false })).toBe(0)
    expect(directScalarValueNumber({ tag: ValueTag.Empty })).toBe(0)
    expect(directScalarValueNumber({ tag: ValueTag.String, value: '1' })).toBe(1)
    expect(directScalarValueNumber({ tag: ValueTag.String, value: '61,111' })).toBe(61111)
    expect(directScalarValueNumber({ tag: ValueTag.String, value: '' })).toBe(0)
    expect(directScalarValueNumber({ tag: ValueTag.String, value: 'not numeric' })).toBeUndefined()
    expect(directScalarValueNumber({ tag: ValueTag.Error, code: ErrorCode.VALUE })).toBeUndefined()

    const cellStore = {
      tags: [ValueTag.Number, ValueTag.Boolean, ValueTag.Empty, ValueTag.String, ValueTag.Error],
      numbers: [-0, 1, 0, 0, 0],
    }
    expect(directScalarCellNumber(cellStore, undefined)).toBe(0)
    expect(directScalarCellNumber(cellStore, 0)).toBe(0)
    expect(directScalarCellNumber(cellStore, 1)).toBe(1)
    expect(directScalarCellNumber(cellStore, 2)).toBe(0)
    expect(directScalarCellNumber(cellStore, 3)).toBeUndefined()
    expect(directScalarCellNumber(cellStore, 4)).toBeUndefined()
  })

  it('evaluates direct scalar formulas with a replacement cell number', () => {
    const touched = { value: false }

    expect(
      evaluateDirectScalarNumber(
        { kind: 'binary', operator: '*', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'cell', cellIndex: 2 } },
        1,
        4,
        touched,
        readCellNumberForDirectScalarTests,
      ),
    ).toBe(40)
    expect(touched.value).toBe(true)
    expect(
      evaluateDirectScalarNumber(
        {
          kind: 'binary',
          operator: '+',
          left: { kind: 'cell', cellIndex: 1 },
          right: { kind: 'literal-number', value: 2 },
          resultOffset: 3,
        },
        1,
        4,
        { value: false },
        readCellNumberForDirectScalarTests,
      ),
    ).toBe(9)
    expect(
      evaluateDirectScalarNumber(
        { kind: 'abs', operand: { kind: 'cell', cellIndex: 1 } },
        1,
        -4,
        { value: false },
        readCellNumberForDirectScalarTests,
      ),
    ).toBe(4)
    expect(
      evaluateDirectScalarNumber(
        { kind: 'binary', operator: '/', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'literal-number', value: 0 } },
        1,
        4,
        { value: false },
        readCellNumberForDirectScalarTests,
      ),
    ).toBeUndefined()
  })

  it('evaluates direct scalar formulas with one or two replacement cell numbers', () => {
    expect(
      evaluateDirectScalarWithReplacementNumbers(
        { kind: 'binary', operator: '*', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'cell', cellIndex: 2 } },
        1,
        4,
        readCellNumberForDirectScalarTests,
      ),
    ).toBe(40)
    expect(
      evaluateDirectScalarWithReplacementNumbers(
        { kind: 'binary', operator: '+', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'cell', cellIndex: 2 } },
        1,
        4,
        readCellNumberForDirectScalarTests,
        2,
        6,
      ),
    ).toBe(10)
    expect(
      evaluateDirectScalarWithReplacementNumbers(
        { kind: 'abs', operand: { kind: 'cell', cellIndex: 2 } },
        1,
        4,
        readCellNumberForDirectScalarTests,
        2,
        -6,
      ),
    ).toBe(6)
    expect(
      evaluateDirectScalarWithReplacementNumbers(
        { kind: 'binary', operator: '/', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'cell', cellIndex: 2 } },
        1,
        4,
        readCellNumberForDirectScalarTests,
        2,
        0,
      ),
    ).toBeUndefined()
  })

  it('computes direct scalar numeric deltas from values and numbers', () => {
    const timesOtherCell = {
      kind: 'binary',
      operator: '*',
      left: { kind: 'cell', cellIndex: 1 },
      right: { kind: 'cell', cellIndex: 2 },
    } as const
    expect(directScalarDeltaFromNumbers(timesOtherCell, 1, 2, 4, readCellNumberForDirectScalarTests)).toBe(20)
    expect(
      directScalarDeltaFromValues(
        timesOtherCell,
        1,
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        readCellNumberForDirectScalarTests,
      ),
    ).toBe(20)

    expect(
      directScalarDeltaFromNumbers(
        { kind: 'binary', operator: '-', left: { kind: 'literal-number', value: 2 }, right: { kind: 'cell', cellIndex: 1 } },
        1,
        2,
        4,
        readCellNumberForDirectScalarTests,
      ),
    ).toBe(-2)
    expect(
      directScalarDeltaFromNumbers({ kind: 'abs', operand: { kind: 'cell', cellIndex: 1 } }, 1, -3, 4, readCellNumberForDirectScalarTests),
    ).toBe(1)
    expect(
      directScalarDeltaFromNumbers(
        { kind: 'binary', operator: '/', left: { kind: 'literal-number', value: 8 }, right: { kind: 'cell', cellIndex: 1 } },
        1,
        2,
        4,
        readCellNumberForDirectScalarTests,
      ),
    ).toBe(-2)
    expect(
      directScalarDeltaFromNumbers(
        { kind: 'binary', operator: '/', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'cell', cellIndex: 2 } },
        3,
        2,
        4,
        readCellNumberForDirectScalarTests,
      ),
    ).toBeUndefined()
    expect(
      directScalarDeltaFromValues(
        timesOtherCell,
        1,
        { tag: ValueTag.String, value: 'x' },
        { tag: ValueTag.Number, value: 4 },
        readCellNumberForDirectScalarTests,
      ),
    ).toBeUndefined()
  })
})
