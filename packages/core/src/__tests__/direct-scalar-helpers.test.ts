import { describe, expect, it } from 'vitest'
import {
  ROW_PAIR_LEFT_DIV_RIGHT,
  ROW_PAIR_LEFT_MINUS_RIGHT,
  ROW_PAIR_LEFT_PLUS_RIGHT,
  ROW_PAIR_LEFT_TIMES_RIGHT,
  ROW_PAIR_RIGHT_DIV_LEFT,
  ROW_PAIR_RIGHT_MINUS_LEFT,
  directScalarLiteralNumericValue,
  evaluateRowPairDirectScalarCode,
  rowPairDirectScalarCode,
  rowPairDirectScalarCodeNeedsZeroGuard,
  singleInputAffineDirectScalar,
} from '../engine/services/direct-scalar-helpers.js'

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
})
