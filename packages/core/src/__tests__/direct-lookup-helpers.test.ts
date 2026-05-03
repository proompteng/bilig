import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { RuntimeDirectLookupDescriptor } from '../engine/runtime-state.js'
import {
  approximateRepeatedUniformLookupCurrentResult,
  approximateUniformLookupCurrentResult,
  approximateUniformLookupNumericResult,
  canSkipUniformApproximateNumericTailWrite,
  canSkipUniformApproximateNumericTailWriteFromCurrentResult,
  canSkipUniformExactNumericTailWriteFromCurrentResult,
  directLookupRowBounds,
  directLookupVersionMatches,
  exactLookupLiteralNumericValue,
  exactUniformLookupCurrentResult,
  exactUniformLookupNumericResult,
  normalizeApproximateNumericValue,
  normalizeApproximateTextValue,
  normalizeExactLookupKey,
  normalizeExactNumericValue,
  sameExactNumericValue,
  withOptionalLookupStringIds,
} from '../engine/services/direct-lookup-helpers.js'

function exactUniform(overrides: Partial<Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>> = {}) {
  return {
    kind: 'exact-uniform-numeric',
    operandCellIndex: 1,
    sheetName: 'Sheet1',
    sheetId: 7,
    rowStart: 0,
    rowEnd: 4,
    col: 0,
    length: 5,
    columnVersion: 2,
    structureVersion: 3,
    sheetColumnVersions: new Uint32Array([2]),
    start: 1,
    step: 1,
    searchMode: 1,
    ...overrides,
  } satisfies Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>
}

function approximateUniform(overrides: Partial<Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>> = {}) {
  return {
    kind: 'approximate-uniform-numeric',
    operandCellIndex: 1,
    sheetName: 'Sheet1',
    sheetId: 7,
    rowStart: 0,
    rowEnd: 4,
    col: 0,
    length: 5,
    columnVersion: 2,
    structureVersion: 3,
    sheetColumnVersions: new Uint32Array([2]),
    start: 1,
    step: 1,
    matchMode: 1,
    ...overrides,
  } satisfies Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>
}

const lookupString = (id: number) => (id === 1 ? 'alpha' : 'beta')

describe('direct lookup helpers', () => {
  it('normalizes exact and approximate lookup operands', () => {
    expect(normalizeExactLookupKey({ tag: ValueTag.Empty }, lookupString)).toBe('e:')
    expect(normalizeExactLookupKey({ tag: ValueTag.Number, value: -0 }, lookupString)).toBe('n:0')
    expect(normalizeExactLookupKey({ tag: ValueTag.Boolean, value: true }, lookupString)).toBe('b:1')
    expect(normalizeExactLookupKey({ tag: ValueTag.Boolean, value: false }, lookupString)).toBe('b:0')
    expect(normalizeExactLookupKey({ tag: ValueTag.String, value: 'local' }, lookupString)).toBe('s:LOCAL')
    expect(normalizeExactLookupKey({ tag: ValueTag.String, value: 'ignored' }, lookupString, 1)).toBe('s:ALPHA')
    expect(normalizeExactLookupKey({ tag: ValueTag.Error, code: ErrorCode.NA }, lookupString)).toBeUndefined()

    expect(normalizeExactNumericValue({ tag: ValueTag.Number, value: -0 })).toBe(0)
    expect(normalizeExactNumericValue({ tag: ValueTag.Empty })).toBeUndefined()
    expect(sameExactNumericValue(-0, 0)).toBe(true)
    expect(exactLookupLiteralNumericValue(-0)).toBe(0)
    expect(exactLookupLiteralNumericValue('1')).toBeUndefined()

    expect(normalizeApproximateNumericValue({ tag: ValueTag.Empty })).toBe(0)
    expect(normalizeApproximateNumericValue({ tag: ValueTag.Number, value: -0 })).toBe(0)
    expect(normalizeApproximateNumericValue({ tag: ValueTag.Boolean, value: true })).toBe(1)
    expect(normalizeApproximateNumericValue({ tag: ValueTag.String, value: '1' })).toBeUndefined()
    expect(normalizeApproximateNumericValue({ tag: ValueTag.Error, code: ErrorCode.NA })).toBeUndefined()

    expect(normalizeApproximateTextValue({ tag: ValueTag.Empty }, lookupString)).toBe('')
    expect(normalizeApproximateTextValue({ tag: ValueTag.String, value: 'local' }, lookupString)).toBe('LOCAL')
    expect(normalizeApproximateTextValue({ tag: ValueTag.String, value: 'ignored' }, lookupString, 2)).toBe('BETA')
    expect(normalizeApproximateTextValue({ tag: ValueTag.Number, value: 1 }, lookupString)).toBeUndefined()
  })

  it('computes uniform exact and approximate lookup results', () => {
    expect(exactUniformLookupNumericResult(exactUniform(), 3)).toBe(3)
    expect(exactUniformLookupNumericResult(exactUniform(), 3.5)).toBeUndefined()
    expect(exactUniformLookupNumericResult(exactUniform({ step: -1, start: 5 }), 3)).toBe(3)
    expect(exactUniformLookupNumericResult(exactUniform({ step: 2, start: 2 }), 8)).toBe(4)
    expect(exactUniformLookupNumericResult(exactUniform({ step: 2, start: 2 }), 9)).toBeUndefined()
    expect(
      exactUniformLookupNumericResult(exactUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 9, columnVersion: 4 } }), 9),
    ).toBe(5)
    expect(
      exactUniformLookupNumericResult(exactUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 9, columnVersion: 4 } }), 5),
    ).toBeUndefined()
    expect(exactUniformLookupCurrentResult(exactUniform(), 8)).toEqual({ kind: 'error', code: ErrorCode.NA })

    expect(approximateUniformLookupNumericResult(approximateUniform(), 3.2)).toBe(3)
    expect(approximateUniformLookupNumericResult(approximateUniform(), 0)).toBeUndefined()
    expect(approximateUniformLookupNumericResult(approximateUniform({ start: 10, step: -1, matchMode: -1 }), 8.8)).toBe(2)
    expect(approximateUniformLookupNumericResult(approximateUniform({ start: 10, step: -1, matchMode: -1 }), 11)).toBeUndefined()
    expect(approximateUniformLookupNumericResult(approximateUniform({ step: 2 }), 4.9)).toBe(2)
    expect(
      approximateUniformLookupNumericResult(
        approximateUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 9, columnVersion: 4 } }),
        6,
      ),
    ).toBe(4)
    expect(
      approximateUniformLookupNumericResult(
        approximateUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 9, columnVersion: 4 } }),
        9,
      ),
    ).toBe(5)
    expect(approximateUniformLookupCurrentResult(approximateUniform(), 0)).toEqual({ kind: 'error', code: ErrorCode.NA })
    expect(approximateUniformLookupCurrentResult(approximateUniform({ start: 10, step: -1, matchMode: -1 }), 11)).toEqual({
      kind: 'error',
      code: ErrorCode.NA,
    })
  })

  it('covers skip guards, row bounds, version checks, and optional string ids', () => {
    const exact = exactUniform()
    const approximate = approximateUniform()
    const cellStore = { tags: [ValueTag.Empty, ValueTag.Number], numbers: [0, 2] }

    expect(directLookupVersionMatches({ structureVersion: 3, columnVersions: [2] }, exact)).toBe(true)
    expect(directLookupVersionMatches({ structureVersion: 4, columnVersions: [2] }, exact)).toBe(false)
    expect(
      directLookupVersionMatches(
        { structureVersion: 3, columnVersions: [4] },
        exactUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 9, columnVersion: 4 } }),
      ),
    ).toBe(true)

    expect(canSkipUniformApproximateNumericTailWrite(approximate, 4, 2, 5, 6)).toBe(true)
    expect(canSkipUniformApproximateNumericTailWrite(approximate, 3, 2, 5, 6)).toBe(false)
    expect(canSkipUniformApproximateNumericTailWrite(approximateUniform({ start: 5, step: -1, matchMode: -1 }), 4, 3, 1, 0)).toBe(true)
    expect(canSkipUniformApproximateNumericTailWriteFromCurrentResult(cellStore, 1, approximate, 4, 5, 6)).toBe(true)
    expect(canSkipUniformApproximateNumericTailWriteFromCurrentResult(cellStore, 1, approximate, 4, 6, 5)).toBe(false)
    expect(canSkipUniformExactNumericTailWriteFromCurrentResult(cellStore, 1, exact, 4, 5, 6)).toBe(true)
    expect(canSkipUniformExactNumericTailWriteFromCurrentResult(cellStore, 1, exact, 3, 5, 6)).toBe(false)

    expect(directLookupRowBounds(exact)).toEqual({ rowStart: 0, rowEnd: 4 })
    expect(
      directLookupRowBounds({
        kind: 'exact',
        operandCellIndex: 1,
        searchMode: 1,
        prepared: {
          sheetName: 'Sheet1',
          rowStart: 2,
          rowEnd: 8,
          col: 3,
          length: 7,
          columnVersion: 1,
          structureVersion: 1,
          sheetColumnVersions: new Uint32Array([1]),
          comparableKind: 'numeric',
          uniformStart: undefined,
          uniformStep: undefined,
          firstPositions: new Map(),
          lastPositions: new Map(),
          firstNumericPositions: undefined,
          lastNumericPositions: undefined,
          firstTextPositions: undefined,
          lastTextPositions: undefined,
        },
      }),
    ).toEqual({ rowStart: 2, rowEnd: 8 })

    expect(
      approximateRepeatedUniformLookupCurrentResult(
        { length: 6, repeatedUniformStart: 10, repeatedUniformStep: 2, repeatedUniformRunLength: 2 },
        1,
        13,
      ),
    ).toEqual({ kind: 'number', value: 4 })

    expect(
      withOptionalLookupStringIds({
        sheetName: 'Sheet1',
        row: 1,
        col: 2,
        oldValue: { tag: ValueTag.String, value: 'a' },
        newValue: { tag: ValueTag.String, value: 'b' },
        oldStringId: undefined,
        newStringId: 4,
        inputCellIndex: 9,
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      row: 1,
      col: 2,
      oldValue: { tag: ValueTag.String, value: 'a' },
      newValue: { tag: ValueTag.String, value: 'b' },
      newStringId: 4,
      inputCellIndex: 9,
    })
  })
})
