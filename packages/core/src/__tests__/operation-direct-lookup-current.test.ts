import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { PreparedApproximateVectorLookup, PreparedExactVectorLookup, RuntimeDirectLookupDescriptor } from '../engine/runtime-state.js'
import { createOperationDirectLookupCurrentService } from '../engine/services/operation-direct-lookup-current.js'

function exactUniform(
  overrides: Partial<Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>> = {},
): Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }> {
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
  }
}

function approximateUniform(
  overrides: Partial<Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>> = {},
): Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }> {
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
  }
}

function preparedExactLookup(overrides: Partial<PreparedExactVectorLookup> = {}): PreparedExactVectorLookup {
  return {
    sheetName: 'Sheet1',
    rowStart: 0,
    rowEnd: 3,
    col: 0,
    length: 4,
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
    ...overrides,
  }
}

function preparedApproximateLookup(overrides: Partial<PreparedApproximateVectorLookup> = {}): PreparedApproximateVectorLookup {
  return {
    sheetName: 'Sheet1',
    rowStart: 0,
    rowEnd: 3,
    col: 0,
    length: 4,
    columnVersion: 1,
    structureVersion: 1,
    sheetColumnVersions: new Uint32Array([1]),
    comparableKind: 'numeric',
    uniformStart: undefined,
    uniformStep: undefined,
    repeatedUniformStart: undefined,
    repeatedUniformStep: undefined,
    repeatedUniformRunLength: undefined,
    sortedAscending: true,
    sortedDescending: false,
    numericValues: Float64Array.of(1, 2, 4, 8),
    textValues: undefined,
    ...overrides,
  }
}

function createService(
  request: {
    directLookup?: RuntimeDirectLookupDescriptor
    tags?: ArrayLike<ValueTag | undefined>
    numbers?: ArrayLike<number | undefined>
    sheetVersion?: { structureVersion: number; columnVersion: number }
    exactPosition?: number | undefined
    sortedPosition?: number | undefined
  } = {},
) {
  const directLookup = request.directLookup
  const exactLookup = {
    findPreparedVectorMatch: vi.fn(() => ({ handled: true as const, position: request.exactPosition })),
  }
  const sortedLookup = {
    findPreparedVectorMatch: vi.fn(() => ({ handled: true as const, position: request.sortedPosition })),
  }
  const service = createOperationDirectLookupCurrentService({
    state: {
      workbook: {
        cellStore: {
          tags: request.tags ?? [],
          numbers: request.numbers ?? [],
        },
        getSheetById: (sheetId: number) =>
          sheetId === 7
            ? {
                id: 7,
                structureVersion: request.sheetVersion?.structureVersion ?? 3,
                columnVersions: new Uint32Array([request.sheetVersion?.columnVersion ?? 2]),
              }
            : undefined,
      },
      formulas: {
        get: (cellIndex: number) => (cellIndex === 10 && directLookup !== undefined ? { directLookup } : undefined),
      },
    },
    exactLookup,
    sortedLookup,
  })
  return { service, exactLookup, sortedLookup }
}

describe('operation direct lookup current results', () => {
  it('evaluates uniform exact and approximate lookup descriptors against current cell values', () => {
    expect(
      createService({
        directLookup: exactUniform(),
        tags: [undefined, ValueTag.Number],
        numbers: [undefined, 3],
      }).service.tryDirectUniformLookupCurrentResult(10),
    ).toEqual({ kind: 'number', value: 3 })

    expect(
      createService({
        directLookup: exactUniform(),
        tags: [undefined, ValueTag.String],
      }).service.tryDirectUniformLookupCurrentResult(10),
    ).toEqual({ kind: 'error', code: ErrorCode.NA })

    expect(
      createService({
        directLookup: exactUniform(),
        tags: [undefined, ValueTag.Number],
        numbers: [undefined, 3],
        sheetVersion: { structureVersion: 3, columnVersion: 99 },
      }).service.tryDirectUniformLookupCurrentResult(10),
    ).toBeUndefined()

    expect(
      createService({
        directLookup: approximateUniform(),
        tags: [undefined, ValueTag.Boolean],
        numbers: [undefined, 1],
      }).service.tryDirectUniformLookupCurrentResult(10),
    ).toEqual({ kind: 'number', value: 1 })
  })

  it('evaluates uniform lookup results from prepared numeric operands and sheet hints', () => {
    const matchingHint = { id: 7, structureVersion: 3, columnVersions: new Uint32Array([2]) }
    const staleHint = { id: 7, structureVersion: 3, columnVersions: new Uint32Array([9]) }
    const exactService = createService({ directLookup: exactUniform() }).service
    const approximateService = createService({ directLookup: approximateUniform() }).service

    expect(exactService.tryDirectUniformLookupCurrentResultFromNumeric(10, 4, undefined, matchingHint)).toEqual({
      kind: 'number',
      value: 4,
    })
    expect(exactService.tryDirectUniformLookupNumericResultFromDescriptor(exactUniform(), 4, undefined, matchingHint)).toBe(4)
    expect(exactService.canEvaluateDirectUniformLookupCurrentResultFromNumeric(10, 4, undefined)).toBe(true)
    expect(exactService.tryDirectUniformLookupCurrentResultFromNumeric(10, 4, undefined, staleHint)).toBeUndefined()

    expect(approximateService.tryDirectUniformLookupCurrentResultFromNumeric(10, undefined, 3.5, matchingHint)).toEqual({
      kind: 'number',
      value: 3,
    })
    expect(approximateService.tryDirectUniformLookupNumericResultFromDescriptor(approximateUniform(), undefined, 3.5, matchingHint)).toBe(3)
    expect(approximateService.canEvaluateDirectUniformLookupCurrentResultFromNumeric(10, undefined, 3.5)).toBe(true)

    const repeatedApproximate = approximateUniform({
      length: 6,
      rowEnd: 5,
      start: 1,
      step: 1,
      repeatedRunLength: 2,
    })
    const repeatedApproximateService = createService({ directLookup: repeatedApproximate }).service
    expect(repeatedApproximateService.tryDirectUniformLookupCurrentResultFromNumeric(10, undefined, 2.5, matchingHint)).toEqual({
      kind: 'number',
      value: 4,
    })
    expect(
      repeatedApproximateService.tryDirectUniformLookupNumericResultFromDescriptor(repeatedApproximate, undefined, 2.5, matchingHint),
    ).toBe(4)
    expect(repeatedApproximateService.tryDirectUniformLookupCurrentResultFromNumeric(10, undefined, 0, matchingHint)).toEqual({
      kind: 'error',
      code: ErrorCode.NA,
    })
  })

  it('evaluates non-uniform exact and approximate prepared lookups', () => {
    const approximateService = createService().service
    const approximateLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate' }> = {
      kind: 'approximate',
      operandCellIndex: 1,
      prepared: preparedApproximateLookup(),
      matchMode: 1,
    }
    expect(approximateService.tryDirectApproximateLookupCurrentResultFromNumeric(approximateLookup, 5)).toEqual({
      kind: 'number',
      value: 3,
    })

    const { service, exactLookup, sortedLookup } = createService({
      exactPosition: 2,
      sortedPosition: 4,
    })
    const exactDirectLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact' }> = {
      kind: 'exact',
      operandCellIndex: 1,
      prepared: preparedExactLookup(),
      searchMode: 1,
    }
    expect(service.tryDirectExactLookupCurrentResult(exactDirectLookup, { tag: ValueTag.String, value: 'needle' })).toEqual({
      kind: 'number',
      value: 2,
    })
    expect(exactLookup.findPreparedVectorMatch).toHaveBeenCalledTimes(1)

    const fallbackApproximateLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate' }> = {
      kind: 'approximate',
      operandCellIndex: 1,
      prepared: preparedApproximateLookup({ numericValues: undefined, repeatedUniformStart: undefined }),
      matchMode: 1,
    }
    expect(service.tryDirectApproximateLookupCurrentResultFromNumeric(fallbackApproximateLookup, 5)).toEqual({
      kind: 'number',
      value: 4,
    })
    expect(sortedLookup.findPreparedVectorMatch).toHaveBeenCalledTimes(1)
  })
})
