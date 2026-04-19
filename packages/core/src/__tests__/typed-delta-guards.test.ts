import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { CELL_MUTATION_TRANSACTION_KIND, isCellMutationTransactionRecord } from '../history/typed-history.js'
import {
  ENGINE_CELL_PATCH_KIND,
  ENGINE_COLUMN_INVALIDATION_PATCH_KIND,
  ENGINE_RANGE_INVALIDATION_PATCH_KIND,
  ENGINE_ROW_INVALIDATION_PATCH_KIND,
  isEngineCellPatch,
  isEngineColumnInvalidationPatch,
  isEngineRangeInvalidationPatch,
  isEngineRowInvalidationPatch,
} from '../patches/patch-types.js'

describe('typed delta guards', () => {
  it('recognizes typed cell-mutation transaction records', () => {
    expect(
      isCellMutationTransactionRecord({
        kind: CELL_MUTATION_TRANSACTION_KIND,
        refs: [],
      }),
    ).toBe(true)
    expect(isCellMutationTransactionRecord({ kind: CELL_MUTATION_TRANSACTION_KIND, refs: null })).toBe(false)
    expect(isCellMutationTransactionRecord(null)).toBe(false)
  })

  it('recognizes tracked cell patches', () => {
    expect(
      isEngineCellPatch({
        kind: ENGINE_CELL_PATCH_KIND,
        cellIndex: 7,
        address: { sheet: 1, row: 2, col: 3 },
        sheetName: 'Sheet1',
        a1: 'D3',
        newValue: { tag: ValueTag.Number, value: 42 },
      }),
    ).toBe(true)
    expect(isEngineCellPatch({ kind: ENGINE_CELL_PATCH_KIND, cellIndex: '7' })).toBe(false)
    expect(isEngineCellPatch(undefined)).toBe(false)
  })

  it('recognizes tracked invalidation patches', () => {
    expect(
      isEngineRangeInvalidationPatch({
        kind: ENGINE_RANGE_INVALIDATION_PATCH_KIND,
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
      }),
    ).toBe(true)
    expect(
      isEngineRowInvalidationPatch({
        kind: ENGINE_ROW_INVALIDATION_PATCH_KIND,
        sheetName: 'Sheet1',
        startIndex: 1,
        endIndex: 2,
      }),
    ).toBe(true)
    expect(
      isEngineColumnInvalidationPatch({
        kind: ENGINE_COLUMN_INVALIDATION_PATCH_KIND,
        sheetName: 'Sheet1',
        startIndex: 3,
        endIndex: 4,
      }),
    ).toBe(true)
    expect(isEngineRangeInvalidationPatch({ kind: ENGINE_RANGE_INVALIDATION_PATCH_KIND, range: null })).toBe(false)
    expect(isEngineRowInvalidationPatch({ kind: ENGINE_ROW_INVALIDATION_PATCH_KIND, sheetName: 'Sheet1' })).toBe(false)
    expect(isEngineColumnInvalidationPatch(undefined)).toBe(false)
  })
})
