import { describe, expect, it } from 'vitest'
import type { EngineCellMutationRef } from '@bilig/core'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import {
  TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY,
  TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY,
  canSkipDimensionUpdateAfterLiteralMutation,
  countPotentialNewTrackedCellMutations,
  readTinySortedPhysicalTrackedEventChanges,
  readTrackedCellChange,
  readTrackedRuntimeCellValue,
  readTrustedPhysicalTrackedChangeMetadata,
  trackedEventHasNoValueChanges,
  trackedEventFromExistingNumericMutationResult,
  tryBuildDirectExistingNumericTrackedChanges,
  tryBuildDirectSingleLiteralTrackedChange,
  withEventChanges,
  type QueuedEvent,
  type WorkPaperTrackedAddressedCellStore,
  type WorkPaperTrackedWorkbookAccess,
} from '../work-paper-tracked-event-helpers.js'
import type { TrackedEngineEvent } from '../tracked-engine-event-refs.js'
import type { WorkPaperChange } from '../work-paper-types.js'

function trackedEvent(input: Partial<TrackedEngineEvent> = {}): TrackedEngineEvent {
  return {
    invalidation: 'cells',
    changedCellIndices: new Uint32Array(),
    changedInputCount: 0,
    changedCellIndicesSortedDisjoint: true,
    hasInvalidatedRanges: false,
    hasInvalidatedRows: false,
    hasInvalidatedColumns: false,
    ...input,
  }
}

function addressedCellStore(input: {
  readonly sheetIds?: readonly number[]
  readonly rows?: readonly number[]
  readonly cols?: readonly number[]
  readonly tags?: readonly ValueTag[]
  readonly numbers?: readonly number[]
}): WorkPaperTrackedAddressedCellStore {
  return {
    sheetIds: input.sheetIds ?? [],
    rows: input.rows ?? [],
    cols: input.cols ?? [],
    tags: input.tags ?? [],
    numbers: input.numbers ?? [],
    errors: [],
    getValue: () => ({ tag: ValueTag.Empty }),
  }
}

function trackedWorkbook(input: {
  readonly cellStore: WorkPaperTrackedAddressedCellStore
  readonly sheetName?: string
  readonly structureVersion?: number
  readonly logicalPositions?: ReadonlyMap<number, { readonly row: number; readonly col: number }>
}): WorkPaperTrackedWorkbookAccess {
  return {
    cellStore: input.cellStore,
    getCellPosition: (cellIndex) => input.logicalPositions?.get(cellIndex),
    getSheetById: (sheetId) => ({
      name: input.sheetName ?? `Sheet${sheetId}`,
      structureVersion: input.structureVersion ?? 1,
    }),
    getSheetNameById: (sheetId) => input.sheetName ?? `Sheet${sheetId}`,
  }
}

describe('work paper tracked event helpers', () => {
  it('reads trusted physical metadata from tracked changed-index arrays', () => {
    const changed = Uint32Array.of(1, 2)
    Reflect.set(changed, TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY, 4)
    Reflect.set(changed, TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY, 1)

    expect(readTrustedPhysicalTrackedChangeMetadata(changed)).toEqual({
      trustedPhysicalSheetId: 4,
      trustedSortedSliceSplit: 1,
    })
    expect(readTrustedPhysicalTrackedChangeMetadata(new Uint32Array())).toBeUndefined()
  })

  it('distinguishes value-changing tracked events from no-op invalidation noise', () => {
    expect(trackedEventHasNoValueChanges(trackedEvent())).toBe(true)
    expect(trackedEventHasNoValueChanges(trackedEvent({ changedCellIndices: Uint32Array.of(1) }))).toBe(false)
    expect(
      trackedEventHasNoValueChanges(
        trackedEvent({
          patches: [
            {
              kind: 'cell',
              cellIndex: 1,
              address: { sheet: 1, row: 0, col: 0 },
              sheetName: 'Sheet1',
              a1: 'A1',
              newValue: { tag: ValueTag.Empty },
            },
          ],
        }),
      ),
    ).toBe(false)
  })

  it('counts potential new tracked cell mutations and literal dimension-safe mutations', () => {
    const existingLiteral: EngineCellMutationRef = {
      sheetId: 1,
      cellIndex: 2,
      mutation: { kind: 'setCellValue', row: 0, col: 0, value: 10 },
    }
    const newLiteral: EngineCellMutationRef = {
      sheetId: 1,
      mutation: { kind: 'setCellValue', row: 1, col: 0, value: 20 },
    }
    const clearMissing: EngineCellMutationRef = {
      sheetId: 1,
      mutation: { kind: 'clearCell', row: 2, col: 0 },
    }

    expect(countPotentialNewTrackedCellMutations([existingLiteral, newLiteral, clearMissing])).toBe(1)
    expect(canSkipDimensionUpdateAfterLiteralMutation([existingLiteral], 0)).toBe(true)
    expect(canSkipDimensionUpdateAfterLiteralMutation([newLiteral], 0)).toBe(false)
    expect(canSkipDimensionUpdateAfterLiteralMutation([existingLiteral], 1)).toBe(false)
  })

  it('reads runtime cell values from compact cell-store arrays', () => {
    const cellStore = {
      tags: [ValueTag.Number, ValueTag.Boolean, ValueTag.String, ValueTag.Error, ValueTag.Empty],
      numbers: [12, 1, 0, 0, 0],
      errors: [ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.Value, ErrorCode.None],
      getValue: (_cellIndex: number, readString: (stringId: number) => string) => ({
        tag: ValueTag.String,
        value: readString(7),
      }),
    }
    const strings = { get: (id: number) => `s${String(id)}` }

    expect(readTrackedRuntimeCellValue(cellStore, 0, strings)).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(readTrackedRuntimeCellValue(cellStore, 1, strings)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(readTrackedRuntimeCellValue(cellStore, 2, strings)).toEqual({ tag: ValueTag.String, value: 's7' })
    expect(readTrackedRuntimeCellValue(cellStore, 3, strings)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(readTrackedRuntimeCellValue(cellStore, 4, strings)).toEqual({ tag: ValueTag.Empty })
  })

  it('converts compact existing-numeric mutation results into tracked events', () => {
    const event = trackedEventFromExistingNumericMutationResult({
      changedCellCount: 2,
      firstChangedCellIndex: 4,
      secondChangedCellIndex: 9,
      explicitChangedCount: 1,
      secondChangedNumericValue: 12,
      secondChangedRow: 3,
      secondChangedCol: 2,
    })

    expect(event).toMatchObject({
      invalidation: 'cells',
      changedInputCount: 1,
      changedCellIndicesSortedDisjoint: true,
      explicitChangedCount: 1,
      firstChangedCellIndex: 4,
      lastChangedCellIndex: 9,
      hasInvalidatedRanges: false,
      hasInvalidatedRows: false,
      hasInvalidatedColumns: false,
    })
    expect([...event.changedCellIndices]).toEqual([4, 9])
    expect(
      trackedEventFromExistingNumericMutationResult({
        changedCellIndices: Uint32Array.of(9, 4),
        explicitChangedCount: 1,
      }).changedCellIndicesSortedDisjoint,
    ).toBe(false)
  })

  it('reads visible tracked cell changes from physical and logical sheet records', () => {
    const cellStore = addressedCellStore({
      sheetIds: [1, 1],
      rows: [2, 99],
      cols: [3, 99],
      tags: [ValueTag.Number, ValueTag.Boolean],
      numbers: [7, 1],
    })

    expect(
      readTrackedCellChange({
        cellIndex: 0,
        workbook: trackedWorkbook({ cellStore, sheetName: 'Data' }),
        strings: { get: String },
        trackedA1: (row, col) => `R${row}C${col}`,
      }),
    ).toEqual({
      kind: 'cell',
      address: { sheet: 1, row: 2, col: 3 },
      sheetName: 'Data',
      a1: 'R2C3',
      newValue: { tag: ValueTag.Number, value: 7 },
    })

    expect(
      readTrackedCellChange({
        cellIndex: 1,
        workbook: trackedWorkbook({
          cellStore,
          sheetName: 'Logical',
          structureVersion: 2,
          logicalPositions: new Map([[1, { row: 4, col: 5 }]]),
        }),
        strings: { get: String },
        trackedA1: (row, col) => `R${row}C${col}`,
      }),
    ).toEqual({
      kind: 'cell',
      address: { sheet: 1, row: 4, col: 5 },
      sheetName: 'Logical',
      a1: 'R4C5',
      newValue: { tag: ValueTag.Boolean, value: true },
    })
  })

  it('reads tiny sorted physical tracked events without generic cell materialization', () => {
    const changes = readTinySortedPhysicalTrackedEventChanges({
      event: trackedEvent({ changedCellIndices: Uint32Array.of(1, 2) }),
      workbook: trackedWorkbook({
        sheetName: 'Sheet1',
        cellStore: addressedCellStore({
          sheetIds: [0, 1, 1],
          rows: [0, 1, 2],
          cols: [0, 0, 0],
          tags: [ValueTag.Empty, ValueTag.Number, ValueTag.Number],
          numbers: [0, 10, 20],
        }),
      }),
      strings: { get: String },
      trackedA1: (row, col) => `R${row}C${col}`,
    })

    expect(changes?.map((change) => [change.a1, change.newValue])).toEqual([
      ['R1C0', { tag: ValueTag.Number, value: 10 }],
      ['R2C0', { tag: ValueTag.Number, value: 20 }],
    ])
    expect(
      readTinySortedPhysicalTrackedEventChanges({
        event: trackedEvent({ changedCellIndices: Uint32Array.of(1, 2) }),
        workbook: trackedWorkbook({
          cellStore: addressedCellStore({
            sheetIds: [0, 1, 1],
            rows: [0, 2, 1],
            cols: [0, 0, 0],
          }),
        }),
        strings: { get: String },
        trackedA1: (row, col) => `R${row}C${col}`,
      }),
    ).toBeNull()
  })

  it('builds direct single-literal tracked changes without visibility snapshots', () => {
    const changes = tryBuildDirectSingleLiteralTrackedChange({
      events: [trackedEvent({ changedCellIndices: Uint32Array.of(4, 8) })],
      expected: {
        address: { sheet: 1, row: 2, col: 1 },
        cellIndex: 4,
        isPhysicalSheet: true,
        sheetName: 'Sheet1',
        value: 9,
      },
      cellStore: addressedCellStore({
        sheetIds: [0, 0, 0, 0, 1, 0, 0, 0, 1],
        rows: [0, 0, 0, 0, 2, 0, 0, 0, 3],
        cols: [0, 0, 0, 0, 1, 0, 0, 0, 1],
        tags: [
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Number,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Number,
        ],
        numbers: [0, 0, 0, 0, 9, 0, 0, 0, 18],
      }),
      strings: { get: String },
      trackedA1: (row, col) => `R${row}C${col}`,
    })

    expect(changes).toEqual([
      {
        kind: 'cell',
        address: { sheet: 1, row: 2, col: 1 },
        sheetName: 'Sheet1',
        a1: 'R2C1',
        newValue: { tag: ValueTag.Number, value: 9 },
      },
      {
        kind: 'cell',
        address: { sheet: 1, row: 3, col: 1 },
        sheetName: 'Sheet1',
        a1: 'R3C1',
        newValue: { tag: ValueTag.Number, value: 18 },
      },
    ])
  })

  it('builds direct single-literal fanout changes when dependents are already in sheet order', () => {
    const changes = tryBuildDirectSingleLiteralTrackedChange({
      events: [trackedEvent({ changedCellIndices: Uint32Array.of(4, 8, 9, 10) })],
      expected: {
        address: { sheet: 1, row: 2, col: 1 },
        cellIndex: 4,
        isPhysicalSheet: true,
        sheetName: 'Sheet1',
        value: 9,
      },
      cellStore: addressedCellStore({
        sheetIds: [0, 0, 0, 0, 1, 0, 0, 0, 1, 1, 1],
        rows: [0, 0, 0, 0, 2, 0, 0, 0, 3, 3, 4],
        cols: [0, 0, 0, 0, 1, 0, 0, 0, 1, 2, 0],
        tags: [
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Number,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Number,
          ValueTag.Number,
          ValueTag.Number,
        ],
        numbers: [0, 0, 0, 0, 9, 0, 0, 0, 18, 27, 36],
      }),
      strings: { get: String },
      trackedA1: (row, col) => `R${row}C${col}`,
    })

    expect(changes?.map((change) => (change.kind === 'cell' ? [change.a1, change.newValue] : []))).toEqual([
      ['R2C1', { tag: ValueTag.Number, value: 9 }],
      ['R3C1', { tag: ValueTag.Number, value: 18 }],
      ['R3C2', { tag: ValueTag.Number, value: 27 }],
      ['R4C0', { tag: ValueTag.Number, value: 36 }],
    ])
  })

  it('builds direct existing-numeric tracked changes and delegates ordering when needed', () => {
    const ordered = tryBuildDirectExistingNumericTrackedChanges({
      result: {
        changedCellIndices: Uint32Array.of(4, 10, 8),
        explicitChangedCount: 3,
      },
      address: { sheet: 1, row: 2, col: 1 },
      cellIndex: 4,
      isPhysicalSheet: true,
      sheetName: 'Sheet1',
      value: 9,
      cellStore: addressedCellStore({
        sheetIds: [0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 1],
        rows: [0, 0, 0, 0, 2, 0, 0, 0, 3, 0, 2],
        cols: [0, 0, 0, 0, 1, 0, 0, 0, 3, 0, 2],
        tags: [
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Number,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Empty,
          ValueTag.Number,
          ValueTag.Empty,
          ValueTag.Number,
        ],
        numbers: [0, 0, 0, 0, 9, 0, 0, 0, 80, 0, 100],
      }),
      strings: { get: String },
      trackedA1: (row, col) => `R${row}C${col}`,
      orderChanges: (changes, explicitChangedCount) => {
        expect(explicitChangedCount).toBe(3)
        return changes.toSorted((left, right) => left.address.row - right.address.row || left.address.col - right.address.col)
      },
    })

    expect(ordered?.map((change) => change.a1)).toEqual(['R2C1', 'R2C2', 'R3C3'])
  })

  it('attaches computed changes to queued semantic events that carry change payloads', () => {
    const changes: WorkPaperChange[] = [
      {
        kind: 'cell',
        address: { sheet: 1, row: 0, col: 0 },
        sheetName: 'Sheet1',
        a1: 'A1',
        newValue: { tag: ValueTag.Number, value: 1 },
      },
    ]
    const event: QueuedEvent = {
      eventName: 'namedExpressionAdded',
      payload: { name: 'Rate', changes: [] },
    }

    expect(withEventChanges(event, changes)).toEqual({
      eventName: 'namedExpressionAdded',
      payload: { name: 'Rate', changes },
    })
    const sheetAdded: QueuedEvent = {
      eventName: 'sheetAdded',
      payload: { sheetId: 1, sheetName: 'Sheet1' },
    }
    expect(withEventChanges(sheetAdded, changes)).toBe(sheetAdded)
  })
})
