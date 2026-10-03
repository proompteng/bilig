import { describe, expect, it } from 'vitest'
import { makeCellKey } from '@bilig/core'
import { ValueTag, type CellValue } from '@bilig/protocol'
import type { TrackedEngineEvent } from '../tracked-engine-event-refs.js'
import {
  computeWorkPaperTrackedCellChangesFromEvents,
  tryReadTinyTrackedEventChangesWithoutVisibility,
  type MaterializedTrackedEventChanges,
  type WorkPaperTrackedChangeSheetRecord,
} from '../work-paper-tracked-change-reducer.js'
import type { WorkPaperCellChange } from '../work-paper-types.js'
import type { VisibilitySnapshot } from '../work-paper-visibility-snapshot.js'

const EMPTY_VALUE: CellValue = { tag: ValueTag.Empty }

const sheets: readonly WorkPaperTrackedChangeSheetRecord[] = [
  { id: 2, order: 0 },
  { id: 1, order: 1 },
]

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function change(sheet: number, row: number, col: number, value: CellValue = numberValue(row * 10 + col)): WorkPaperCellChange {
  return {
    kind: 'cell',
    address: { sheet, row, col },
    sheetName: sheet === 1 ? 'Sheet1' : 'Sheet2',
    a1: `${String.fromCharCode(65 + col)}${row + 1}`,
    newValue: value,
  }
}

function trackedEvent(input: Partial<TrackedEngineEvent> = {}): TrackedEngineEvent {
  return {
    invalidation: 'cells',
    changedCellIndices: [],
    changedInputCount: 1,
    changedCellIndicesSortedDisjoint: true,
    hasInvalidatedRanges: false,
    hasInvalidatedRows: false,
    hasInvalidatedColumns: false,
    ...input,
  }
}

function sheetOrder(sheetId: number): number {
  return sheets.find((sheet) => sheet.id === sheetId)?.order ?? 0
}

function materialized(changes: WorkPaperCellChange[], ordered = false): MaterializedTrackedEventChanges {
  return { changes, canReusePublicChanges: false, ordered }
}

describe('work-paper tracked change reducer', () => {
  it('orders tiny patch events by sheet order and address without visibility snapshots', () => {
    const changes = tryReadTinyTrackedEventChangesWithoutVisibility({
      event: trackedEvent({
        patches: [
          { ...change(1, 4, 0), cellIndex: 1 },
          { ...change(2, 0, 1), cellIndex: 2 },
        ],
      }),
      listSheets: () => sheets,
      materializeTrackedEventChanges: () => materialized([]),
      readSingleTrackedCellChange: () => undefined,
      readTinySortedPhysicalTrackedEventChanges: () => null,
      sheetOrder,
    })

    expect(changes?.map((entry) => (entry.kind === 'cell' ? `${entry.sheetName}!${entry.a1}` : ''))).toEqual(['Sheet2!B1', 'Sheet1!A5'])
  })

  it('rejects tiny index events with duplicate logical cells', () => {
    const changes = tryReadTinyTrackedEventChangesWithoutVisibility({
      event: trackedEvent({ changedCellIndices: [4, 8] }),
      listSheets: () => sheets,
      materializeTrackedEventChanges: () => materialized([]),
      readSingleTrackedCellChange: () => change(1, 0, 0),
      readTinySortedPhysicalTrackedEventChanges: () => null,
      sheetOrder,
    })

    expect(changes).toBeNull()
  })

  it('updates visibility for single small tracked events', () => {
    const beforeVisibility: VisibilitySnapshot = new Map([
      [
        1,
        {
          sheetId: 1,
          sheetName: 'Sheet1',
          order: 1,
          cells: new Map([[makeCellKey(1, 0, 0), numberValue(1)]]),
        },
      ],
    ])

    const result = computeWorkPaperTrackedCellChangesFromEvents({
      beforeVisibility,
      events: [trackedEvent({ changedCellIndices: [1, 2] })],
      listSheets: () => sheets,
      materializeTrackedEventChanges: () => materialized([]),
      readSingleTrackedCellChange: (cellIndex) => (cellIndex === 1 ? change(1, 0, 0, EMPTY_VALUE) : change(1, 0, 1, numberValue(7))),
      readTinySortedPhysicalTrackedEventChanges: () => null,
      sheetOrder,
    })

    expect(result?.changes.map((entry) => (entry.kind === 'cell' ? `${entry.sheetName}!${entry.a1}` : ''))).toEqual([
      'Sheet1!A1',
      'Sheet1!B1',
    ])
    expect(result?.nextVisibility.get(1)?.cells.has(makeCellKey(1, 0, 0))).toBe(false)
    expect(result?.nextVisibility.get(1)?.cells.get(makeCellKey(1, 0, 1))).toEqual(numberValue(7))
  })

  it('collapses repeated cells across tracked events to the latest change', () => {
    const eventA = trackedEvent({ changedCellIndices: [1] })
    const eventB = trackedEvent({ changedCellIndices: [2] })

    const result = computeWorkPaperTrackedCellChangesFromEvents({
      beforeVisibility: new Map(),
      events: [eventA, eventB],
      updateVisibility: false,
      listSheets: () => sheets,
      materializeTrackedEventChanges: (event) =>
        event === eventA ? materialized([change(1, 0, 0, numberValue(1))]) : materialized([change(1, 0, 0, numberValue(9))]),
      readSingleTrackedCellChange: () => undefined,
      readTinySortedPhysicalTrackedEventChanges: () => null,
      sheetOrder,
    })

    expect(result?.changes).toEqual([change(1, 0, 0, numberValue(9))])
  })
})
