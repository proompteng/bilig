import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { makeCellKey } from '@bilig/core'
import { runProperty } from '@bilig/test-fuzz'
import { ValueTag, type CellValue } from '@bilig/protocol'
import type { TrackedEngineEvent } from '../tracked-engine-event-refs.js'
import { computeWorkPaperTrackedCellChangesFromEvents, type MaterializedTrackedEventChanges } from '../work-paper-tracked-change-reducer.js'
import type { WorkPaperCellChange } from '../work-paper-types.js'
import type { VisibilitySnapshot } from '../work-paper-visibility-snapshot.js'

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function trackedEvent(): TrackedEngineEvent {
  return {
    invalidation: 'cells',
    changedCellIndices: [1],
    changedInputCount: 1,
    changedCellIndicesSortedDisjoint: true,
    hasInvalidatedRanges: false,
    hasInvalidatedRows: false,
    hasInvalidatedColumns: false,
  }
}

function materialized(changes: WorkPaperCellChange[]): MaterializedTrackedEventChanges {
  return { changes, canReusePublicChanges: false, ordered: true }
}

describe('work paper tracked change reducer fuzz', () => {
  it('should not mutate the caller visibility snapshot while computing next visibility', async () => {
    await runProperty({
      suite: 'headless/work-paper-tracked-change-reducer/input-visibility-immutable',
      arbitrary: fc.record({
        row: fc.integer({ min: 0, max: 20 }),
        col: fc.integer({ min: 0, max: 12 }),
        beforeValue: fc.integer({ min: -1_000, max: 1_000 }),
        nextValue: fc.integer({ min: 1_001, max: 2_000 }),
      }),
      predicate: async ({ row, col, beforeValue, nextValue }) => {
        const beforeVisibility: VisibilitySnapshot = new Map([
          [
            1,
            {
              sheetId: 1,
              sheetName: 'Sheet1',
              order: 0,
              cells: new Map([[makeCellKey(1, row, col), numberValue(beforeValue)]]),
            },
          ],
        ])
        const beforeDigest = visibilityDigest(beforeVisibility)
        const change: WorkPaperCellChange = {
          kind: 'cell',
          address: { sheet: 1, row, col },
          sheetName: 'Sheet1',
          a1: `A${row + 1}`,
          newValue: numberValue(nextValue),
        }

        const result = computeWorkPaperTrackedCellChangesFromEvents({
          beforeVisibility,
          events: [trackedEvent()],
          listSheets: () => [{ id: 1, order: 0 }],
          materializeTrackedEventChanges: () => materialized([change]),
          readSingleTrackedCellChange: () => change,
          readTinySortedPhysicalTrackedEventChanges: () => null,
          sheetOrder: () => 0,
        })

        expect(result?.nextVisibility.get(1)?.cells.get(makeCellKey(1, row, col))).toEqual(numberValue(nextValue))
        expect(visibilityDigest(beforeVisibility)).toEqual(beforeDigest)
      },
      parameters: { numRuns: 120 },
    })
  })
})

function visibilityDigest(snapshot: VisibilitySnapshot): readonly string[] {
  const entries: string[] = []
  for (const [sheetId, sheet] of snapshot.entries()) {
    entries.push(`${sheetId}:sheet:${sheet.sheetName}:${sheet.order}`)
    for (const [cellKey, value] of sheet.cells.entries()) {
      entries.push(`${sheetId}:cell:${cellKey}:${JSON.stringify(value)}`)
    }
  }
  return entries.toSorted()
}
