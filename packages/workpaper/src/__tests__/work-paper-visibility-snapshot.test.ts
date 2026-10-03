import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'
import {
  captureWorkPaperVisibilitySnapshot,
  computeWorkPaperCellChangesFromVisibilitySnapshots,
  type WorkPaperVisibilitySheetRecord,
} from '../work-paper-visibility-snapshot.js'

const empty = (): CellValue => ({ tag: ValueTag.Empty })
const number = (value: number): CellValue => ({ tag: ValueTag.Number, value })

function sheet(input: {
  readonly id: number
  readonly name: string
  readonly order: number
  readonly entries: readonly { readonly cellIndex: number; readonly row: number; readonly col: number }[]
}): WorkPaperVisibilitySheetRecord {
  return {
    id: input.id,
    name: input.name,
    order: input.order,
    grid: {
      forEachCellEntry: (visit) => {
        input.entries.forEach((entry) => visit(entry.cellIndex, entry.row, entry.col))
      },
    },
  }
}

function snapshot(sheets: readonly WorkPaperVisibilitySheetRecord[], values: ReadonlyMap<number, CellValue>) {
  return captureWorkPaperVisibilitySnapshot({
    sheets,
    cellStore: {
      getValue: (cellIndex) => values.get(cellIndex) ?? empty(),
    },
    strings: { get: String },
  })
}

describe('work paper visibility snapshots', () => {
  it('captures only visible non-empty cells', () => {
    const captured = snapshot(
      [
        sheet({
          id: 1,
          name: 'Sheet1',
          order: 0,
          entries: [
            { cellIndex: 1, row: 0, col: 0 },
            { cellIndex: 2, row: 1, col: 0 },
          ],
        }),
      ],
      new Map([
        [1, number(10)],
        [2, empty()],
      ]),
    )

    expect(captured.get(1)?.cells.size).toBe(1)
    expect([...captured.get(1)!.cells.values()]).toEqual([number(10)])
  })

  it('computes ordered cell changes between visibility snapshots', () => {
    const sheets = [
      sheet({ id: 1, name: 'Later', order: 1, entries: [{ cellIndex: 1, row: 0, col: 0 }] }),
      sheet({ id: 2, name: 'Earlier', order: 0, entries: [{ cellIndex: 2, row: 1, col: 1 }] }),
    ]
    const before = snapshot(
      sheets,
      new Map([
        [1, number(1)],
        [2, number(2)],
      ]),
    )
    const after = snapshot(
      sheets,
      new Map([
        [1, number(3)],
        [2, number(4)],
      ]),
    )

    expect(
      computeWorkPaperCellChangesFromVisibilitySnapshots({
        beforeVisibility: before,
        afterVisibility: after,
        sheets,
      }),
    ).toEqual([
      {
        kind: 'cell',
        address: { sheet: 2, row: 1, col: 1 },
        sheetName: 'Earlier',
        a1: 'B2',
        newValue: number(4),
      },
      {
        kind: 'cell',
        address: { sheet: 1, row: 0, col: 0 },
        sheetName: 'Later',
        a1: 'A1',
        newValue: number(3),
      },
    ])
  })
})
