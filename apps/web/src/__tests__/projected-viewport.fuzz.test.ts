import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import type { CellSnapshot, RecalcMetrics } from '@bilig/protocol'
import { ValueTag } from '@bilig/protocol'
import type { ViewportPatch } from '@bilig/worker-transport'
import { runProperty } from '@bilig/test-fuzz'
import { ProjectedViewportStore } from '../projected-viewport-store.js'

type ViewportAction =
  | {
      kind: 'cell'
      row: number
      col: number
      version: number
      value: number | boolean | string | null
    }
  | {
      kind: 'column'
      index: number
      size: number
      hidden: boolean
    }
  | {
      kind: 'row'
      index: number
      size: number
      hidden: boolean
    }

const TEST_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

describe('projected viewport fuzz', () => {
  it('preserves patched cell and axis state for monotonic viewport updates', async () => {
    await runProperty({
      suite: 'web/projected-viewport/monotonic-patch-parity',
      arbitrary: fc.array(viewportActionArbitrary, { minLength: 4, maxLength: 24 }),
      predicate: async (actions) => {
        const store = new ProjectedViewportStore()
        const expectedCells = new Map<string, CellSnapshot>()
        const expectedColumnSizes = new Map<number, number>()
        const expectedHiddenColumns = new Map<number, true>()
        const expectedRowSizes = new Map<number, number>()
        const expectedHiddenRows = new Map<number, true>()

        actions.forEach((action) => {
          store.applyViewportPatch(patchFromAction(action))
          switch (action.kind) {
            case 'cell': {
              const snapshot = snapshotFromAction(action)
              const key = cellKey(snapshot.sheetName, snapshot.address)
              const previous = expectedCells.get(key)
              if (!previous || snapshot.version >= previous.version) {
                expectedCells.set(key, snapshot)
              }
              return
            }
            case 'column':
              expectedColumnSizes.set(action.index, action.size)
              if (action.hidden) {
                expectedHiddenColumns.set(action.index, true)
              } else {
                expectedHiddenColumns.delete(action.index)
              }
              return
            case 'row':
              expectedRowSizes.set(action.index, action.size)
              if (action.hidden) {
                expectedHiddenRows.set(action.index, true)
              } else {
                expectedHiddenRows.delete(action.index)
              }
              return
          }
        })

        expectedCells.forEach((snapshot) => {
          expect(store.getCell(snapshot.sheetName, snapshot.address)).toEqual(snapshot)
        })
        expect(store.getColumnSizes(sheetName)).toEqual(Object.fromEntries(expectedColumnSizes))
        expect(store.getHiddenColumns(sheetName)).toEqual(Object.fromEntries(expectedHiddenColumns))
        expect(store.getRowSizes(sheetName)).toEqual(Object.fromEntries(expectedRowSizes))
        expect(store.getHiddenRows(sheetName)).toEqual(Object.fromEntries(expectedHiddenRows))
      },
    })
  })
})

// Helpers

const sheetName = 'Sheet1'

const viewportActionArbitrary = fc.oneof<ViewportAction>(
  fc
    .record({
      row: fc.integer({ min: 0, max: 6 }),
      col: fc.integer({ min: 0, max: 4 }),
      version: fc.integer({ min: 1, max: 10 }),
      value: fc.oneof<number | boolean | string | null>(fc.integer({ min: -50, max: 50 }), fc.boolean(), fc.string(), fc.constant(null)),
    })
    .map((action) => Object.assign({ kind: 'cell' as const }, action)),
  fc
    .record({
      index: fc.integer({ min: 0, max: 4 }),
      size: fc.integer({ min: 60, max: 180 }),
      hidden: fc.boolean(),
    })
    .map((action) => Object.assign({ kind: 'column' as const }, action)),
  fc
    .record({
      index: fc.integer({ min: 0, max: 6 }),
      size: fc.integer({ min: 18, max: 60 }),
      hidden: fc.boolean(),
    })
    .map((action) => Object.assign({ kind: 'row' as const }, action)),
)

function patchFromAction(action: ViewportAction): ViewportPatch {
  return {
    version: 1,
    full: false,
    freezeRows: 0,
    freezeCols: 0,
    viewport: {
      sheetName,
      rowStart: 0,
      rowEnd: 8,
      colStart: 0,
      colEnd: 6,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells:
      action.kind === 'cell'
        ? [
            {
              row: action.row,
              col: action.col,
              snapshot: snapshotFromAction(action),
              displayText: `${action.value ?? ''}`,
              copyText: `${action.value ?? ''}`,
              editorText: `${action.value ?? ''}`,
              formatId: 0,
              styleId: 'style-0',
            },
          ]
        : [],
    columns: action.kind === 'column' ? [{ index: action.index, size: action.size, hidden: action.hidden }] : [],
    rows: action.kind === 'row' ? [{ index: action.index, size: action.size, hidden: action.hidden }] : [],
  }
}

function snapshotFromAction(action: Extract<ViewportAction, { kind: 'cell' }>): CellSnapshot {
  return {
    sheetName,
    address: addressFor(action.row, action.col),
    value: toCellValue(action.value),
    flags: 0,
    version: action.version,
  }
}

function toCellValue(value: number | boolean | string | null): CellSnapshot['value'] {
  if (value === null) {
    return { tag: ValueTag.Empty }
  }
  if (typeof value === 'number') {
    return { tag: ValueTag.Number, value }
  }
  if (typeof value === 'boolean') {
    return { tag: ValueTag.Boolean, value }
  }
  return { tag: ValueTag.String, value, stringId: 0 }
}

function addressFor(row: number, col: number): string {
  return `${String.fromCharCode(65 + col)}${row + 1}`
}

function cellKey(targetSheetName: string, address: string): string {
  return `${targetSheetName}!${address}`
}
