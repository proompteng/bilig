import { AxisResidentCellIndex } from './storage/axis-resident-cell-index.js'
import { CellAxisIdentityStore } from './storage/cell-axis-identity-store.js'
import { CellPageStore } from './storage/cell-page-store.js'
import { LogicalSheetStore } from './storage/logical-sheet-store.js'
import { SheetAxisMap } from './storage/sheet-axis-map.js'
import { SheetGrid } from './sheet-grid.js'
import type { EngineCounters } from './perf/engine-counters.js'
import { makeLogicalCellKey } from './workbook-cell-key-index.js'
import type { WorkbookSheetFormatPrSnapshot } from '@bilig/protocol'
import type { WorkbookAxisEntryRecord, WorkbookFormatRangeRecord, WorkbookStyleRangeRecord } from './workbook-metadata-types.js'

export interface SheetRecord {
  id: number
  name: string
  order: number
  grid: SheetGrid
  axisMap: SheetAxisMap
  logicalAxisMap: SheetAxisMap
  logical: LogicalSheetStore
  cellIdentities: CellAxisIdentityStore
  residentCells: AxisResidentCellIndex
  columnVersions: Uint32Array
  structureVersion: number
  rowAxis: Array<WorkbookAxisEntryRecord | undefined>
  columnAxis: Array<WorkbookAxisEntryRecord | undefined>
  sheetFormatPr?: WorkbookSheetFormatPrSnapshot
  styleRanges: WorkbookStyleRangeRecord[]
  formatRanges: WorkbookFormatRangeRecord[]
}

export function createWorkbookSheetRecord(args: {
  readonly id: number
  readonly name: string
  readonly order: number
  readonly counters: EngineCounters | undefined
}): SheetRecord {
  const axisMap = new SheetAxisMap()
  const logicalAxisMap = new SheetAxisMap()
  const cellIdentities = new CellAxisIdentityStore()
  const residentCells = new AxisResidentCellIndex()
  const logical = new LogicalSheetStore(
    args.id,
    logicalAxisMap,
    new CellPageStore(
      new Map<string, number>(),
      (location) => makeLogicalCellKey(location.sheetId, location.rowId, location.colId),
      undefined,
      makeLogicalCellKey,
    ),
    cellIdentities,
    residentCells,
  )

  return {
    id: args.id,
    name: args.name,
    order: args.order,
    grid: new SheetGrid(args.counters, {
      get: (row, col) => logical.getVisibleCell(row, col),
      forEachCellEntry: (fn) => {
        logical.forEachVisibleCellEntry(fn)
      },
      someCellInAxisScope: (axis, scope, predicate) => logical.someResidentCellInAxisScope(axis, scope, predicate),
    }),
    axisMap,
    logicalAxisMap,
    logical,
    cellIdentities,
    residentCells,
    columnVersions: new Uint32Array(0),
    structureVersion: 1,
    rowAxis: [],
    columnAxis: [],
    styleRanges: [],
    formatRanges: [],
  }
}
