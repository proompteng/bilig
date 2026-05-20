import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import {
  createWrittenColumnTracker,
  markWrittenColumn,
  materializeWrittenColumns,
  type WrittenColumnTracker,
} from '../../written-column-tracker.js'
import type { EngineRuntimeState } from '../runtime-state.js'

export interface InitialFormulaValueWriter {
  readonly writeValue: (cellIndex: number, value: CellValue) => void
  readonly writeValueAt: (cellIndex: number, sheetId: number, col: number, value: CellValue) => void
  readonly writeNumber: (cellIndex: number, value: number) => void
  readonly writeNumberAt: (cellIndex: number, sheetId: number, col: number, value: number) => void
  readonly flush: () => void
}

export function createInitialFormulaValueWriter(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'>
}): InitialFormulaValueWriter {
  let singleSheetId: number | undefined
  let singleSheetTracker: WrittenColumnTracker | undefined
  let writtenColumnsBySheetId: Map<number, WrittenColumnTracker> | undefined
  const promoteSingleSheetTracker = (): Map<number, WrittenColumnTracker> => {
    writtenColumnsBySheetId = new Map()
    if (singleSheetId !== undefined && singleSheetTracker !== undefined) {
      writtenColumnsBySheetId.set(singleSheetId, singleSheetTracker)
    }
    singleSheetId = undefined
    singleSheetTracker = undefined
    return writtenColumnsBySheetId
  }
  const markKnownColumn = (sheetId: number, col: number): void => {
    if (!writtenColumnsBySheetId && (singleSheetId === undefined || singleSheetId === sheetId)) {
      singleSheetId = sheetId
      singleSheetTracker ??= createWrittenColumnTracker()
      markWrittenColumn(singleSheetTracker, col)
      return
    }
    const trackers = writtenColumnsBySheetId ?? promoteSingleSheetTracker()
    let tracker = trackers.get(sheetId)
    if (!tracker) {
      tracker = createWrittenColumnTracker()
      trackers.set(sheetId, tracker)
    }
    markWrittenColumn(tracker, col)
  }
  const markCellColumn = (cellIndex: number): void => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const col = args.state.workbook.cellStore.cols[cellIndex]
    if (sheetId === undefined || col === undefined) {
      return
    }
    markKnownColumn(sheetId, col)
  }
  const clearDerivedFlags = (cellIndex: number): void => {
    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  }
  const writeNumberCore = (cellIndex: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    clearDerivedFlags(cellIndex)
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = 0
    cellStore.numbers[cellIndex] = value
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  }
  const writeValueCore = (cellIndex: number, value: CellValue): void => {
    const cellStore = args.state.workbook.cellStore
    clearDerivedFlags(cellIndex)
    cellStore.tags[cellIndex] = value.tag
    cellStore.errors[cellIndex] = value.tag === ValueTag.Error ? value.code : ErrorCode.None
    cellStore.stringIds[cellIndex] = value.tag === ValueTag.String ? args.state.strings.intern(value.value) : 0
    cellStore.numbers[cellIndex] = value.tag === ValueTag.Number ? value.value : value.tag === ValueTag.Boolean ? (value.value ? 1 : 0) : 0
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  }
  return {
    writeValue(cellIndex, value) {
      writeValueCore(cellIndex, value)
      markCellColumn(cellIndex)
    },
    writeValueAt(cellIndex, sheetId, col, value) {
      writeValueCore(cellIndex, value)
      markKnownColumn(sheetId, col)
    },
    writeNumber(cellIndex, value) {
      writeNumberCore(cellIndex, value)
      markCellColumn(cellIndex)
    },
    writeNumberAt(cellIndex, sheetId, col, value) {
      writeNumberCore(cellIndex, value)
      markKnownColumn(sheetId, col)
    },
    flush() {
      if (!writtenColumnsBySheetId) {
        if (singleSheetId !== undefined && singleSheetTracker !== undefined && singleSheetTracker.count > 0) {
          args.state.workbook.notifyColumnsWritten(singleSheetId, materializeWrittenColumns(singleSheetTracker))
        }
        return
      }
      writtenColumnsBySheetId.forEach((tracker, sheetId) => {
        if (tracker.count > 0) {
          args.state.workbook.notifyColumnsWritten(sheetId, materializeWrittenColumns(tracker))
        }
      })
    },
  }
}
