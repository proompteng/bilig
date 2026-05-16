import {
  isCellNumberFormatInputValue,
  isCellRangeRef,
  isCellStyleFieldList,
  isCellStylePatchValue,
  isCommitOps,
  isLiteralInput,
  isPendingWorkbookMutationInput,
  isWorkbookSheetName,
  isWorkbookStructuralCount,
  isWorkbookStructuralIndex,
  isWorkbookStructuralSize,
  type PendingWorkbookMutation,
} from './workbook-sync.js'
import type { WorkerEngine } from './worker-runtime-support.js'

const MIN_COLUMN_WIDTH = 44
const MAX_COLUMN_WIDTH = 480

export function applyPendingWorkbookMutationToEngine(engine: WorkerEngine, mutation: PendingWorkbookMutation): void {
  if (!isPendingWorkbookMutationInput(mutation)) {
    return
  }
  const { method, args } = mutation
  switch (method) {
    case 'setCellValue': {
      const [sheetName, address, value] = args
      if (typeof sheetName === 'string' && typeof address === 'string' && isLiteralInput(value)) {
        engine.setCellValue(sheetName, address, value)
      }
      return
    }
    case 'setCellFormula': {
      const [sheetName, address, formula] = args
      if (typeof sheetName === 'string' && typeof address === 'string' && typeof formula === 'string') {
        engine.setCellFormula(sheetName, address, formula)
      }
      return
    }
    case 'clearCell': {
      const [sheetName, address] = args
      if (typeof sheetName === 'string' && typeof address === 'string') {
        engine.clearCell(sheetName, address)
      }
      return
    }
    case 'clearRange': {
      const [range] = args
      if (isCellRangeRef(range)) {
        engine.clearRange(range)
      }
      return
    }
    case 'renderCommit': {
      const [ops] = args
      if (isCommitOps(ops)) {
        engine.renderCommit(ops)
      }
      return
    }
    case 'fillRange':
    case 'copyRange':
    case 'moveRange': {
      const [source, target] = args
      if (isCellRangeRef(source) && isCellRangeRef(target)) {
        if (method === 'fillRange') {
          engine.fillRange(source, target)
        } else if (method === 'copyRange') {
          engine.copyRange(source, target)
        } else {
          engine.moveRange(source, target)
        }
      }
      return
    }
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns': {
      const [sheetName, start, count] = args
      if (isWorkbookSheetName(sheetName) && isWorkbookStructuralIndex(start) && isWorkbookStructuralCount(count)) {
        if (method === 'insertRows') {
          engine.insertRows(sheetName, start, count)
        } else if (method === 'deleteRows') {
          engine.deleteRows(sheetName, start, count)
        } else if (method === 'insertColumns') {
          engine.insertColumns(sheetName, start, count)
        } else {
          engine.deleteColumns(sheetName, start, count)
        }
      }
      return
    }
    case 'updateRowMetadata': {
      const [sheetName, startRow, count, height, hidden] = args
      if (
        isWorkbookSheetName(sheetName) &&
        isWorkbookStructuralIndex(startRow) &&
        isWorkbookStructuralCount(count) &&
        (height === null || isWorkbookStructuralSize(height)) &&
        (hidden === null || typeof hidden === 'boolean')
      ) {
        engine.updateRowMetadata(sheetName, startRow, count, height === null ? null : Math.max(1, Math.round(height)), hidden)
      }
      return
    }
    case 'updateColumnMetadata': {
      const [sheetName, startCol, count, width, hidden] = args
      if (
        isWorkbookSheetName(sheetName) &&
        isWorkbookStructuralIndex(startCol) &&
        isWorkbookStructuralCount(count) &&
        (width === null || isWorkbookStructuralSize(width)) &&
        (hidden === null || typeof hidden === 'boolean')
      ) {
        const clampedWidth = width === null ? null : Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(width)))
        engine.updateColumnMetadata(sheetName, startCol, count, clampedWidth, hidden)
      }
      return
    }
    case 'updateColumnWidth': {
      const [sheetName, columnIndex, width] = args
      if (isWorkbookSheetName(sheetName) && isWorkbookStructuralIndex(columnIndex) && isWorkbookStructuralSize(width)) {
        const clamped = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(width)))
        engine.updateColumnMetadata(sheetName, columnIndex, 1, clamped, null)
      }
      return
    }
    case 'setFreezePane': {
      const [sheetName, rows, cols] = args
      if (isWorkbookSheetName(sheetName) && isWorkbookStructuralIndex(rows) && isWorkbookStructuralIndex(cols)) {
        engine.setFreezePane(sheetName, Math.max(0, Math.round(rows)), Math.max(0, Math.round(cols)))
      }
      return
    }
    case 'mergeCells': {
      const [range] = args
      if (isCellRangeRef(range)) {
        engine.mergeCells(range)
      }
      return
    }
    case 'unmergeCells': {
      const [range] = args
      if (isCellRangeRef(range)) {
        engine.unmergeCells(range)
      }
      return
    }
    case 'setRangeStyle': {
      const [range, patch] = args
      if (isCellRangeRef(range) && isCellStylePatchValue(patch)) {
        engine.setRangeStyle(range, patch)
      }
      return
    }
    case 'clearRangeStyle': {
      const [range, fields] = args
      if (isCellRangeRef(range) && (fields === undefined || isCellStyleFieldList(fields))) {
        engine.clearRangeStyle(range, fields)
      }
      return
    }
    case 'setRangeNumberFormat': {
      const [range, format] = args
      if (isCellRangeRef(range) && isCellNumberFormatInputValue(format)) {
        engine.setRangeNumberFormat(range, format)
      }
      return
    }
    case 'clearRangeNumberFormat': {
      const [range] = args
      if (isCellRangeRef(range)) {
        engine.clearRangeNumberFormat(range)
      }
      return
    }
  }
}
