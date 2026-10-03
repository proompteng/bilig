import type { SpreadsheetEngine } from '@bilig/core/headless-runtime'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { WorkPaperCellChange } from './work-paper-types.js'

type TrackedCellStore = SpreadsheetEngine['workbook']['cellStore']

export interface DetachedPhysicalTrackedIndexChanges {
  readonly cellIndices?: Uint32Array
  readonly rows?: Uint32Array
  readonly cols?: Uint32Array
  readonly tags?: Uint8Array
  readonly constantTag?: ValueTag
  readonly numbers: Float64Array
  readonly errors?: Int32Array
  readonly stringIds?: Int32Array
  readonly strings?: readonly (string | undefined)[]
}

export function readTrackedCellValue(cellStore: TrackedCellStore, cellIndex: number, engine: SpreadsheetEngine): CellValue {
  const tag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
  switch (tag) {
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: cellStore.numbers[cellIndex] ?? 0 }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: (cellStore.numbers[cellIndex] ?? 0) !== 0 }
    case ValueTag.String:
      return cellStore.getValue(cellIndex, (stringId) => engine.strings.get(stringId))
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: cellStore.errors[cellIndex]! }
    case ValueTag.Empty:
    default:
      return { tag: ValueTag.Empty }
  }
}

export function readDetachedPhysicalTrackedIndexChange(
  index: number,
  sheetId: number,
  sheetName: string,
  cellStore: TrackedCellStore,
  detached: DetachedPhysicalTrackedIndexChanges,
  formatAddressCached: (row: number, col: number) => string,
): WorkPaperCellChange {
  const cellIndex = detached.cellIndices?.[index]
  const row = detached.rows?.[index] ?? (cellIndex === undefined ? 0 : (cellStore.rows[cellIndex] ?? 0))
  const col = detached.cols?.[index] ?? (cellIndex === undefined ? 0 : (cellStore.cols[cellIndex] ?? 0))
  return {
    kind: 'cell',
    address: { sheet: sheetId, row, col },
    sheetName,
    a1: formatAddressCached(row, col),
    newValue: readDetachedCellValue(detached, index),
  }
}

function readDetachedCellValue(detached: DetachedPhysicalTrackedIndexChanges, index: number): CellValue {
  const tag = (detached.tags?.[index] ?? detached.constantTag ?? ValueTag.Empty) as ValueTag
  switch (tag) {
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: detached.numbers[index] ?? 0 }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: (detached.numbers[index] ?? 0) !== 0 }
    case ValueTag.String:
      return { tag: ValueTag.String, value: detached.strings?.[index] ?? '', stringId: detached.stringIds?.[index] ?? 0 }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: detached.errors?.[index] ?? ErrorCode.None }
    case ValueTag.Empty:
    default:
      return { tag: ValueTag.Empty }
  }
}
