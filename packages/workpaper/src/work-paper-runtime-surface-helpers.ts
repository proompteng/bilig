import { CellFlags } from '@bilig/core/headless-runtime'
import type { CellRangeRef, CellValue } from '@bilig/protocol'
import { parseCellAddress } from '@bilig/formula'
import type { TrackedEngineEvent } from './tracked-engine-event-refs.js'
import { TINY_TRACKED_CHANGE_LIMIT } from './work-paper-tracked-event-helpers.js'
import type { WorkPaperNamedExpressionValueSnapshot } from './work-paper-named-expression-helpers.js'

export type NamedExpressionValueSnapshot = WorkPaperNamedExpressionValueSnapshot
export type RebuildValueSnapshot = Map<number, CellValue>

export const FORMULA_REBUILD_VALUE_FLAGS = CellFlags.HasFormula | CellFlags.SpillChild | CellFlags.PivotOutput
export const EMPTY_NAMED_EXPRESSION_VALUES: NamedExpressionValueSnapshot = new Map()

export function shouldPreferLazyPublicChanges(events: readonly TrackedEngineEvent[], shouldEmitValuesUpdated: boolean): boolean {
  if (!shouldEmitValuesUpdated) {
    return true
  }
  return events.some(
    (event) =>
      event.changedCellIndices.length > TINY_TRACKED_CHANGE_LIMIT &&
      event.invalidation !== 'full' &&
      event.patches === undefined &&
      !event.hasInvalidatedRanges &&
      !event.hasInvalidatedRows &&
      !event.hasInvalidatedColumns,
  )
}

export function cellAddressInRange(row: number, col: number, range: CellRangeRef): boolean {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  if (!start || !end) {
    return false
  }
  return row >= start.row && row <= end.row && col >= start.col && col <= end.col
}
