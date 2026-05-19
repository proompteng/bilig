import type { SheetRecord } from '@bilig/core/headless-runtime'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { WORKPAPER_PUBLIC_ERROR_NAMES } from './work-paper-config.js'
import { WorkPaperOperationError } from './work-paper-errors.js'
import { normalizeAxisIntervals, type WorkPaperAxisIntervalEditMode, type WorkPaperAxisKind } from './work-paper-axis-helpers.js'
import { assertRowAndColumn } from './work-paper-runtime-helpers.js'
import type { WorkPaperAxisInterval, WorkPaperChange, WorkPaperConfig } from './work-paper-types.js'

export interface WorkPaperStructuralAxisFastPathRuntime {
  readonly applyAxisIntervalEditForSheet: (
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheet: SheetRecord,
    start: number,
    amount: number,
    options?: { readonly emitTracked?: boolean },
  ) => void
  readonly assertNotDisposed: () => void
  readonly batchStructuralChanges: (operations: () => void) => WorkPaperChange[]
  readonly canUseTrackedStructuralFastPath: () => boolean
  readonly captureTrackedChangesWithoutVisibilityCache: (mutate: () => void) => WorkPaperChange[]
  readonly getBatchUsesTrackedFastPath: () => boolean
  readonly getConfig: () => WorkPaperConfig
  readonly hasPendingLazyTrackedChanges: () => boolean
  readonly hasTrackedEngineEvents: () => boolean
  readonly materializePendingLazyTrackedChanges: () => void
  readonly drainTrackedEngineEvents: () => void
  readonly messageOf: (error: unknown, fallback: string) => string
  readonly sheetRecord: (sheetId: number) => SheetRecord
}

export function editWorkPaperAxisIntervalsWithoutRuntimeAdapters(
  runtime: WorkPaperStructuralAxisFastPathRuntime,
  axis: WorkPaperAxisKind,
  mode: WorkPaperAxisIntervalEditMode,
  sheetId: number,
  startOrInterval: number | WorkPaperAxisInterval,
  countOrInterval: number | WorkPaperAxisInterval | undefined,
  restIntervals: readonly WorkPaperAxisInterval[],
): WorkPaperChange[] {
  if (typeof startOrInterval === 'number' && (countOrInterval === undefined || typeof countOrInterval === 'number')) {
    return editSingleAxisIntervalWithoutRuntimeAdapters(runtime, axis, mode, sheetId, startOrInterval, countOrInterval ?? 1)
  }
  const indexes = normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals)
  if (!canEditAxisIntervalsWithoutRuntimeAdapters(runtime, axis, sheetId, indexes)) {
    throw new WorkPaperOperationError(`${axis === 'row' ? 'Rows' : 'Columns'} cannot be ${mode === 'add' ? 'added' : 'removed'}`)
  }
  const orderedIndexes = mode === 'remove' ? indexes.toSorted((left, right) => right[0] - left[0]) : indexes
  if (mode === 'add' && indexes.length === 1 && runtime.canUseTrackedStructuralFastPath()) {
    const [start, amount] = indexes[0]!
    return applyValuePreservingAxisInsert(runtime, axis, runtime.sheetRecord(sheetId), start, amount)
  }
  if (indexes.length === 1 && runtime.canUseTrackedStructuralFastPath()) {
    const [start, amount] = indexes[0]!
    return runtime.captureTrackedChangesWithoutVisibilityCache(() => {
      runtime.applyAxisIntervalEditForSheet(axis, mode, runtime.sheetRecord(sheetId), start, amount)
    })
  }
  if (indexes.length === 1 && runtime.getBatchUsesTrackedFastPath()) {
    const [start, amount] = indexes[0]!
    runtime.applyAxisIntervalEditForSheet(axis, mode, runtime.sheetRecord(sheetId), start, amount)
    return []
  }
  return runtime.batchStructuralChanges(() => {
    orderedIndexes.forEach(([start, amount]) => {
      runtime.applyAxisIntervalEditForSheet(axis, mode, runtime.sheetRecord(sheetId), start, amount)
    })
  })
}

function editSingleAxisIntervalWithoutRuntimeAdapters(
  runtime: WorkPaperStructuralAxisFastPathRuntime,
  axis: WorkPaperAxisKind,
  mode: WorkPaperAxisIntervalEditMode,
  sheetId: number,
  start: number,
  amount: number,
): WorkPaperChange[] {
  const sheet = sheetRecordForSingleAxisInterval(runtime, axis, mode, sheetId, start, amount)
  if (mode === 'add' && runtime.canUseTrackedStructuralFastPath()) {
    return applyValuePreservingAxisInsert(runtime, axis, sheet, start, amount)
  }
  if (runtime.canUseTrackedStructuralFastPath()) {
    return runtime.captureTrackedChangesWithoutVisibilityCache(() => {
      runtime.applyAxisIntervalEditForSheet(axis, mode, sheet, start, amount)
    })
  }
  if (runtime.getBatchUsesTrackedFastPath()) {
    runtime.applyAxisIntervalEditForSheet(axis, mode, sheet, start, amount)
    return []
  }
  return runtime.batchStructuralChanges(() => {
    runtime.applyAxisIntervalEditForSheet(axis, mode, sheet, start, amount)
  })
}

function sheetRecordForSingleAxisInterval(
  runtime: WorkPaperStructuralAxisFastPathRuntime,
  axis: WorkPaperAxisKind,
  mode: WorkPaperAxisIntervalEditMode,
  sheetId: number,
  start: number,
  amount: number,
): SheetRecord {
  runtime.assertNotDisposed()
  const sheet = runtime.sheetRecord(sheetId)
  const config = runtime.getConfig()
  const limit = axis === 'row' ? (config.maxRows ?? MAX_ROWS) : (config.maxColumns ?? MAX_COLS)
  assertRowAndColumn(start, 'start')
  assertRowAndColumn(amount, 'count')
  if (amount <= 0 || start + amount > limit) {
    throw new WorkPaperOperationError(`${axis === 'row' ? 'Rows' : 'Columns'} cannot be ${mode === 'add' ? 'added' : 'removed'}`)
  }
  return sheet
}

function applyValuePreservingAxisInsert(
  runtime: WorkPaperStructuralAxisFastPathRuntime,
  axis: WorkPaperAxisKind,
  sheet: SheetRecord,
  start: number,
  amount: number,
): WorkPaperChange[] {
  if (runtime.hasPendingLazyTrackedChanges()) {
    runtime.materializePendingLazyTrackedChanges()
  }
  if (runtime.hasTrackedEngineEvents()) {
    runtime.drainTrackedEngineEvents()
  }
  try {
    runtime.applyAxisIntervalEditForSheet(axis, 'add', sheet, start, amount, { emitTracked: false })
  } catch (error) {
    if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
      throw error
    }
    throw new WorkPaperOperationError(runtime.messageOf(error, 'Mutation failed'))
  }
  return []
}

function canEditAxisIntervalsWithoutRuntimeAdapters(
  runtime: WorkPaperStructuralAxisFastPathRuntime,
  axis: WorkPaperAxisKind,
  sheetId: number,
  indexes: readonly [number, number][],
): boolean {
  runtime.assertNotDisposed()
  void runtime.sheetRecord(sheetId)
  const config = runtime.getConfig()
  const limit = axis === 'row' ? (config.maxRows ?? MAX_ROWS) : (config.maxColumns ?? MAX_COLS)
  return indexes.every(([start, count]) => {
    assertRowAndColumn(start, 'start')
    assertRowAndColumn(count, 'count')
    return count > 0 && start + count <= limit
  })
}
