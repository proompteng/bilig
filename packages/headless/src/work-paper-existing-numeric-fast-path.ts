import {
  CellFlags,
  type EngineExistingLiteralCellMutationRef,
  type EngineExistingNumericCellMutationRef,
  type EngineExistingNumericCellMutationResult,
  type SheetRecord,
  type SpreadsheetEngine,
} from '@bilig/core'
import { ValueTag, type LiteralInput } from '@bilig/protocol'
import { WORKPAPER_PUBLIC_ERROR_NAMES } from './work-paper-config.js'
import { WorkPaperOperationError } from './work-paper-errors.js'
import {
  trackedEventFromExistingNumericMutationResult,
  tryBuildDirectExistingNumericTrackedChanges,
} from './work-paper-tracked-event-helpers.js'
import type { TrackedEngineEvent } from './tracked-engine-event-refs.js'
import type { WorkPaperCellAddress, WorkPaperCellChange, WorkPaperChange } from './work-paper-types.js'

const FAST_EXISTING_NUMERIC_LITERAL_FLAGS =
  CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput

type ExistingNumericMutationEngine = SpreadsheetEngine & {
  readonly tryApplyExistingNumericCellMutationAt?: (
    request: EngineExistingNumericCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyExistingLiteralCellMutationAt?: (
    request: EngineExistingLiteralCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
}

export interface WorkPaperExistingNumericFastPathRuntime {
  readonly canUseTrackedMutationFastPath: () => boolean
  readonly getEngine: () => SpreadsheetEngine
  readonly hasPendingLazyTrackedChanges: () => boolean
  readonly materializePendingLazyTrackedChanges: () => void
  readonly hasTrackedEngineEvents: () => boolean
  readonly drainTrackedEngineEvents: () => readonly TrackedEngineEvent[]
  readonly clearTrackedEngineEvents: () => void
  readonly getEngineEventCaptureEnabled: () => boolean
  readonly setEngineEventCaptureEnabled: (enabled: boolean) => void
  readonly hasPendingBatchOps: () => boolean
  readonly flushPendingBatchOps: () => void
  readonly messageOf: (error: unknown, fallback: string) => string
  readonly trackedA1: (row: number, col: number) => string
  readonly orderChanges: (changes: WorkPaperCellChange[], explicitChangedCount: number | undefined) => WorkPaperChange[]
  readonly computeTrackedChangesWithoutVisibilityCache: (
    events: readonly TrackedEngineEvent[],
    options: { readonly preferLazyPublicChanges?: boolean },
  ) => WorkPaperChange[]
  readonly hasValuesUpdatedListeners: () => boolean
  readonly emitValuesUpdated: (changes: WorkPaperChange[]) => void
}

export interface ExistingNumericWorkPaperCellContentsRequest {
  readonly sheet: SheetRecord
  readonly address: WorkPaperCellAddress
  readonly cellIndex: number
  readonly value: number
}

export interface ExistingLiteralWorkPaperCellContentsRequest {
  readonly sheet: SheetRecord
  readonly address: WorkPaperCellAddress
  readonly cellIndex: number
  readonly value: LiteralInput
}

export function trySetExistingNumericWorkPaperCellContentsWithTrackedFastPath(
  runtime: WorkPaperExistingNumericFastPathRuntime,
  request: ExistingNumericWorkPaperCellContentsRequest,
): WorkPaperChange[] | null {
  if (!runtime.canUseTrackedMutationFastPath() || request.sheet.structureVersion !== 1) {
    return null
  }
  const engine = runtime.getEngine()
  const cellStore = engine.workbook.cellStore
  if (
    cellStore.sheetIds[request.cellIndex] !== request.address.sheet ||
    cellStore.rows[request.cellIndex] !== request.address.row ||
    cellStore.cols[request.cellIndex] !== request.address.col ||
    (cellStore.formulaIds[request.cellIndex] ?? 0) !== 0 ||
    ((cellStore.flags[request.cellIndex] ?? 0) & FAST_EXISTING_NUMERIC_LITERAL_FLAGS) !== 0 ||
    cellStore.tags[request.cellIndex] !== ValueTag.Number
  ) {
    return null
  }
  const existingNumericMutationEngine: ExistingNumericMutationEngine = engine
  if (typeof existingNumericMutationEngine.tryApplyExistingNumericCellMutationAt !== 'function') {
    return null
  }

  if (runtime.hasPendingLazyTrackedChanges()) {
    runtime.materializePendingLazyTrackedChanges()
  }
  if (runtime.hasTrackedEngineEvents()) {
    runtime.drainTrackedEngineEvents()
  }
  let result: EngineExistingNumericCellMutationResult | null = null
  const oldNumericValue = cellStore.numbers[request.cellIndex] ?? 0
  const previousCaptureEnabled = runtime.getEngineEventCaptureEnabled()
  runtime.setEngineEventCaptureEnabled(false)
  try {
    if (runtime.hasPendingBatchOps()) {
      runtime.flushPendingBatchOps()
    }
    result = existingNumericMutationEngine.tryApplyExistingNumericCellMutationAt({
      sheetId: request.address.sheet,
      row: request.address.row,
      col: request.address.col,
      cellIndex: request.cellIndex,
      value: request.value,
      emitTracked: false,
      trustedExistingNumericLiteral: true,
      oldNumericValue,
    })
  } catch (error) {
    if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
      throw error
    }
    throw new WorkPaperOperationError(runtime.messageOf(error, 'Mutation failed'))
  } finally {
    runtime.setEngineEventCaptureEnabled(previousCaptureEnabled)
  }
  if (!result) {
    return null
  }

  if (runtime.hasTrackedEngineEvents()) {
    runtime.clearTrackedEngineEvents()
  }
  if (
    result.changedCellIndices === undefined &&
    result.changedCellCount === 2 &&
    result.firstChangedCellIndex === request.cellIndex &&
    result.secondChangedCellIndex !== undefined &&
    result.secondChangedNumericValue !== undefined &&
    result.secondChangedRow !== undefined &&
    result.secondChangedCol !== undefined
  ) {
    const changes: WorkPaperCellChange[] = [
      {
        kind: 'cell',
        address: { sheet: request.address.sheet, row: request.address.row, col: request.address.col },
        sheetName: request.sheet.name,
        a1: runtime.trackedA1(request.address.row, request.address.col),
        newValue: { tag: ValueTag.Number, value: request.value },
      },
      {
        kind: 'cell',
        address: { sheet: request.address.sheet, row: result.secondChangedRow, col: result.secondChangedCol },
        sheetName: request.sheet.name,
        a1: runtime.trackedA1(result.secondChangedRow, result.secondChangedCol),
        newValue: { tag: ValueTag.Number, value: result.secondChangedNumericValue },
      },
    ]
    if (runtime.hasValuesUpdatedListeners()) {
      runtime.emitValuesUpdated(changes)
    }
    return changes
  }

  let changes: WorkPaperChange[] | null = tryBuildDirectExistingNumericTrackedChanges({
    result,
    address: request.address,
    cellIndex: request.cellIndex,
    isPhysicalSheet: true,
    sheetName: request.sheet.name,
    value: request.value,
    cellStore,
    strings: engine.strings,
    trackedA1: (row, col) => runtime.trackedA1(row, col),
    orderChanges: (changesToOrder, explicitChangedCount) => runtime.orderChanges(changesToOrder, explicitChangedCount),
  })
  if (changes === null) {
    const shouldEmitValuesUpdated = runtime.hasValuesUpdatedListeners()
    const events = [trackedEventFromExistingNumericMutationResult(result)]
    changes = runtime.computeTrackedChangesWithoutVisibilityCache(events, {
      preferLazyPublicChanges: !shouldEmitValuesUpdated,
    })
    if (changes.length > 0 && shouldEmitValuesUpdated) {
      runtime.emitValuesUpdated(changes)
    }
    return changes
  }
  if (changes.length > 0 && runtime.hasValuesUpdatedListeners()) {
    runtime.emitValuesUpdated(changes)
  }
  return changes
}

export function trySetExistingLiteralWorkPaperCellContentsWithTrackedFastPath(
  runtime: WorkPaperExistingNumericFastPathRuntime,
  request: ExistingLiteralWorkPaperCellContentsRequest,
): WorkPaperChange[] | null {
  if (!runtime.canUseTrackedMutationFastPath() || request.sheet.structureVersion !== 1 || request.value === null) {
    return null
  }
  const engine = runtime.getEngine()
  const cellStore = engine.workbook.cellStore
  if (
    cellStore.sheetIds[request.cellIndex] !== request.address.sheet ||
    cellStore.rows[request.cellIndex] !== request.address.row ||
    cellStore.cols[request.cellIndex] !== request.address.col ||
    (cellStore.formulaIds[request.cellIndex] ?? 0) !== 0 ||
    ((cellStore.flags[request.cellIndex] ?? 0) & FAST_EXISTING_NUMERIC_LITERAL_FLAGS) !== 0
  ) {
    return null
  }
  const existingLiteralMutationEngine: ExistingNumericMutationEngine = engine
  if (typeof existingLiteralMutationEngine.tryApplyExistingLiteralCellMutationAt !== 'function') {
    return null
  }

  if (runtime.hasPendingLazyTrackedChanges()) {
    runtime.materializePendingLazyTrackedChanges()
  }
  if (runtime.hasTrackedEngineEvents()) {
    runtime.drainTrackedEngineEvents()
  }
  let result: EngineExistingNumericCellMutationResult | null = null
  const previousCaptureEnabled = runtime.getEngineEventCaptureEnabled()
  runtime.setEngineEventCaptureEnabled(false)
  try {
    if (runtime.hasPendingBatchOps()) {
      runtime.flushPendingBatchOps()
    }
    result = existingLiteralMutationEngine.tryApplyExistingLiteralCellMutationAt({
      sheetId: request.address.sheet,
      row: request.address.row,
      col: request.address.col,
      cellIndex: request.cellIndex,
      value: request.value,
      emitTracked: false,
    })
  } catch (error) {
    if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
      throw error
    }
    throw new WorkPaperOperationError(runtime.messageOf(error, 'Mutation failed'))
  } finally {
    runtime.setEngineEventCaptureEnabled(previousCaptureEnabled)
  }
  if (!result) {
    return null
  }

  if (runtime.hasTrackedEngineEvents()) {
    runtime.clearTrackedEngineEvents()
  }
  let changes: WorkPaperChange[] | null = tryBuildDirectExistingNumericTrackedChanges({
    result,
    address: request.address,
    cellIndex: request.cellIndex,
    isPhysicalSheet: true,
    sheetName: request.sheet.name,
    value: request.value,
    cellStore,
    strings: engine.strings,
    trackedA1: (row, col) => runtime.trackedA1(row, col),
    orderChanges: (changesToOrder, explicitChangedCount) => runtime.orderChanges(changesToOrder, explicitChangedCount),
  })
  if (changes === null) {
    const shouldEmitValuesUpdated = runtime.hasValuesUpdatedListeners()
    const events = [trackedEventFromExistingNumericMutationResult(result)]
    changes = runtime.computeTrackedChangesWithoutVisibilityCache(events, {
      preferLazyPublicChanges: !shouldEmitValuesUpdated,
    })
    if (changes.length > 0 && shouldEmitValuesUpdated) {
      runtime.emitValuesUpdated(changes)
    }
    return changes
  }
  if (changes.length > 0 && runtime.hasValuesUpdatedListeners()) {
    runtime.emitValuesUpdated(changes)
  }
  return changes
}
