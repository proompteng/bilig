import {
  CellFlags,
  type EngineExistingLiteralCellMutationRef,
  type EngineExistingNumericCellMutationRef,
  type EngineExistingNumericCellMutationResult,
  type SheetRecord,
  type SpreadsheetEngine,
} from '@bilig/core/headless-runtime'
import { ValueTag, type LiteralInput } from '@bilig/protocol'
import { WORKPAPER_PUBLIC_ERROR_NAMES } from './work-paper-config.js'
import { WorkPaperOperationError } from './work-paper-errors.js'
import {
  trackedEventFromExistingNumericMutationResult,
  tryBuildDirectExistingNumericTrackedChanges,
} from './work-paper-tracked-event-helpers.js'
import {
  createLazyPhysicalTrackedIndexChanges,
  createPrefixedLazyTrackedIndexChanges,
  formatTrackedAddress,
  tryCreateLazyPhysicalTrackedIndexChanges,
} from './tracked-cell-lazy-physical-changes.js'
import type { TrackedEngineEvent } from './tracked-engine-event-refs.js'
import { scalarValueFromLiteral } from './work-paper-runtime-helpers.js'
import type { WorkPaperCellAddress, WorkPaperCellChange, WorkPaperChange } from './work-paper-types.js'

const FAST_EXISTING_NUMERIC_LITERAL_FLAGS =
  CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput
const LAZY_DIRECT_EXISTING_NUMERIC_CHANGE_THRESHOLD = 32

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
  readonly trackLazyChanges: (changes: WorkPaperCellChange[]) => void
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
  const compactLazyChanges = tryBuildCompactLazyDirectExistingNumericTrackedChanges(runtime, engine, request, result)
  if (compactLazyChanges !== null) {
    if (runtime.hasValuesUpdatedListeners()) {
      runtime.emitValuesUpdated(compactLazyChanges)
    }
    return compactLazyChanges
  }
  if (result.changedCellIndices === undefined && result.changedCellCount === 1 && result.firstChangedCellIndex === request.cellIndex) {
    const changes: WorkPaperCellChange[] = [
      {
        kind: 'cell',
        address: { sheet: request.address.sheet, row: request.address.row, col: request.address.col },
        sheetName: request.sheet.name,
        a1: runtime.trackedA1(request.address.row, request.address.col),
        newValue: { tag: ValueTag.Number, value: request.value },
      },
    ]
    if (runtime.hasValuesUpdatedListeners()) {
      runtime.emitValuesUpdated(changes)
    }
    return changes
  }
  const lazyChanges = tryBuildLazyDirectExistingNumericTrackedChanges(
    runtime,
    engine,
    {
      address: request.address,
      cellIndex: request.cellIndex,
      sheet: request.sheet,
      value: request.value,
    },
    result,
  )
  if (lazyChanges !== null) {
    if (runtime.hasValuesUpdatedListeners()) {
      runtime.emitValuesUpdated(lazyChanges)
    }
    return lazyChanges
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
      preferLazyPublicChanges:
        !shouldEmitValuesUpdated ||
        events.some((event) => event.changedCellIndices.length >= LAZY_DIRECT_EXISTING_NUMERIC_CHANGE_THRESHOLD),
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

function tryBuildCompactLazyDirectExistingNumericTrackedChanges(
  runtime: WorkPaperExistingNumericFastPathRuntime,
  engine: SpreadsheetEngine,
  request: ExistingNumericWorkPaperCellContentsRequest,
  result: EngineExistingNumericCellMutationResult,
): WorkPaperCellChange[] | null {
  if (runtime.hasValuesUpdatedListeners() || result.changedCellIndices !== undefined) {
    return null
  }
  const changedCellCount = result.changedCellCount
  if (changedCellCount === undefined || changedCellCount < LAZY_DIRECT_EXISTING_NUMERIC_CHANGE_THRESHOLD) {
    return null
  }
  let changedCellIndices: Uint32Array | undefined
  if (changedCellCount === 1 && result.firstChangedCellIndex === request.cellIndex) {
    changedCellIndices = Uint32Array.of(request.cellIndex)
  } else if (
    changedCellCount === 2 &&
    result.firstChangedCellIndex === request.cellIndex &&
    result.secondChangedCellIndex !== undefined &&
    result.secondChangedRow !== undefined &&
    result.secondChangedCol !== undefined
  ) {
    changedCellIndices = Uint32Array.of(request.cellIndex, result.secondChangedCellIndex)
  }
  if (changedCellIndices === undefined) {
    return null
  }
  const cellStore = engine.workbook.cellStore
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    if (cellStore.sheetIds[changedCellIndices[index]!] !== request.sheet.id) {
      return null
    }
  }
  const changes = createLazyPhysicalTrackedIndexChanges(
    request.sheet.id,
    request.sheet.name,
    cellStore,
    engine,
    changedCellIndices,
    formatTrackedAddress,
  )
  runtime.trackLazyChanges(changes)
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
  const lazyChanges = tryBuildLazyDirectExistingNumericTrackedChanges(
    runtime,
    engine,
    {
      address: request.address,
      cellIndex: request.cellIndex,
      sheet: request.sheet,
      value: request.value,
    },
    result,
  )
  if (lazyChanges !== null) {
    if (runtime.hasValuesUpdatedListeners()) {
      runtime.emitValuesUpdated(lazyChanges)
    }
    return lazyChanges
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
      preferLazyPublicChanges:
        !shouldEmitValuesUpdated ||
        events.some((event) => event.changedCellIndices.length >= LAZY_DIRECT_EXISTING_NUMERIC_CHANGE_THRESHOLD),
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

function tryBuildLazyDirectExistingNumericTrackedChanges(
  runtime: WorkPaperExistingNumericFastPathRuntime,
  engine: SpreadsheetEngine,
  request: {
    readonly address: WorkPaperCellAddress
    readonly cellIndex: number
    readonly sheet: SheetRecord
    readonly value: LiteralInput
  },
  result: EngineExistingNumericCellMutationResult,
): WorkPaperCellChange[] | null {
  const changedCellIndices = result.changedCellIndices
  if (
    changedCellIndices === undefined ||
    changedCellIndices.length < LAZY_DIRECT_EXISTING_NUMERIC_CHANGE_THRESHOLD ||
    changedCellIndices[0] !== request.cellIndex
  ) {
    return null
  }
  const cellStore = engine.workbook.cellStore
  let previousRow = -1
  let previousCol = -1
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    if (cellStore.sheetIds[cellIndex] !== request.sheet.id) {
      return tryBuildPrefixedCrossSheetLazyChanges(runtime, engine, request, changedCellIndices)
    }
    const row = cellStore.rows[cellIndex]
    const col = cellStore.cols[cellIndex]
    if (row === undefined || col === undefined || row < previousRow || (row === previousRow && col < previousCol)) {
      return null
    }
    previousRow = row
    previousCol = col
  }
  const changes = createLazyPhysicalTrackedIndexChanges(
    request.sheet.id,
    request.sheet.name,
    cellStore,
    engine,
    changedCellIndices,
    formatTrackedAddress,
  )
  runtime.trackLazyChanges(changes)
  return changes
}

function tryBuildPrefixedCrossSheetLazyChanges(
  runtime: WorkPaperExistingNumericFastPathRuntime,
  engine: SpreadsheetEngine,
  request: {
    readonly address: WorkPaperCellAddress
    readonly cellIndex: number
    readonly sheet: SheetRecord
    readonly value: LiteralInput
  },
  changedCellIndices: Uint32Array,
): WorkPaperCellChange[] | null {
  const tailStart = 1
  const firstTailCellIndex = changedCellIndices[tailStart]
  if (firstTailCellIndex === undefined) {
    return null
  }
  const cellStore = engine.workbook.cellStore
  const tailSheetId = cellStore.sheetIds[firstTailCellIndex]
  if (tailSheetId === undefined || tailSheetId === request.sheet.id) {
    return null
  }
  const tailSheet = engine.workbook.getSheetById(tailSheetId)
  if (tailSheet === undefined || tailSheet.order < request.sheet.order) {
    return null
  }
  const tailChanges = tryCreateLazyPhysicalTrackedIndexChanges(
    engine,
    changedCellIndices.subarray(tailStart),
    tailSheetId,
    formatTrackedAddress,
  )
  if (tailChanges === null) {
    return null
  }
  const literalChange: WorkPaperCellChange = {
    kind: 'cell',
    address: { sheet: request.address.sheet, row: request.address.row, col: request.address.col },
    sheetName: request.sheet.name,
    a1: runtime.trackedA1(request.address.row, request.address.col),
    newValue: scalarValueFromLiteral(request.value),
  }
  const changes = createPrefixedLazyTrackedIndexChanges([literalChange], tailChanges)
  runtime.trackLazyChanges(changes)
  return changes
}
