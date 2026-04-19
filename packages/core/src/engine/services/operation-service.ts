import { Effect } from 'effect'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import { ValueTag, type CellRangeRef, type CellValue, type EngineChangedCell, type EngineEvent, type SelectionState } from '@bilig/protocol'
import type { EdgeSlice } from '../../edge-arena.js'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import {
  entityPayload,
  isRangeEntity,
  makeCellEntity,
  makeExactLookupColumnEntity,
  makeSortedLookupColumnEntity,
} from '../../entity-ids.js'
import { batchOpOrder, compareOpOrder, createBatch, markBatchApplied, type OpOrder } from '../../replica-state.js'
import { CellFlags } from '../../cell-store.js'
import { emptyValue, literalToValue, writeLiteralToCellStore } from '../../engine-value-utils.js'
import { spillDependencyKey, tableDependencyKey } from '../../engine-metadata-utils.js'
import { makeCellKey, normalizeDefinedName, pivotKey, type WorkbookPivotRecord } from '../../workbook-store.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import type {
  EngineRuntimeState,
  PreparedCellAddress,
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectLookupDescriptor,
  U32,
} from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'
import type { EngineCellPatch } from '../../patches/patch-types.js'

type MutationSource = 'local' | 'remote' | 'restore' | 'undo' | 'redo'

type StructuralAxisOp = Extract<
  EngineOp,
  {
    kind: 'insertRows' | 'deleteRows' | 'moveRows' | 'insertColumns' | 'deleteColumns' | 'moveColumns'
  }
>

type DerivedOp = Extract<EngineOp, { kind: 'upsertSpillRange' | 'deleteSpillRange' | 'upsertPivotTable' | 'deletePivotTable' }>

export interface EngineOperationService {
  readonly applyBatch: (
    batch: EngineOpBatch,
    source: MutationSource,
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
  ) => Effect.Effect<void, EngineMutationError>
  readonly applyCellMutationsAt: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly applyDerivedOp: (op: DerivedOp) => Effect.Effect<number[], EngineMutationError>
}

interface VersionStore {
  get(key: string): OpOrder | undefined
  set(key: string, value: OpOrder): void
}

const noopVersionStore: VersionStore = {
  get() {
    return undefined
  },
  set() {
    return
  },
}

const FAST_LITERAL_OVERWRITE_FLAGS =
  CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput

function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function assertNever(value: never): never {
  throw new Error(`Unexpected value: ${String(value)}`)
}

function collectTrackedDependents(registry: Map<string | number, Set<number>>, keys: readonly (string | number)[]): number[] {
  const candidates = new Set<number>()
  keys.forEach((key) => {
    registry.get(key)?.forEach((cellIndex) => {
      candidates.add(cellIndex)
    })
  })
  return [...candidates]
}

function normalizeRange(range: CellRangeRef): CellRangeRef & {
  startRow: number
  endRow: number
  startCol: number
  endCol: number
} {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  return {
    ...range,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    startRow,
    endRow,
    startCol,
    endCol,
  }
}

function rangesIntersect(left: CellRangeRef, right: CellRangeRef): boolean {
  const a = normalizeRange(left)
  const b = normalizeRange(right)
  return !(a.sheetName !== b.sheetName || a.endRow < b.startRow || b.endRow < a.startRow || a.endCol < b.startCol || b.endCol < a.startCol)
}

function cellRange(sheetName: string, address: string): CellRangeRef {
  return {
    sheetName,
    startAddress: address,
    endAddress: address,
  }
}

function throwProtectionBlocked(message: string): never {
  throw new Error(`Workbook protection blocks this change: ${message}`)
}

function withOptionalLookupStringIds(request: {
  sheetName: string
  row: number
  col: number
  oldValue: CellValue
  newValue: CellValue
  oldStringId: number | undefined
  newStringId: number | undefined
}): {
  sheetName: string
  row: number
  col: number
  oldValue: CellValue
  newValue: CellValue
  oldStringId?: number
  newStringId?: number
} {
  return {
    sheetName: request.sheetName,
    row: request.row,
    col: request.col,
    oldValue: request.oldValue,
    newValue: request.newValue,
    ...(request.oldStringId === undefined ? {} : { oldStringId: request.oldStringId }),
    ...(request.newStringId === undefined ? {} : { newStringId: request.newStringId }),
  }
}

function normalizeExactLookupKey(value: CellValue, lookupString: (id: number) => string, stringId = 0): string | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'e:'
    case ValueTag.Number:
      return `n:${Object.is(value.value, -0) ? 0 : value.value}`
    case ValueTag.Boolean:
      return value.value ? 'b:1' : 'b:0'
    case ValueTag.String:
      return `s:${(stringId !== 0 ? lookupString(stringId) : value.value).toUpperCase()}`
    case ValueTag.Error:
      return undefined
  }
}

function normalizeApproximateNumericValue(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return 0
    case ValueTag.Number:
      return Object.is(value.value, -0) ? 0 : value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.String:
    case ValueTag.Error:
      return undefined
  }
}

function normalizeApproximateTextValue(value: CellValue, lookupString: (id: number) => string, stringId = 0): string | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.String:
      return (stringId !== 0 ? lookupString(stringId) : value.value).toUpperCase()
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.Error:
      return undefined
  }
}

function directCriteriaTouchesPoint(
  directCriteria: RuntimeDirectCriteriaDescriptor,
  request: { sheetName: string; row: number; col: number },
): boolean {
  if (directCriteria.aggregateRange) {
    const aggregateRange = directCriteria.aggregateRange
    if (
      aggregateRange.sheetName === request.sheetName &&
      aggregateRange.col === request.col &&
      request.row >= aggregateRange.rowStart &&
      request.row <= aggregateRange.rowEnd
    ) {
      return true
    }
  }
  return directCriteria.criteriaPairs.some(
    (pair) =>
      pair.range.sheetName === request.sheetName &&
      pair.range.col === request.col &&
      request.row >= pair.range.rowStart &&
      request.row <= pair.range.rowEnd,
  )
}

function mergeChangedCellIndices(base: readonly number[] | U32, extras: readonly number[]): U32 {
  if (extras.length === 0) {
    return base instanceof Uint32Array ? base : Uint32Array.from(base)
  }
  const merged = new Set<number>()
  for (let index = 0; index < base.length; index += 1) {
    merged.add(base[index]!)
  }
  for (let index = 0; index < extras.length; index += 1) {
    merged.add(extras[index]!)
  }
  return Uint32Array.from(merged)
}

export function createEngineOperationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | 'workbook'
    | 'strings'
    | 'events'
    | 'formulas'
    | 'replicaState'
    | 'entityVersions'
    | 'sheetDeleteVersions'
    | 'batchListeners'
    | 'redoStack'
    | 'trackReplicaVersions'
    | 'getSyncClientConnection'
    | 'getLastMetrics'
    | 'setLastMetrics'
  >
  readonly reverseState: {
    readonly reverseSpillEdges: Map<string, Set<number>>
    readonly reverseAggregateColumnEdges: Map<number, Set<number>>
    readonly reverseExactLookupColumnEdges: Map<number, EdgeSlice>
    readonly reverseSortedLookupColumnEdges: Map<number, EdgeSlice>
  }
  readonly getSelectionState: () => SelectionState
  readonly setSelection: (sheetName: string, address: string) => void
  readonly rewriteDefinedNamesForSheetRename: (oldSheetName: string, newSheetName: string) => void
  readonly rewriteCellFormulasForSheetRename: (oldSheetName: string, newSheetName: string, formulaChangedCount: number) => number
  readonly rebindDefinedNameDependents: (names: readonly string[], formulaChangedCount: number) => number
  readonly rebindTableDependents: (tableNames: readonly string[], formulaChangedCount: number) => number
  readonly rebindFormulaCells: (candidates: readonly number[], formulaChangedCount: number) => number
  readonly refreshRangeDependencies: (rangeIndices: readonly number[]) => void
  readonly rebindFormulasForSheet: (sheetName: string, formulaChangedCount: number, candidates?: readonly number[] | U32) => number
  readonly removeSheetRuntime: (
    sheetName: string,
    explicitChangedCount: number,
  ) => { changedInputCount: number; formulaChangedCount: number; explicitChangedCount: number }
  readonly applyStructuralAxisOp: (op: StructuralAxisOp) => {
    transaction: StructuralTransaction
    changedCellIndices: number[]
    formulaCellIndices: number[]
    topologyChanged: boolean
    graphRefreshRequired: boolean
  }
  readonly clearOwnedSpill: (cellIndex: number) => number[]
  readonly clearPivotForCell: (cellIndex: number) => number[]
  readonly clearOwnedPivot: (pivot: WorkbookPivotRecord) => number[]
  readonly materializePivot: (pivot: WorkbookPivotRecord) => number[]
  readonly removeFormula: (cellIndex: number) => boolean
  readonly bindFormula: (cellIndex: number, ownerSheetName: string, source: string) => boolean
  readonly setInvalidFormulaValue: (cellIndex: number) => void
  readonly beginMutationCollection: () => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly markVolatileFormulasChanged: (count: number) => number
  readonly markSpillRootsChanged: (cellIndices: readonly number[], count: number) => number
  readonly markPivotRootsChanged: (cellIndices: readonly number[], count: number) => number
  readonly markExplicitChanged: (cellIndex: number, count: number) => number
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32
  readonly composeEventChanges: (recalculated: U32, explicitChangedCount: number) => U32
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
  readonly captureChangedPatches: (changedCellIndices: readonly number[] | U32) => readonly EngineCellPatch[]
  readonly getChangedInputBuffer: () => U32
  readonly getChangedFormulaBuffer: () => U32
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly ensureCellTracked: (sheetName: string, address: string) => number
  readonly estimatePotentialNewCells: (ops: readonly EngineOp[]) => number
  readonly resetMaterializedCellScratch: (expectedSize: number) => void
  readonly syncDynamicRanges: (formulaChangedCount: number) => number
  readonly rebuildTopoRanks: () => void
  readonly repairTopoRanks: (changedFormulaCells: readonly number[] | U32) => boolean
  readonly detectCycles: () => void
  readonly recalculate: (changedRoots: readonly number[] | U32, kernelSyncRoots?: readonly number[] | U32) => U32
  readonly evaluateDirectFormula: (cellIndex: number) => readonly number[] | undefined
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => U32
  readonly prepareRegionQueryIndices: () => void
  readonly flushWasmProgramSync: () => void
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly getEntityDependents: (entityId: number) => Uint32Array
  readonly collectFormulaDependents: (entityId: number) => Uint32Array
  readonly hasRegionFormulaSubscriptionsForColumn: (sheetName: string, col: number) => boolean
  readonly collectRegionFormulaDependentsForCell: (sheetName: string, row: number, col: number) => Uint32Array
  readonly noteExactLookupLiteralWrite: (request: {
    sheetName: string
    row: number
    col: number
    oldValue: CellValue
    newValue: CellValue
    oldStringId?: number
    newStringId?: number
  }) => void
  readonly noteAggregateLiteralWrite: (request: {
    sheetName: string
    row: number
    col: number
    oldValue: CellValue
    newValue: CellValue
  }) => void
  readonly noteSortedLookupLiteralWrite: (request: {
    sheetName: string
    row: number
    col: number
    oldValue: CellValue
    newValue: CellValue
    oldStringId?: number
    newStringId?: number
  }) => void
  readonly invalidateExactLookupColumn: (request: { sheetName: string; col: number }) => void
  readonly invalidateSortedLookupColumn: (request: { sheetName: string; col: number }) => void
  readonly invalidateAggregateColumn: (request: { sheetName: string; col: number }) => void
}): EngineOperationService {
  const emitBatch = (batch: EngineOpBatch): void => {
    args.state.batchListeners.forEach((listener) => {
      listener(batch)
    })
  }
  const entityVersions: VersionStore = args.state.trackReplicaVersions ? args.state.entityVersions : noopVersionStore
  const sheetDeleteVersions: VersionStore = args.state.trackReplicaVersions ? args.state.sheetDeleteVersions : noopVersionStore
  const setEntityVersionForOp = (op: EngineOp, order: OpOrder): void => {
    if (!args.state.trackReplicaVersions) {
      return
    }
    entityVersions.set(entityKeyForOp(op), order)
  }
  const setCellEntityVersion = (sheetName: string, address: string, order: OpOrder): void => {
    if (!args.state.trackReplicaVersions) {
      return
    }
    entityVersions.set(`cell:${sheetName}!${address}`, order)
  }
  const setSheetDeleteVersion = (sheetName: string, order: OpOrder): void => {
    if (!args.state.trackReplicaVersions) {
      return
    }
    sheetDeleteVersions.set(sheetName, order)
  }

  const sheetHasProtection = (sheetName: string): boolean =>
    args.state.workbook.getSheetProtection(sheetName) !== undefined || args.state.workbook.listRangeProtections(sheetName).length > 0

  const rangeIsProtected = (range: CellRangeRef): boolean => {
    if (args.state.workbook.getSheetProtection(range.sheetName)) {
      return true
    }
    return args.state.workbook.listRangeProtections(range.sheetName).some((protection) => rangesIntersect(protection.range, range))
  }

  const assertProtectionAllowsOp = (op: EngineOp): void => {
    switch (op.kind) {
      case 'setSheetProtection':
      case 'clearSheetProtection':
      case 'upsertRangeProtection':
      case 'deleteRangeProtection':
      case 'upsertWorkbook':
      case 'setWorkbookMetadata':
      case 'setCalculationSettings':
      case 'setVolatileContext':
      case 'upsertDefinedName':
      case 'deleteDefinedName':
      case 'upsertCellStyle':
      case 'upsertCellNumberFormat':
        return
      case 'upsertSheet':
        return
      case 'renameSheet':
      case 'deleteSheet':
        if (sheetHasProtection(op.kind === 'renameSheet' ? op.oldName : op.name)) {
          throwProtectionBlocked(`sheet ${op.kind === 'renameSheet' ? op.oldName : op.name} is protected`)
        }
        return
      case 'insertRows':
      case 'deleteRows':
      case 'moveRows':
      case 'insertColumns':
      case 'deleteColumns':
      case 'moveColumns':
      case 'updateRowMetadata':
      case 'updateColumnMetadata':
      case 'setFreezePane':
      case 'clearFreezePane':
        if (sheetHasProtection(op.sheetName)) {
          throwProtectionBlocked(`sheet ${op.sheetName} is protected`)
        }
        return
      case 'setFilter':
      case 'clearFilter':
      case 'setSort':
      case 'clearSort':
      case 'setStyleRange':
      case 'setFormatRange':
        if (rangeIsProtected(op.range)) {
          throwProtectionBlocked(`range ${op.range.sheetName}!${op.range.startAddress}:${op.range.endAddress} is protected`)
        }
        return
      case 'setDataValidation':
        if (rangeIsProtected(op.validation.range)) {
          throwProtectionBlocked(
            `range ${op.validation.range.sheetName}!${op.validation.range.startAddress}:${op.validation.range.endAddress} is protected`,
          )
        }
        return
      case 'clearDataValidation':
        if (rangeIsProtected(op.range)) {
          throwProtectionBlocked(`range ${op.range.sheetName}!${op.range.startAddress}:${op.range.endAddress} is protected`)
        }
        return
      case 'upsertConditionalFormat':
        if (rangeIsProtected(op.format.range)) {
          throwProtectionBlocked(
            `range ${op.format.range.sheetName}!${op.format.range.startAddress}:${op.format.range.endAddress} is protected`,
          )
        }
        return
      case 'deleteConditionalFormat': {
        const existing = args.state.workbook.getConditionalFormat(op.id)
        if (existing && rangeIsProtected(existing.range)) {
          throwProtectionBlocked(`conditional format ${op.id} targets a protected range`)
        }
        return
      }
      case 'upsertCommentThread':
        if (rangeIsProtected(cellRange(op.thread.sheetName, op.thread.address))) {
          throwProtectionBlocked(`cell ${op.thread.sheetName}!${op.thread.address} is protected`)
        }
        return
      case 'deleteCommentThread':
        if (rangeIsProtected(cellRange(op.sheetName, op.address))) {
          throwProtectionBlocked(`cell ${op.sheetName}!${op.address} is protected`)
        }
        return
      case 'upsertNote':
        if (rangeIsProtected(cellRange(op.note.sheetName, op.note.address))) {
          throwProtectionBlocked(`cell ${op.note.sheetName}!${op.note.address} is protected`)
        }
        return
      case 'deleteNote':
        if (rangeIsProtected(cellRange(op.sheetName, op.address))) {
          throwProtectionBlocked(`cell ${op.sheetName}!${op.address} is protected`)
        }
        return
      case 'setCellValue':
      case 'setCellFormula':
      case 'setCellFormat':
      case 'clearCell':
        if (rangeIsProtected(cellRange(op.sheetName, op.address))) {
          throwProtectionBlocked(`cell ${op.sheetName}!${op.address} is protected`)
        }
        return
      case 'upsertTable':
        if (
          rangeIsProtected({
            sheetName: op.table.sheetName,
            startAddress: op.table.startAddress,
            endAddress: op.table.endAddress,
          })
        ) {
          throwProtectionBlocked(`table ${op.table.name} overlaps a protected range`)
        }
        return
      case 'deleteTable': {
        const existing = args.state.workbook.getTable(op.name)
        if (
          existing &&
          rangeIsProtected({
            sheetName: existing.sheetName,
            startAddress: existing.startAddress,
            endAddress: existing.endAddress,
          })
        ) {
          throwProtectionBlocked(`table ${op.name} overlaps a protected range`)
        }
        return
      }
      case 'upsertSpillRange':
      case 'deleteSpillRange':
        if (rangeIsProtected(cellRange(op.sheetName, op.address))) {
          throwProtectionBlocked(`cell ${op.sheetName}!${op.address} is protected`)
        }
        return
      case 'upsertPivotTable':
        if (sheetHasProtection(op.sheetName) || rangeIsProtected(op.source) || rangeIsProtected(cellRange(op.sheetName, op.address))) {
          throwProtectionBlocked(`pivot ${op.name} touches protected workbook state`)
        }
        return
      case 'deletePivotTable': {
        const existing = args.state.workbook.getPivot(op.sheetName, op.address)
        if (
          existing &&
          (sheetHasProtection(existing.sheetName) ||
            rangeIsProtected(existing.source) ||
            rangeIsProtected(cellRange(existing.sheetName, existing.address)))
        ) {
          throwProtectionBlocked(`pivot at ${op.sheetName}!${op.address} touches protected workbook state`)
        }
        return
      }
      case 'upsertChart':
        if (
          sheetHasProtection(op.chart.sheetName) ||
          rangeIsProtected(op.chart.source) ||
          rangeIsProtected(cellRange(op.chart.sheetName, op.chart.address))
        ) {
          throwProtectionBlocked(`chart ${op.chart.id} touches protected workbook state`)
        }
        return
      case 'deleteChart': {
        const existing = args.state.workbook.getChart(op.id)
        if (
          existing &&
          (sheetHasProtection(existing.sheetName) ||
            rangeIsProtected(existing.source) ||
            rangeIsProtected(cellRange(existing.sheetName, existing.address)))
        ) {
          throwProtectionBlocked(`chart ${op.id} touches protected workbook state`)
        }
        return
      }
      case 'upsertImage':
        if (sheetHasProtection(op.image.sheetName) || rangeIsProtected(cellRange(op.image.sheetName, op.image.address))) {
          throwProtectionBlocked(`image ${op.image.id} touches protected workbook state`)
        }
        return
      case 'deleteImage': {
        const existing = args.state.workbook.getImage(op.id)
        if (existing && (sheetHasProtection(existing.sheetName) || rangeIsProtected(cellRange(existing.sheetName, existing.address)))) {
          throwProtectionBlocked(`image ${op.id} touches protected workbook state`)
        }
        return
      }
      case 'upsertShape':
        if (sheetHasProtection(op.shape.sheetName) || rangeIsProtected(cellRange(op.shape.sheetName, op.shape.address))) {
          throwProtectionBlocked(`shape ${op.shape.id} touches protected workbook state`)
        }
        return
      case 'deleteShape': {
        const existing = args.state.workbook.getShape(op.id)
        if (existing && (sheetHasProtection(existing.sheetName) || rangeIsProtected(cellRange(existing.sheetName, existing.address)))) {
          throwProtectionBlocked(`shape ${op.id} touches protected workbook state`)
        }
        return
      }
      default:
        assertNever(op)
        return
    }
  }

  const readCellValueForLookup = (cellIndex: number | undefined): { value: CellValue; stringId: number | undefined } => {
    if (cellIndex === undefined) {
      return { value: emptyValue(), stringId: undefined }
    }
    const stringId = args.state.workbook.cellStore.stringIds[cellIndex]
    return {
      value: args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id)),
      stringId,
    }
  }

  const readCellValueAtForLookup = (sheetName: string, row: number, col: number): { value: CellValue; stringId: number | undefined } => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return { value: emptyValue(), stringId: undefined }
    }
    return readCellValueForLookup(sheet.logical.getVisibleCell(row, col))
  }

  const isLocallySortedNumericWrite = (
    sheetName: string,
    row: number,
    col: number,
    rowStart: number,
    rowEnd: number,
    matchMode: 1 | -1,
  ): boolean => {
    const current = normalizeApproximateNumericValue(readCellValueAtForLookup(sheetName, row, col).value)
    if (current === undefined) {
      return false
    }
    if (row > rowStart) {
      const previous = normalizeApproximateNumericValue(readCellValueAtForLookup(sheetName, row - 1, col).value)
      if (previous === undefined || (matchMode === 1 ? previous > current : previous < current)) {
        return false
      }
    }
    if (row < rowEnd) {
      const next = normalizeApproximateNumericValue(readCellValueAtForLookup(sheetName, row + 1, col).value)
      if (next === undefined || (matchMode === 1 ? current > next : current < next)) {
        return false
      }
    }
    return true
  }

  const isLocallySortedTextWrite = (
    sheetName: string,
    row: number,
    col: number,
    rowStart: number,
    rowEnd: number,
    matchMode: 1 | -1,
  ): boolean => {
    const currentCell = readCellValueAtForLookup(sheetName, row, col)
    const current = normalizeApproximateTextValue(currentCell.value, (id) => args.state.strings.get(id), currentCell.stringId)
    if (current === undefined) {
      return false
    }
    if (row > rowStart) {
      const previousCell = readCellValueAtForLookup(sheetName, row - 1, col)
      const previous = normalizeApproximateTextValue(previousCell.value, (id) => args.state.strings.get(id), previousCell.stringId)
      if (previous === undefined || (matchMode === 1 ? previous > current : previous < current)) {
        return false
      }
    }
    if (row < rowEnd) {
      const nextCell = readCellValueAtForLookup(sheetName, row + 1, col)
      const next = normalizeApproximateTextValue(nextCell.value, (id) => args.state.strings.get(id), nextCell.stringId)
      if (next === undefined || (matchMode === 1 ? current > next : current < next)) {
        return false
      }
    }
    return true
  }

  const canSkipApproximateLookupDirtyMark = (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate' | 'approximate-uniform-numeric' }>,
    request: {
      sheetName: string
      row: number
      col: number
      oldValue: CellValue
      newValue: CellValue
      oldStringId?: number
      newStringId?: number
    },
  ): boolean => {
    const rowStart = directLookup.kind === 'approximate' ? directLookup.prepared.rowStart : directLookup.rowStart
    const rowEnd = directLookup.kind === 'approximate' ? directLookup.prepared.rowEnd : directLookup.rowEnd
    const matchMode = directLookup.kind === 'approximate' ? directLookup.matchMode : directLookup.matchMode
    const operand = readCellValueForLookup(directLookup.operandCellIndex)
    const operandNumeric = normalizeApproximateNumericValue(operand.value)
    if (operandNumeric !== undefined) {
      const oldNumeric = normalizeApproximateNumericValue(request.oldValue)
      const newNumeric = normalizeApproximateNumericValue(request.newValue)
      if (oldNumeric === undefined || newNumeric === undefined) {
        return false
      }
      if (matchMode === 1) {
        return (
          oldNumeric > operandNumeric &&
          newNumeric > operandNumeric &&
          isLocallySortedNumericWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode)
        )
      }
      return (
        oldNumeric < operandNumeric &&
        newNumeric < operandNumeric &&
        isLocallySortedNumericWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode)
      )
    }
    const operandText = normalizeApproximateTextValue(operand.value, (id) => args.state.strings.get(id), operand.stringId)
    if (operandText === undefined) {
      return false
    }
    const oldText = normalizeApproximateTextValue(request.oldValue, (id) => args.state.strings.get(id), request.oldStringId)
    const newText = normalizeApproximateTextValue(request.newValue, (id) => args.state.strings.get(id), request.newStringId)
    if (oldText === undefined || newText === undefined) {
      return false
    }
    if (matchMode === 1) {
      return (
        oldText > operandText &&
        newText > operandText &&
        isLocallySortedTextWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode)
      )
    }
    return (
      oldText < operandText &&
      newText < operandText &&
      isLocallySortedTextWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode)
    )
  }

  const pruneCellIfOrphaned = (cellIndex: number): void => {
    if (args.collectFormulaDependents(makeCellEntity(cellIndex)).length > 0) {
      return
    }
    args.state.workbook.pruneCellIfEmpty(cellIndex)
  }

  const normalizeHistoryDependencyPlaceholder = (cellIndex: number, source: MutationSource): void => {
    if (source !== 'undo' && source !== 'restore') {
      return
    }
    if (args.state.workbook.getCellFormat(cellIndex) !== undefined) {
      return
    }
    const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
    if (
      (flags &
        (CellFlags.HasFormula |
          CellFlags.JsOnly |
          CellFlags.InCycle |
          CellFlags.SpillChild |
          CellFlags.PivotOutput |
          CellFlags.PendingDelete)) !==
      0
    ) {
      return
    }
    const value = args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id))
    if (value.tag !== ValueTag.Empty) {
      return
    }
    if (args.collectFormulaDependents(makeCellEntity(cellIndex)).length === 0) {
      return
    }
    args.state.workbook.cellStore.versions[cellIndex] = 0
  }

  const markCycleMemberInputsChanged = (changedInputCount: number): number => {
    args.state.formulas.forEach((_formula, cellIndex) => {
      if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) === 0) {
        return
      }
      changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
    })
    return changedInputCount
  }

  const hasCycleMembersNow = (): boolean => {
    let found = false
    args.state.formulas.forEach((_formula, cellIndex) => {
      if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
        found = true
      }
    })
    return found
  }

  const hasTrackedExactLookupDependents = (sheetId: number, col: number): boolean => {
    const slice = args.reverseState.reverseExactLookupColumnEdges.get(entityPayload(makeExactLookupColumnEntity(sheetId, col)))
    return slice !== undefined && slice.len > 0
  }

  const hasTrackedSortedLookupDependents = (sheetId: number, col: number): boolean => {
    const slice = args.reverseState.reverseSortedLookupColumnEdges.get(entityPayload(makeSortedLookupColumnEntity(sheetId, col)))
    return slice !== undefined && slice.len > 0
  }

  const hasTrackedDirectRangeDependents = (sheetId: number, col: number): boolean => {
    const sheetName = args.state.workbook.getSheetNameById(sheetId)
    return sheetName ? args.hasRegionFormulaSubscriptionsForColumn(sheetName, col) : false
  }

  const markAffectedExactLookupDependents = (
    request: {
      sheetName: string
      row: number
      col: number
      oldValue: CellValue
      newValue: CellValue
      oldStringId?: number
      newStringId?: number
    },
    formulaChangedCount: number,
  ): number => {
    const sheet = args.state.workbook.getSheet(request.sheetName)
    if (!sheet) {
      return formulaChangedCount
    }
    const dependents = args.collectFormulaDependents(makeExactLookupColumnEntity(sheet.id, request.col))
    if (dependents.length === 0) {
      return formulaChangedCount
    }
    const oldKey = normalizeExactLookupKey(request.oldValue, (id) => args.state.strings.get(id), request.oldStringId)
    const newKey = normalizeExactLookupKey(request.newValue, (id) => args.state.strings.get(id), request.newStringId)
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      const formula = args.state.formulas.get(formulaCellIndex)
      const directLookup = formula?.directLookup
      if (!directLookup) {
        continue
      }
      if (directLookup.kind !== 'exact' && directLookup.kind !== 'exact-uniform-numeric') {
        continue
      }
      const rowStart = directLookup.kind === 'exact' ? directLookup.prepared.rowStart : directLookup.rowStart
      const rowEnd = directLookup.kind === 'exact' ? directLookup.prepared.rowEnd : directLookup.rowEnd
      if (request.row < rowStart || request.row > rowEnd) {
        continue
      }
      const operand = readCellValueForLookup(directLookup.operandCellIndex)
      const operandKey = normalizeExactLookupKey(operand.value, (id) => args.state.strings.get(id), operand.stringId)
      if (operandKey === undefined || (operandKey !== oldKey && operandKey !== newKey)) {
        continue
      }
      formulaChangedCount = args.markFormulaChanged(formulaCellIndex, formulaChangedCount)
    }
    return formulaChangedCount
  }

  const markAffectedApproximateLookupDependents = (
    request: {
      sheetName: string
      row: number
      col: number
      oldValue: CellValue
      newValue: CellValue
      oldStringId?: number
      newStringId?: number
    },
    formulaChangedCount: number,
  ): number => {
    const sheet = args.state.workbook.getSheet(request.sheetName)
    if (!sheet) {
      return formulaChangedCount
    }
    const dependents = args.collectFormulaDependents(makeSortedLookupColumnEntity(sheet.id, request.col))
    if (dependents.length === 0) {
      return formulaChangedCount
    }
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
      if (directLookup?.kind !== 'approximate' && directLookup?.kind !== 'approximate-uniform-numeric') {
        continue
      }
      const rowStart = directLookup.kind === 'approximate' ? directLookup.prepared.rowStart : directLookup.rowStart
      const rowEnd = directLookup.kind === 'approximate' ? directLookup.prepared.rowEnd : directLookup.rowEnd
      if (request.row < rowStart || request.row > rowEnd) {
        continue
      }
      if (canSkipApproximateLookupDirtyMark(directLookup, request)) {
        continue
      }
      formulaChangedCount = args.markFormulaChanged(formulaCellIndex, formulaChangedCount)
    }
    return formulaChangedCount
  }

  const collectAffectedDirectRangeDependents = (request: { sheetName: string; row: number; col: number }): number[] => {
    const dependents = args.collectRegionFormulaDependentsForCell(request.sheetName, request.row, request.col)
    if (dependents.length === 0) {
      return []
    }
    return [...dependents].filter((formulaCellIndex) => {
      const formula = args.state.formulas.get(formulaCellIndex)
      const directAggregate = formula?.directAggregate
      if (
        directAggregate &&
        directAggregate.sheetName === request.sheetName &&
        directAggregate.col === request.col &&
        request.row >= directAggregate.rowStart &&
        request.row <= directAggregate.rowEnd
      ) {
        return true
      }
      const directCriteria = formula?.directCriteria
      return directCriteria ? directCriteriaTouchesPoint(directCriteria, request) : false
    })
  }

  const markAffectedDirectRangeDependents = (
    request: {
      sheetName: string
      row: number
      col: number
    },
    formulaChangedCount: number,
    postRecalcDirectFormulaIndices?: Set<number>,
  ): number => {
    const dependents = collectAffectedDirectRangeDependents(request)
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (postRecalcDirectFormulaIndices && (args.state.formulas.get(formulaCellIndex)?.dependencyIndices.length ?? 0) > 0) {
        postRecalcDirectFormulaIndices.add(formulaCellIndex)
      }
      formulaChangedCount = args.markFormulaChanged(formulaCellIndex, formulaChangedCount)
    }
    return formulaChangedCount
  }

  const refreshDependentRangesAndRebindFormulaDependents = (cellIndex: number, formulaChangedCount: number): number => {
    const directDependents = args.getEntityDependents(makeCellEntity(cellIndex))
    const rangeIndices: number[] = []
    for (let index = 0; index < directDependents.length; index += 1) {
      const dependent = directDependents[index]!
      if (isRangeEntity(dependent)) {
        rangeIndices.push(entityPayload(dependent))
      }
    }
    if (rangeIndices.length > 0) {
      args.refreshRangeDependencies(rangeIndices)
    }
    const formulas = Array.from(args.collectFormulaDependents(makeCellEntity(cellIndex))).filter((candidate) => candidate !== cellIndex)
    if (formulas.length === 0) {
      return formulaChangedCount
    }
    return args.rebindFormulaCells(formulas, formulaChangedCount)
  }

  const entityKeyForOp = (op: EngineOp): string => {
    switch (op.kind) {
      case 'upsertWorkbook':
        return 'workbook'
      case 'setWorkbookMetadata':
        return `workbook-meta:${op.key}`
      case 'setCalculationSettings':
        return 'workbook-calc'
      case 'setVolatileContext':
        return 'workbook-volatile'
      case 'upsertSheet':
      case 'deleteSheet':
        return `sheet:${op.name}`
      case 'renameSheet':
        return `sheet:${op.oldName}`
      case 'insertRows':
      case 'deleteRows':
      case 'moveRows':
        return `row-structure:${op.sheetName}`
      case 'insertColumns':
      case 'deleteColumns':
      case 'moveColumns':
        return `column-structure:${op.sheetName}`
      case 'updateRowMetadata':
        return `row-meta:${op.sheetName}:${op.start}:${op.count}`
      case 'updateColumnMetadata':
        return `column-meta:${op.sheetName}:${op.start}:${op.count}`
      case 'setFreezePane':
      case 'clearFreezePane':
        return `freeze:${op.sheetName}`
      case 'setSheetProtection':
      case 'clearSheetProtection':
        return `sheet-protection:${op.kind === 'setSheetProtection' ? op.protection.sheetName : op.sheetName}`
      case 'setFilter':
      case 'clearFilter':
        return `filter:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
      case 'setSort':
      case 'clearSort':
        return `sort:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
      case 'setDataValidation':
        return `validation:${op.validation.range.sheetName}:${op.validation.range.startAddress}:${op.validation.range.endAddress}`
      case 'clearDataValidation':
        return `validation:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
      case 'upsertConditionalFormat':
        return `conditional-format:${op.format.id}`
      case 'deleteConditionalFormat':
        return `conditional-format:${op.id}`
      case 'upsertRangeProtection':
        return `range-protection:${op.protection.id}`
      case 'deleteRangeProtection':
        return `range-protection:${op.id}`
      case 'upsertCommentThread':
        return `comment:${op.thread.sheetName}!${op.thread.address}`
      case 'deleteCommentThread':
        return `comment:${op.sheetName}!${op.address}`
      case 'upsertNote':
        return `note:${op.note.sheetName}!${op.note.address}`
      case 'deleteNote':
        return `note:${op.sheetName}!${op.address}`
      case 'setCellFormat':
        return `format:${op.sheetName}!${op.address}`
      case 'upsertCellStyle':
        return `style:${op.style.id}`
      case 'upsertCellNumberFormat':
        return `number-format:${op.format.id}`
      case 'setStyleRange':
        return `style-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
      case 'setFormatRange':
        return `format-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
      case 'setCellValue':
      case 'setCellFormula':
      case 'clearCell':
        return `cell:${op.sheetName}!${op.address}`
      case 'upsertDefinedName':
      case 'deleteDefinedName':
        return `defined-name:${normalizeDefinedName(op.name)}`
      case 'upsertTable':
        return `table:${normalizeDefinedName(op.table.name)}`
      case 'deleteTable':
        return `table:${normalizeDefinedName(op.name)}`
      case 'upsertSpillRange':
      case 'deleteSpillRange':
        return `spill:${op.sheetName}!${op.address}`
      case 'upsertPivotTable':
      case 'deletePivotTable':
        return `pivot:${pivotKey(op.sheetName, op.address)}`
      case 'upsertChart':
        return `chart:${op.chart.id.trim().toUpperCase()}`
      case 'deleteChart':
        return `chart:${op.id.trim().toUpperCase()}`
      case 'upsertImage':
        return `image:${op.image.id.trim().toUpperCase()}`
      case 'deleteImage':
        return `image:${op.id.trim().toUpperCase()}`
      case 'upsertShape':
        return `shape:${op.shape.id.trim().toUpperCase()}`
      case 'deleteShape':
        return `shape:${op.id.trim().toUpperCase()}`
      default:
        return assertNever(op)
    }
  }
  const canFastPathLiteralOverwrite = (cellIndex: number): boolean => {
    const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
    return (flags & FAST_LITERAL_OVERWRITE_FLAGS) === 0 && args.state.formulas.get(cellIndex) === undefined
  }

  const isNullLiteralWriteNoOp = (cellIndex: number): boolean => {
    if (args.state.formulas.get(cellIndex) !== undefined) {
      return false
    }
    if (args.state.workbook.getCellFormat(cellIndex) !== undefined) {
      return false
    }
    const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
    if ((flags & (CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
      return false
    }
    const value = args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id))
    return value.tag === ValueTag.Empty
  }

  const isClearCellNoOp = (cellIndex: number): boolean => {
    if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.AuthoredBlank) !== 0) {
      return false
    }
    return isNullLiteralWriteNoOp(cellIndex)
  }

  const sheetDeleteBarrierForOp = (op: EngineOp): OpOrder | undefined => {
    switch (op.kind) {
      case 'upsertWorkbook':
      case 'setWorkbookMetadata':
      case 'setCalculationSettings':
      case 'setVolatileContext':
      case 'deleteSheet':
      case 'upsertDefinedName':
      case 'deleteDefinedName':
      case 'upsertTable':
      case 'deleteTable':
        return undefined
      case 'updateRowMetadata':
      case 'updateColumnMetadata':
      case 'insertRows':
      case 'deleteRows':
      case 'moveRows':
      case 'insertColumns':
      case 'deleteColumns':
      case 'moveColumns':
      case 'setFreezePane':
      case 'clearFreezePane':
      case 'clearSheetProtection':
      case 'setFilter':
      case 'clearFilter':
      case 'setSort':
      case 'clearSort':
      case 'clearDataValidation':
      case 'deleteConditionalFormat':
      case 'deleteRangeProtection':
      case 'deleteCommentThread':
      case 'deleteNote':
      case 'setCellFormat':
      case 'setCellValue':
      case 'setCellFormula':
      case 'clearCell':
      case 'upsertSpillRange':
      case 'deleteSpillRange':
      case 'deletePivotTable':
        return sheetDeleteVersions.get(op.sheetName)
      case 'setStyleRange':
      case 'setFormatRange':
        return sheetDeleteVersions.get(op.range.sheetName)
      case 'upsertCellNumberFormat':
      case 'upsertCellStyle':
        return undefined
      case 'upsertSheet':
        return sheetDeleteVersions.get(op.name)
      case 'renameSheet':
        return sheetDeleteVersions.get(op.oldName)
      case 'setDataValidation':
        return sheetDeleteVersions.get(op.validation.range.sheetName)
      case 'setSheetProtection':
        return sheetDeleteVersions.get(op.protection.sheetName)
      case 'upsertConditionalFormat':
        return sheetDeleteVersions.get(op.format.range.sheetName)
      case 'upsertRangeProtection':
        return sheetDeleteVersions.get(op.protection.range.sheetName)
      case 'upsertCommentThread':
        return sheetDeleteVersions.get(op.thread.sheetName)
      case 'upsertNote':
        return sheetDeleteVersions.get(op.note.sheetName)
      case 'upsertPivotTable':
        return sheetDeleteVersions.get(op.sheetName) ?? sheetDeleteVersions.get(op.source.sheetName)
      case 'upsertChart':
        return sheetDeleteVersions.get(op.chart.sheetName) ?? sheetDeleteVersions.get(op.chart.source.sheetName)
      case 'deleteChart':
        return undefined
      case 'upsertImage':
        return sheetDeleteVersions.get(op.image.sheetName)
      case 'deleteImage':
        return undefined
      case 'upsertShape':
        return sheetDeleteVersions.get(op.shape.sheetName)
      case 'deleteShape':
        return undefined
      default:
        return assertNever(op)
    }
  }

  const shouldApplyOp = (op: EngineOp, order: OpOrder): boolean => {
    const sheetDeleteOrder = sheetDeleteBarrierForOp(op)
    if (sheetDeleteOrder && compareOpOrder(order, sheetDeleteOrder) <= 0) {
      return false
    }
    const existingOrder = entityVersions.get(entityKeyForOp(op))
    if (existingOrder && compareOpOrder(order, existingOrder) <= 0) {
      return false
    }
    return true
  }

  const applySpillRangeOp = (op: Extract<EngineOp, { kind: 'upsertSpillRange' | 'deleteSpillRange' }>, order: OpOrder): number[] => {
    if (op.kind === 'upsertSpillRange') {
      args.state.workbook.setSpill(op.sheetName, op.address, op.rows, op.cols)
    } else {
      args.state.workbook.deleteSpill(op.sheetName, op.address)
    }
    setEntityVersionForOp(op, order)
    return collectTrackedDependents(args.reverseState.reverseSpillEdges, [spillDependencyKey(op.sheetName, op.address)])
  }

  const applyPivotUpsertOp = (op: Extract<EngineOp, { kind: 'upsertPivotTable' }>, order: OpOrder): number[] => {
    const pivot = {
      name: op.name,
      sheetName: op.sheetName,
      address: op.address,
      source: op.source,
      groupBy: op.groupBy,
      values: op.values,
      rows: op.rows,
      cols: op.cols,
    } satisfies WorkbookPivotRecord
    args.state.workbook.setPivot(pivot)
    setEntityVersionForOp(op, order)
    return args.materializePivot(pivot)
  }

  const applyPivotDeleteOp = (op: Extract<EngineOp, { kind: 'deletePivotTable' }>, order: OpOrder): number[] => {
    const pivot = args.state.workbook.getPivot(op.sheetName, op.address)
    if (!pivot) {
      setEntityVersionForOp(op, order)
      return []
    }
    const changedPivotOutputs = args.clearOwnedPivot(pivot)
    args.state.workbook.deletePivot(op.sheetName, op.address)
    setEntityVersionForOp(op, order)
    return changedPivotOutputs
  }

  const applyDerivedOpNow = (op: DerivedOp): number[] => {
    const batch = createBatch(args.state.replicaState, [op])
    const order = batchOpOrder(batch, 0)
    switch (op.kind) {
      case 'upsertSpillRange':
      case 'deleteSpillRange': {
        const candidates = applySpillRangeOp(op, order)
        args.rebindFormulaCells(candidates, 0)
        return candidates
      }
      case 'upsertPivotTable':
        return applyPivotUpsertOp(op, order)
      case 'deletePivotTable':
        return applyPivotDeleteOp(op, order)
      default:
        return assertNever(op)
    }
  }

  const applyBatchNow = (
    batch: EngineOpBatch,
    source: MutationSource,
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
  ): void => {
    if (preparedCellAddressesByOpIndex && preparedCellAddressesByOpIndex.length !== batch.ops.length) {
      throw new Error('Prepared cell addresses must align with batch operations')
    }
    const isRestore = source === 'restore'
    args.beginMutationCollection()
    let changedInputCount = 0
    let formulaChangedCount = 0
    let explicitChangedCount = 0
    let topologyChanged = false
    let sheetDeleted = false
    let structuralInvalidation = false
    let compileMs = 0
    const invalidatedRanges: CellRangeRef[] = []
    const invalidatedRows: { sheetName: string; startIndex: number; endIndex: number }[] = []
    const invalidatedColumns: { sheetName: string; startIndex: number; endIndex: number }[] = []
    const postRecalcDirectFormulaIndices = new Set<number>()
    let refreshAllPivots = false
    let appliedOps = 0
    const canSkipOrderChecks = source !== 'remote'
    const hadCycleMembersBefore = hasCycleMembersNow()

    const reservedNewCells = potentialNewCells ?? args.estimatePotentialNewCells(batch.ops)
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.capacity + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const preparedSheetIdByName = new Map<string, number>()
    const resolvePreparedSheetId = (sheetName: string, create: boolean): number | undefined => {
      const cachedSheetId = preparedSheetIdByName.get(sheetName)
      if (cachedSheetId !== undefined) {
        if (args.state.workbook.getSheetById(cachedSheetId)) {
          return cachedSheetId
        }
        preparedSheetIdByName.delete(sheetName)
      }
      const sheet = create ? args.state.workbook.getOrCreateSheet(sheetName) : args.state.workbook.getSheet(sheetName)
      if (!sheet) {
        return undefined
      }
      preparedSheetIdByName.set(sheetName, sheet.id)
      return sheet.id
    }
    const getPreparedExistingCellIndex = (
      sheetName: string,
      address: string,
      preparedCellAddress: PreparedCellAddress | null,
    ): number | undefined => {
      if (!preparedCellAddress) {
        return args.state.workbook.getCellIndex(sheetName, address)
      }
      const sheetId = resolvePreparedSheetId(sheetName, false)
      if (sheetId === undefined) {
        return undefined
      }
      return args.state.workbook.cellKeyToIndex.get(makeCellKey(sheetId, preparedCellAddress.row, preparedCellAddress.col))
    }
    const ensurePreparedCellTracked = (sheetName: string, address: string, preparedCellAddress: PreparedCellAddress | null): number => {
      if (!preparedCellAddress) {
        return args.ensureCellTracked(sheetName, address)
      }
      const sheetId = resolvePreparedSheetId(sheetName, true)
      if (sheetId === undefined) {
        throw new Error(`Unknown sheet: ${sheetName}`)
      }
      return args.state.workbook.ensureCellAt(sheetId, preparedCellAddress.row, preparedCellAddress.col).cellIndex
    }

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      if (!isRestore && source !== 'undo' && source !== 'redo') {
        batch.ops.forEach((op) => {
          assertProtectionAllowsOp(op)
        })
      }
      batch.ops.forEach((op, opIndex) => {
        const order = batchOpOrder(batch, opIndex)
        const preparedCellAddress = preparedCellAddressesByOpIndex?.[opIndex] ?? null
        if (!canSkipOrderChecks && !shouldApplyOp(op, order)) {
          return
        }

        switch (op.kind) {
          case 'upsertWorkbook':
            args.state.workbook.workbookName = op.name
            setEntityVersionForOp(op, order)
            break
          case 'setWorkbookMetadata':
            args.state.workbook.setWorkbookProperty(op.key, op.value)
            setEntityVersionForOp(op, order)
            break
          case 'setCalculationSettings':
            args.state.workbook.setCalculationSettings(op.settings)
            setEntityVersionForOp(op, order)
            break
          case 'setVolatileContext':
            args.state.workbook.setVolatileContext(op.context)
            setEntityVersionForOp(op, order)
            break
          case 'upsertSheet': {
            preparedSheetIdByName.delete(op.name)
            args.state.workbook.createSheet(op.name, op.order, op.id)
            setEntityVersionForOp(op, order)
            const tombstone = sheetDeleteVersions.get(op.name)
            if (!tombstone || compareOpOrder(order, tombstone) > 0) {
              args.state.sheetDeleteVersions.delete(op.name)
            }
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindFormulasForSheet(op.name, formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            refreshAllPivots = true
            break
          }
          case 'renameSheet': {
            preparedSheetIdByName.delete(op.oldName)
            preparedSheetIdByName.delete(op.newName)
            const renamedSheet = args.state.workbook.renameSheet(op.oldName, op.newName)
            if (args.state.trackReplicaVersions) {
              entityVersions.set(`sheet:${op.oldName}`, order)
              entityVersions.set(`sheet:${op.newName}`, order)
            }
            setSheetDeleteVersion(op.oldName, order)
            const renamedTombstone = sheetDeleteVersions.get(op.newName)
            if (!renamedTombstone || compareOpOrder(order, renamedTombstone) > 0) {
              args.state.sheetDeleteVersions.delete(op.newName)
            }
            if (!renamedSheet) {
              break
            }
            const selection = args.getSelectionState()
            if (selection.sheetName === op.oldName) {
              args.setSelection(op.newName, selection.address ?? 'A1')
            }
            args.rewriteDefinedNamesForSheetRename(op.oldName, op.newName)
            formulaChangedCount = args.rewriteCellFormulasForSheetRename(op.oldName, op.newName, formulaChangedCount)
            topologyChanged = true
            sheetDeleted = true
            structuralInvalidation = true
            refreshAllPivots = true
            break
          }
          case 'deleteSheet': {
            preparedSheetIdByName.delete(op.name)
            const removal = args.removeSheetRuntime(op.name, explicitChangedCount)
            changedInputCount += removal.changedInputCount
            formulaChangedCount += removal.formulaChangedCount
            explicitChangedCount = removal.explicitChangedCount
            setEntityVersionForOp(op, order)
            setSheetDeleteVersion(op.name, order)
            topologyChanged = true
            sheetDeleted = true
            structuralInvalidation = true
            refreshAllPivots = true
            break
          }
          case 'insertRows':
          case 'deleteRows':
          case 'moveRows':
          case 'insertColumns':
          case 'deleteColumns':
          case 'moveColumns': {
            const structural = args.applyStructuralAxisOp(op)
            structural.formulaCellIndices.forEach((cellIndex) => {
              formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
            })
            structural.transaction.invalidationSpans.forEach((invalidation) => {
              if (invalidation.axis === 'row') {
                invalidatedRows.push({
                  sheetName: op.sheetName,
                  startIndex: invalidation.start,
                  endIndex: invalidation.end - 1,
                })
                return
              }
              invalidatedColumns.push({
                sheetName: op.sheetName,
                startIndex: invalidation.start,
                endIndex: invalidation.end - 1,
              })
            })
            topologyChanged = structural.graphRefreshRequired || topologyChanged
            refreshAllPivots = true
            setEntityVersionForOp(op, order)
            break
          }
          case 'updateRowMetadata':
            args.state.workbook.setRowMetadata(op.sheetName, op.start, op.count, op.size, op.hidden)
            invalidatedRows.push({
              sheetName: op.sheetName,
              startIndex: op.start,
              endIndex: op.start + op.count - 1,
            })
            setEntityVersionForOp(op, order)
            break
          case 'updateColumnMetadata':
            args.state.workbook.setColumnMetadata(op.sheetName, op.start, op.count, op.size, op.hidden)
            invalidatedColumns.push({
              sheetName: op.sheetName,
              startIndex: op.start,
              endIndex: op.start + op.count - 1,
            })
            setEntityVersionForOp(op, order)
            break
          case 'setFreezePane':
            args.state.workbook.setFreezePane(op.sheetName, op.rows, op.cols)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'clearFreezePane':
            args.state.workbook.clearFreezePane(op.sheetName)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'setSheetProtection':
            args.state.workbook.setSheetProtection(op.protection)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'clearSheetProtection':
            args.state.workbook.clearSheetProtection(op.sheetName)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'setFilter':
            args.state.workbook.setFilter(op.sheetName, op.range)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'clearFilter':
            args.state.workbook.deleteFilter(op.sheetName, op.range)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'setSort':
            args.state.workbook.setSort(op.sheetName, op.range, op.keys)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'clearSort':
            args.state.workbook.deleteSort(op.sheetName, op.range)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'setDataValidation':
            args.state.workbook.setDataValidation(op.validation)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'clearDataValidation':
            args.state.workbook.deleteDataValidation(op.sheetName, op.range)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'upsertConditionalFormat':
            args.state.workbook.setConditionalFormat(op.format)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'deleteConditionalFormat':
            args.state.workbook.deleteConditionalFormat(op.id)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'upsertRangeProtection':
            args.state.workbook.setRangeProtection(op.protection)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'deleteRangeProtection':
            args.state.workbook.deleteRangeProtection(op.id)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'upsertCommentThread':
            args.state.workbook.setCommentThread(op.thread)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'deleteCommentThread':
            args.state.workbook.deleteCommentThread(op.sheetName, op.address)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'upsertNote':
            args.state.workbook.setNote(op.note)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'deleteNote':
            args.state.workbook.deleteNote(op.sheetName, op.address)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'upsertTable': {
            args.state.workbook.setTable(op.table)
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindTableDependents([tableDependencyKey(op.table.name)], formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            setEntityVersionForOp(op, order)
            break
          }
          case 'deleteTable': {
            args.state.workbook.deleteTable(op.name)
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindTableDependents([tableDependencyKey(op.name)], formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            setEntityVersionForOp(op, order)
            break
          }
          case 'upsertSpillRange':
          case 'deleteSpillRange': {
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindFormulaCells(applySpillRangeOp(op, order), formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            break
          }
          case 'setCellValue': {
            const existingIndex = getPreparedExistingCellIndex(op.sheetName, op.address, preparedCellAddress)
            const parsedAddress = preparedCellAddress ?? parseCellAddress(op.address, op.sheetName)
            const sheet = args.state.workbook.getSheet(op.sheetName)
            const sheetId = sheet?.id
            const hasExactLookupDependents = sheetId !== undefined ? hasTrackedExactLookupDependents(sheetId, parsedAddress.col) : false
            const hasSortedLookupDependents = sheetId !== undefined ? hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) : false
            const hasAggregateDependents = sheetId !== undefined ? hasTrackedDirectRangeDependents(sheetId, parsedAddress.col) : false
            const needsLookupValueRead = hasExactLookupDependents || hasSortedLookupDependents || hasAggregateDependents
            const prior = needsLookupValueRead ? readCellValueForLookup(existingIndex) : { value: emptyValue(), stringId: undefined }
            if (!isRestore) {
              if (op.value === null && (existingIndex === undefined || isNullLiteralWriteNoOp(existingIndex))) {
                break
              }
              if (existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              }
            }
            const cellIndex = ensurePreparedCellTracked(op.sheetName, op.address, preparedCellAddress)
            if (!isRestore) {
              changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
              const removedFormula = args.removeFormula(cellIndex)
              if (removedFormula) {
                args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
              }
              if (removedFormula) {
                formulaChangedCount = refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
              }
              topologyChanged = removedFormula || topologyChanged
            }
            writeLiteralToCellStore(args.state.workbook.cellStore, cellIndex, op.value, args.state.strings)
            if (op.value === null) {
              args.state.workbook.cellStore.flags[cellIndex] =
                (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.AuthoredBlank
            }
            args.state.workbook.notifyCellValueWritten(cellIndex)
            if (needsLookupValueRead) {
              const newValue = literalToValue(op.value, args.state.strings)
              const newStringId = typeof op.value === 'string' ? args.state.workbook.cellStore.stringIds[cellIndex] : undefined
              if (hasExactLookupDependents || hasAggregateDependents) {
                const exactLookupRequest = withOptionalLookupStringIds({
                  sheetName: op.sheetName,
                  row: parsedAddress.row,
                  col: parsedAddress.col,
                  oldValue: prior.value,
                  newValue,
                  oldStringId: prior.stringId,
                  newStringId,
                })
                if (hasExactLookupDependents) {
                  args.noteExactLookupLiteralWrite(exactLookupRequest)
                  formulaChangedCount = markAffectedExactLookupDependents(exactLookupRequest, formulaChangedCount)
                }
                if (hasAggregateDependents) {
                  args.noteAggregateLiteralWrite({
                    sheetName: exactLookupRequest.sheetName,
                    row: exactLookupRequest.row,
                    col: exactLookupRequest.col,
                    oldValue: exactLookupRequest.oldValue,
                    newValue: exactLookupRequest.newValue,
                  })
                  formulaChangedCount = markAffectedDirectRangeDependents(
                    exactLookupRequest,
                    formulaChangedCount,
                    postRecalcDirectFormulaIndices,
                  )
                }
              }
              if (hasSortedLookupDependents) {
                const sortedLookupRequest = withOptionalLookupStringIds({
                  sheetName: op.sheetName,
                  row: parsedAddress.row,
                  col: parsedAddress.col,
                  oldValue: prior.value,
                  newValue,
                  oldStringId: prior.stringId,
                  newStringId,
                })
                args.noteSortedLookupLiteralWrite(sortedLookupRequest)
                formulaChangedCount = markAffectedApproximateLookupDependents(sortedLookupRequest, formulaChangedCount)
              }
            }
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
            if (!isRestore && op.value === null) {
              pruneCellIfOrphaned(cellIndex)
            }
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            if (!isRestore) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              setEntityVersionForOp(op, order)
            }
            break
          }
          case 'setCellFormula': {
            const parsedAddress = parseCellAddress(op.address, op.sheetName)
            args.invalidateExactLookupColumn({ sheetName: op.sheetName, col: parsedAddress.col })
            args.invalidateSortedLookupColumn({ sheetName: op.sheetName, col: parsedAddress.col })
            if (!isRestore) {
              const existingIndex = getPreparedExistingCellIndex(op.sheetName, op.address, preparedCellAddress)
              if (existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              }
            }
            const cellIndex = ensurePreparedCellTracked(op.sheetName, op.address, preparedCellAddress)
            const priorHadFormula = args.state.formulas.get(cellIndex) !== undefined
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.AuthoredBlank
            if (!isRestore) {
              changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
            }
            const compileStarted = isRestore ? 0 : performance.now()
            try {
              const changedTopology = args.bindFormula(cellIndex, op.sheetName, op.formula)
              args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
              if (!isRestore) {
                compileMs += performance.now() - compileStarted
              }
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
              formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
              topologyChanged = topologyChanged || changedTopology
              if (!priorHadFormula) {
                formulaChangedCount = refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
                topologyChanged = true
              }
              const aggregateDependents = collectAffectedDirectRangeDependents({
                sheetName: op.sheetName,
                row: parsedAddress.row,
                col: parsedAddress.col,
              }).filter((candidate) => candidate !== cellIndex)
              if (aggregateDependents.length > 0) {
                formulaChangedCount = args.rebindFormulaCells(aggregateDependents, formulaChangedCount)
                for (let index = 0; index < aggregateDependents.length; index += 1) {
                  postRecalcDirectFormulaIndices.add(aggregateDependents[index]!)
                  formulaChangedCount = args.markFormulaChanged(aggregateDependents[index]!, formulaChangedCount)
                  changedInputCount = args.markInputChanged(aggregateDependents[index]!, changedInputCount)
                }
                topologyChanged = true
              }
            } catch {
              if (!isRestore) {
                compileMs += performance.now() - compileStarted
              }
              topologyChanged = args.removeFormula(cellIndex) || topologyChanged
              args.setInvalidFormulaValue(cellIndex)
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            }
            if (!isRestore) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              setEntityVersionForOp(op, order)
            }
            break
          }
          case 'setCellFormat': {
            const cellIndex = args.ensureCellTracked(op.sheetName, op.address)
            args.state.workbook.setCellFormat(cellIndex, op.format)
            pruneCellIfOrphaned(cellIndex)
            if (!isRestore) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              setEntityVersionForOp(op, order)
            }
            break
          }
          case 'upsertCellStyle':
            args.state.workbook.upsertCellStyle(op.style)
            setEntityVersionForOp(op, order)
            break
          case 'upsertCellNumberFormat':
            args.state.workbook.upsertCellNumberFormat(op.format)
            setEntityVersionForOp(op, order)
            break
          case 'setStyleRange':
            args.state.workbook.setStyleRange(op.range, op.styleId)
            if (source !== 'restore') {
              args.state.workbook.coalesceStyleRanges(op.range.sheetName)
            }
            invalidatedRanges.push(op.range)
            setEntityVersionForOp(op, order)
            break
          case 'setFormatRange':
            args.state.workbook.setFormatRange(op.range, op.formatId)
            invalidatedRanges.push(op.range)
            setEntityVersionForOp(op, order)
            break
          case 'clearCell': {
            const cellIndex = getPreparedExistingCellIndex(op.sheetName, op.address, preparedCellAddress)
            const parsedAddress = preparedCellAddress ?? parseCellAddress(op.address, op.sheetName)
            const sheet = args.state.workbook.getSheet(op.sheetName)
            const sheetId = sheet?.id
            const hasExactLookupDependents = sheetId !== undefined ? hasTrackedExactLookupDependents(sheetId, parsedAddress.col) : false
            const hasSortedLookupDependents = sheetId !== undefined ? hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) : false
            const hasAggregateDependents = sheetId !== undefined ? hasTrackedDirectRangeDependents(sheetId, parsedAddress.col) : false
            const needsLookupValueRead = hasExactLookupDependents || hasSortedLookupDependents || hasAggregateDependents
            const prior = needsLookupValueRead ? readCellValueForLookup(cellIndex) : { value: emptyValue(), stringId: undefined }
            if (cellIndex === undefined) {
              setEntityVersionForOp(op, order)
              break
            }
            if (isClearCellNoOp(cellIndex)) {
              break
            }
            changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(cellIndex), changedInputCount)
            changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
            const removedFormula = args.removeFormula(cellIndex)
            if (removedFormula) {
              args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
            }
            if (removedFormula) {
              formulaChangedCount = refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
            }
            topologyChanged = removedFormula || topologyChanged
            args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
            args.state.workbook.notifyCellValueWritten(cellIndex)
            if (needsLookupValueRead) {
              if (hasExactLookupDependents || hasAggregateDependents) {
                const exactLookupRequest = withOptionalLookupStringIds({
                  sheetName: op.sheetName,
                  row: parsedAddress.row,
                  col: parsedAddress.col,
                  oldValue: prior.value,
                  newValue: emptyValue(),
                  oldStringId: prior.stringId,
                  newStringId: undefined,
                })
                if (hasExactLookupDependents) {
                  args.noteExactLookupLiteralWrite(exactLookupRequest)
                  formulaChangedCount = markAffectedExactLookupDependents(exactLookupRequest, formulaChangedCount)
                }
                if (hasAggregateDependents) {
                  args.noteAggregateLiteralWrite({
                    sheetName: exactLookupRequest.sheetName,
                    row: exactLookupRequest.row,
                    col: exactLookupRequest.col,
                    oldValue: exactLookupRequest.oldValue,
                    newValue: exactLookupRequest.newValue,
                  })
                  formulaChangedCount = markAffectedDirectRangeDependents(
                    exactLookupRequest,
                    formulaChangedCount,
                    postRecalcDirectFormulaIndices,
                  )
                }
              }
              if (hasSortedLookupDependents) {
                const sortedLookupRequest = withOptionalLookupStringIds({
                  sheetName: op.sheetName,
                  row: parsedAddress.row,
                  col: parsedAddress.col,
                  oldValue: prior.value,
                  newValue: emptyValue(),
                  oldStringId: prior.stringId,
                  newStringId: undefined,
                })
                args.noteSortedLookupLiteralWrite(sortedLookupRequest)
                formulaChangedCount = markAffectedApproximateLookupDependents(sortedLookupRequest, formulaChangedCount)
              }
            }
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(
                CellFlags.HasFormula |
                CellFlags.JsOnly |
                CellFlags.InCycle |
                CellFlags.SpillChild |
                CellFlags.PivotOutput |
                CellFlags.AuthoredBlank
              )
            normalizeHistoryDependencyPlaceholder(cellIndex, source)
            pruneCellIfOrphaned(cellIndex)
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            setEntityVersionForOp(op, order)
            break
          }
          case 'upsertDefinedName': {
            const normalizedName = normalizeDefinedName(op.name)
            args.state.workbook.setDefinedName(op.name, op.value)
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindDefinedNameDependents([normalizedName], formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            setEntityVersionForOp(op, order)
            break
          }
          case 'deleteDefinedName': {
            const normalizedName = normalizeDefinedName(op.name)
            args.state.workbook.deleteDefinedName(op.name)
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindDefinedNameDependents([normalizedName], formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            setEntityVersionForOp(op, order)
            break
          }
          case 'upsertPivotTable': {
            const changedPivotUpsertOutputs = applyPivotUpsertOp(op, order)
            changedInputCount = args.markPivotRootsChanged(changedPivotUpsertOutputs, changedInputCount)
            changedPivotUpsertOutputs.forEach((cellIndex) => {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            })
            refreshAllPivots = true
            break
          }
          case 'deletePivotTable': {
            const changedPivotOutputs = applyPivotDeleteOp(op, order)
            changedInputCount = args.markPivotRootsChanged(changedPivotOutputs, changedInputCount)
            changedPivotOutputs.forEach((cellIndex) => {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            })
            refreshAllPivots = true
            break
          }
          case 'upsertChart':
            args.state.workbook.setChart(op.chart)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'deleteChart':
            args.state.workbook.deleteChart(op.id)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'upsertImage':
            args.state.workbook.setImage(op.image)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'deleteImage':
            args.state.workbook.deleteImage(op.id)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'upsertShape':
            args.state.workbook.setShape(op.shape)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          case 'deleteShape':
            args.state.workbook.deleteShape(op.id)
            structuralInvalidation = true
            setEntityVersionForOp(op, order)
            break
          default:
            assertNever(op)
        }
        appliedOps += 1
      })

      const reboundCount = formulaChangedCount
      formulaChangedCount = args.syncDynamicRanges(formulaChangedCount)
      topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
      args.flushWasmProgramSync()
    }

    markBatchApplied(args.state.replicaState, batch)
    if (appliedOps === 0) {
      if (source === 'local') {
        emitBatch(batch)
      }
      return
    }

    if (topologyChanged) {
      const repaired =
        !hadCycleMembersBefore &&
        !sheetDeleted &&
        !structuralInvalidation &&
        formulaChangedCount > 0 &&
        args.repairTopoRanks(args.getChangedFormulaBuffer().subarray(0, formulaChangedCount))
      if (!repaired) {
        args.rebuildTopoRanks()
        args.detectCycles()
        changedInputCount = markCycleMemberInputsChanged(changedInputCount)
      }
    }
    args.prepareRegionQueryIndices()
    const hasActiveFormulas = args.state.formulas.size > 0
    const hasActivePivots = args.state.workbook.listPivots().length > 0
    let recalculated: U32 = new Uint32Array()
    if (hasActiveFormulas || hasActivePivots || refreshAllPivots) {
      formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount)
      const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
      recalculated = args.recalculate(args.composeMutationRoots(changedInputCount, formulaChangedCount), changedInputArray)
      if (postRecalcDirectFormulaIndices.size > 0) {
        const postRecalcChanged: number[] = []
        postRecalcDirectFormulaIndices.forEach((cellIndex) => {
          if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
            return
          }
          const changedCellIndices = args.evaluateDirectFormula(cellIndex)
          postRecalcChanged.push(cellIndex)
          if (changedCellIndices) {
            for (let index = 0; index < changedCellIndices.length; index += 1) {
              postRecalcChanged.push(changedCellIndices[index]!)
            }
          }
        })
        recalculated = mergeChangedCellIndices(recalculated, postRecalcChanged)
      }
      recalculated = args.reconcilePivotOutputs(recalculated, refreshAllPivots)
    }
    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const changed: U32 = isRestore || !requiresChangedSet ? new Uint32Array() : args.composeEventChanges(recalculated, explicitChangedCount)
    const lastMetrics = {
      ...args.state.getLastMetrics(),
      batchId: args.state.getLastMetrics().batchId + 1,
      changedInputCount: changedInputCount + formulaChangedCount,
      compileMs,
    }
    args.state.setLastMetrics(lastMetrics)
    const invalidation = isRestore || sheetDeleted || structuralInvalidation ? 'full' : 'cells'
    if (hasGeneralEventListeners || hasWatchedCellListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation,
        changedCellIndices: changed,
        changedCells: hasGeneralEventListeners ? args.captureChangedCells(changed) : [],
        invalidatedRanges,
        invalidatedRows,
        invalidatedColumns,
        metrics: lastMetrics,
        explicitChangedCount,
      }
      if (event.invalidation === 'full') {
        args.state.events.emitAllWatched(event)
      } else {
        args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
      }
    }
    if (hasTrackedEventListeners) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation,
        changedCellIndices: changed,
        patches: args.captureChangedPatches(changed),
        invalidatedRanges,
        invalidatedRows,
        invalidatedColumns,
        metrics: lastMetrics,
        explicitChangedCount,
      })
    }
    if (source === 'local') {
      void args.state.getSyncClientConnection()?.send(batch)
      emitBatch(batch)
    } else if (source === 'remote' && args.state.redoStack.length > 0) {
      args.state.redoStack.length = 0
    }
  }

  const applyCellMutationsAtNow = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ): void => {
    const isRestore = source === 'restore'
    args.beginMutationCollection()
    let changedInputCount = 0
    let formulaChangedCount = 0
    let explicitChangedCount = 0
    let topologyChanged = false
    let compileMs = 0
    const postRecalcDirectFormulaIndices = new Set<number>()
    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const trackExplicitChanges = !isRestore && requiresChangedSet
    const hadCycleMembersBefore = hasCycleMembersNow()
    const reservedNewCells = potentialNewCells ?? refs.length
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.capacity + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const sheetNameById = new Map<number, string>()
    const resolveSheetName = (sheetId: number): string => {
      const cached = sheetNameById.get(sheetId)
      if (cached !== undefined) {
        return cached
      }
      const sheet = args.state.workbook.getSheetById(sheetId)
      if (!sheet) {
        throw new Error(`Unknown sheet id: ${sheetId}`)
      }
      sheetNameById.set(sheetId, sheet.name)
      return sheet.name
    }
    const trackedColumnDependencyFlagsBySheet = new Map<
      number,
      Map<
        number,
        {
          hasExactLookupDependents: boolean
          hasSortedLookupDependents: boolean
          hasAggregateDependents: boolean
          needsLookupValueRead: boolean
        }
      >
    >()
    const clearTrackedColumnDependencyFlagCache = (): void => {
      trackedColumnDependencyFlagsBySheet.clear()
    }
    const resolveTrackedColumnDependencyFlags = (
      sheetId: number,
      col: number,
    ): {
      hasExactLookupDependents: boolean
      hasSortedLookupDependents: boolean
      hasAggregateDependents: boolean
      needsLookupValueRead: boolean
    } => {
      let flagsByColumn = trackedColumnDependencyFlagsBySheet.get(sheetId)
      if (flagsByColumn === undefined) {
        flagsByColumn = new Map()
        trackedColumnDependencyFlagsBySheet.set(sheetId, flagsByColumn)
      }
      const cached = flagsByColumn.get(col)
      if (cached !== undefined) {
        return cached
      }
      const hasExactLookupDependents = hasTrackedExactLookupDependents(sheetId, col)
      const hasSortedLookupDependents = hasTrackedSortedLookupDependents(sheetId, col)
      const hasAggregateDependents = hasTrackedDirectRangeDependents(sheetId, col)
      const next = {
        hasExactLookupDependents,
        hasSortedLookupDependents,
        hasAggregateDependents,
        needsLookupValueRead: hasExactLookupDependents || hasSortedLookupDependents || hasAggregateDependents,
      }
      flagsByColumn.set(col, next)
      return next
    }

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        refs.forEach((ref, refIndex) => {
          const { sheetId, mutation } = ref
          const order = args.state.trackReplicaVersions && batch ? batchOpOrder(batch, refIndex) : undefined
          const existingIndex = args.state.workbook.cellKeyToIndex.get(makeCellKey(sheetId, mutation.row, mutation.col))

          switch (mutation.kind) {
            case 'setCellValue': {
              const sheetName = resolveSheetName(sheetId)
              const { hasExactLookupDependents, hasSortedLookupDependents, hasAggregateDependents, needsLookupValueRead } =
                resolveTrackedColumnDependencyFlags(sheetId, mutation.col)
              const prior = needsLookupValueRead ? readCellValueForLookup(existingIndex) : { value: emptyValue(), stringId: undefined }
              if (mutation.value === null && !isRestore && (existingIndex === undefined || isNullLiteralWriteNoOp(existingIndex))) {
                break
              }
              if (existingIndex !== undefined && canFastPathLiteralOverwrite(existingIndex)) {
                writeLiteralToCellStore(args.state.workbook.cellStore, existingIndex, mutation.value, args.state.strings)
                args.state.workbook.notifyCellValueWritten(existingIndex)
                if (needsLookupValueRead) {
                  const newValue = literalToValue(mutation.value, args.state.strings)
                  const newStringId =
                    typeof mutation.value === 'string' ? args.state.workbook.cellStore.stringIds[existingIndex] : undefined
                  if (hasExactLookupDependents || hasAggregateDependents) {
                    const exactLookupRequest = withOptionalLookupStringIds({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                      oldValue: prior.value,
                      newValue,
                      oldStringId: prior.stringId,
                      newStringId,
                    })
                    if (hasExactLookupDependents) {
                      args.noteExactLookupLiteralWrite(exactLookupRequest)
                      formulaChangedCount = markAffectedExactLookupDependents(exactLookupRequest, formulaChangedCount)
                    }
                    if (hasAggregateDependents) {
                      args.noteAggregateLiteralWrite({
                        sheetName: exactLookupRequest.sheetName,
                        row: exactLookupRequest.row,
                        col: exactLookupRequest.col,
                        oldValue: exactLookupRequest.oldValue,
                        newValue: exactLookupRequest.newValue,
                      })
                      formulaChangedCount = markAffectedDirectRangeDependents(
                        exactLookupRequest,
                        formulaChangedCount,
                        postRecalcDirectFormulaIndices,
                      )
                    }
                  }
                  if (hasSortedLookupDependents) {
                    const sortedLookupRequest = withOptionalLookupStringIds({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                      oldValue: prior.value,
                      newValue,
                      oldStringId: prior.stringId,
                      newStringId,
                    })
                    args.noteSortedLookupLiteralWrite(sortedLookupRequest)
                    formulaChangedCount = markAffectedApproximateLookupDependents(sortedLookupRequest, formulaChangedCount)
                  }
                }
                changedInputCount = args.markInputChanged(existingIndex, changedInputCount)
                if (trackExplicitChanges) {
                  explicitChangedCount = args.markExplicitChanged(existingIndex, explicitChangedCount)
                }
                if (!isRestore && args.state.trackReplicaVersions) {
                  setCellEntityVersion(sheetName, formatAddress(mutation.row, mutation.col), order!)
                }
                break
              }
              if (existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              }
              const cellIndex = args.state.workbook.ensureCellAt(sheetId, mutation.row, mutation.col).cellIndex
              if (!isRestore) {
                changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
                const removedFormula = args.removeFormula(cellIndex)
                topologyChanged = removedFormula || topologyChanged
                if (removedFormula) {
                  args.invalidateAggregateColumn({ sheetName, col: mutation.col })
                }
                if (removedFormula) {
                  clearTrackedColumnDependencyFlagCache()
                }
              }
              writeLiteralToCellStore(args.state.workbook.cellStore, cellIndex, mutation.value, args.state.strings)
              args.state.workbook.notifyCellValueWritten(cellIndex)
              if (needsLookupValueRead) {
                const newValue = literalToValue(mutation.value, args.state.strings)
                const newStringId = typeof mutation.value === 'string' ? args.state.workbook.cellStore.stringIds[cellIndex] : undefined
                if (hasExactLookupDependents || hasAggregateDependents) {
                  const exactLookupRequest = withOptionalLookupStringIds({
                    sheetName,
                    row: mutation.row,
                    col: mutation.col,
                    oldValue: prior.value,
                    newValue,
                    oldStringId: prior.stringId,
                    newStringId,
                  })
                  if (hasExactLookupDependents) {
                    args.noteExactLookupLiteralWrite(exactLookupRequest)
                    formulaChangedCount = markAffectedExactLookupDependents(exactLookupRequest, formulaChangedCount)
                  }
                  if (hasAggregateDependents) {
                    args.noteAggregateLiteralWrite({
                      sheetName: exactLookupRequest.sheetName,
                      row: exactLookupRequest.row,
                      col: exactLookupRequest.col,
                      oldValue: exactLookupRequest.oldValue,
                      newValue: exactLookupRequest.newValue,
                    })
                    formulaChangedCount = markAffectedDirectRangeDependents(
                      exactLookupRequest,
                      formulaChangedCount,
                      postRecalcDirectFormulaIndices,
                    )
                  }
                }
                if (hasSortedLookupDependents) {
                  const sortedLookupRequest = withOptionalLookupStringIds({
                    sheetName,
                    row: mutation.row,
                    col: mutation.col,
                    oldValue: prior.value,
                    newValue,
                    oldStringId: prior.stringId,
                    newStringId,
                  })
                  args.noteSortedLookupLiteralWrite(sortedLookupRequest)
                  formulaChangedCount = markAffectedApproximateLookupDependents(sortedLookupRequest, formulaChangedCount)
                }
              }
              args.state.workbook.cellStore.flags[cellIndex] =
                (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
                ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
              if (!isRestore && mutation.value === null) {
                pruneCellIfOrphaned(cellIndex)
              }
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
              if (trackExplicitChanges) {
                explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              }
              if (!isRestore && args.state.trackReplicaVersions) {
                setCellEntityVersion(sheetName, formatAddress(mutation.row, mutation.col), order!)
              }
              break
            }
            case 'setCellFormula': {
              const sheetName = resolveSheetName(sheetId)
              const { hasExactLookupDependents, hasSortedLookupDependents } = resolveTrackedColumnDependencyFlags(sheetId, mutation.col)
              if (hasExactLookupDependents) {
                args.invalidateExactLookupColumn({ sheetName, col: mutation.col })
              }
              if (hasSortedLookupDependents) {
                args.invalidateSortedLookupColumn({ sheetName, col: mutation.col })
              }
              if (!isRestore && existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              }
              const cellIndex = args.state.workbook.ensureCellAt(sheetId, mutation.row, mutation.col).cellIndex
              if (!isRestore) {
                changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
              }
              const compileStarted = isRestore ? 0 : performance.now()
              try {
                const changedTopology = args.bindFormula(cellIndex, sheetName, mutation.formula)
                args.invalidateAggregateColumn({ sheetName, col: mutation.col })
                if (!isRestore) {
                  compileMs += performance.now() - compileStarted
                }
                clearTrackedColumnDependencyFlagCache()
                changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
                formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
                topologyChanged = topologyChanged || changedTopology
                const aggregateDependents = collectAffectedDirectRangeDependents({
                  sheetName,
                  row: mutation.row,
                  col: mutation.col,
                }).filter((candidate) => candidate !== cellIndex)
                if (aggregateDependents.length > 0) {
                  formulaChangedCount = args.rebindFormulaCells(aggregateDependents, formulaChangedCount)
                  for (let index = 0; index < aggregateDependents.length; index += 1) {
                    postRecalcDirectFormulaIndices.add(aggregateDependents[index]!)
                    formulaChangedCount = args.markFormulaChanged(aggregateDependents[index]!, formulaChangedCount)
                    changedInputCount = args.markInputChanged(aggregateDependents[index]!, changedInputCount)
                  }
                  topologyChanged = true
                }
              } catch {
                if (!isRestore) {
                  compileMs += performance.now() - compileStarted
                }
                const removedFormula = args.removeFormula(cellIndex)
                topologyChanged = removedFormula || topologyChanged
                clearTrackedColumnDependencyFlagCache()
                args.setInvalidFormulaValue(cellIndex)
                changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
              }
              if (trackExplicitChanges) {
                explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              }
              if (!isRestore && args.state.trackReplicaVersions) {
                setCellEntityVersion(sheetName, formatAddress(mutation.row, mutation.col), order!)
              }
              break
            }
            case 'clearCell': {
              const { hasExactLookupDependents, hasSortedLookupDependents, hasAggregateDependents, needsLookupValueRead } =
                resolveTrackedColumnDependencyFlags(sheetId, mutation.col)
              const prior = needsLookupValueRead ? readCellValueForLookup(existingIndex) : { value: emptyValue(), stringId: undefined }
              if (existingIndex !== undefined && isClearCellNoOp(existingIndex)) {
                break
              }
              if (existingIndex !== undefined && canFastPathLiteralOverwrite(existingIndex)) {
                args.state.workbook.cellStore.setValue(existingIndex, emptyValue())
                args.state.workbook.notifyCellValueWritten(existingIndex)
                if (needsLookupValueRead) {
                  const sheetName = resolveSheetName(sheetId)
                  if (hasExactLookupDependents || hasAggregateDependents) {
                    const exactLookupRequest = withOptionalLookupStringIds({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                      oldValue: prior.value,
                      newValue: emptyValue(),
                      oldStringId: prior.stringId,
                      newStringId: undefined,
                    })
                    if (hasExactLookupDependents) {
                      args.noteExactLookupLiteralWrite(exactLookupRequest)
                      formulaChangedCount = markAffectedExactLookupDependents(exactLookupRequest, formulaChangedCount)
                    }
                    if (hasAggregateDependents) {
                      args.noteAggregateLiteralWrite({
                        sheetName: exactLookupRequest.sheetName,
                        row: exactLookupRequest.row,
                        col: exactLookupRequest.col,
                        oldValue: exactLookupRequest.oldValue,
                        newValue: exactLookupRequest.newValue,
                      })
                      formulaChangedCount = markAffectedDirectRangeDependents(
                        exactLookupRequest,
                        formulaChangedCount,
                        postRecalcDirectFormulaIndices,
                      )
                    }
                  }
                  if (hasSortedLookupDependents) {
                    const sortedLookupRequest = withOptionalLookupStringIds({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                      oldValue: prior.value,
                      newValue: emptyValue(),
                      oldStringId: prior.stringId,
                      newStringId: undefined,
                    })
                    args.noteSortedLookupLiteralWrite(sortedLookupRequest)
                    formulaChangedCount = markAffectedApproximateLookupDependents(sortedLookupRequest, formulaChangedCount)
                  }
                }
                changedInputCount = args.markInputChanged(existingIndex, changedInputCount)
                if (trackExplicitChanges) {
                  explicitChangedCount = args.markExplicitChanged(existingIndex, explicitChangedCount)
                }
                if (!isRestore && args.state.trackReplicaVersions) {
                  setCellEntityVersion(resolveSheetName(sheetId), formatAddress(mutation.row, mutation.col), order!)
                }
                break
              }
              if (existingIndex === undefined) {
                if (!isRestore && args.state.trackReplicaVersions) {
                  setCellEntityVersion(resolveSheetName(sheetId), formatAddress(mutation.row, mutation.col), order!)
                }
                break
              }
              changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(existingIndex), changedInputCount)
              const removedFormula = args.removeFormula(existingIndex)
              topologyChanged = removedFormula || topologyChanged
              if (removedFormula) {
                args.invalidateAggregateColumn({ sheetName: resolveSheetName(sheetId), col: mutation.col })
              }
              if (removedFormula) {
                clearTrackedColumnDependencyFlagCache()
              }
              args.state.workbook.cellStore.setValue(existingIndex, emptyValue())
              args.state.workbook.notifyCellValueWritten(existingIndex)
              if (needsLookupValueRead) {
                const sheetName = resolveSheetName(sheetId)
                if (hasExactLookupDependents || hasAggregateDependents) {
                  const exactLookupRequest = withOptionalLookupStringIds({
                    sheetName,
                    row: mutation.row,
                    col: mutation.col,
                    oldValue: prior.value,
                    newValue: emptyValue(),
                    oldStringId: prior.stringId,
                    newStringId: undefined,
                  })
                  if (hasExactLookupDependents) {
                    args.noteExactLookupLiteralWrite(exactLookupRequest)
                    formulaChangedCount = markAffectedExactLookupDependents(exactLookupRequest, formulaChangedCount)
                  }
                  if (hasAggregateDependents) {
                    args.noteAggregateLiteralWrite({
                      sheetName: exactLookupRequest.sheetName,
                      row: exactLookupRequest.row,
                      col: exactLookupRequest.col,
                      oldValue: exactLookupRequest.oldValue,
                      newValue: exactLookupRequest.newValue,
                    })
                    formulaChangedCount = markAffectedDirectRangeDependents(
                      exactLookupRequest,
                      formulaChangedCount,
                      postRecalcDirectFormulaIndices,
                    )
                  }
                }
                if (hasSortedLookupDependents) {
                  const sortedLookupRequest = withOptionalLookupStringIds({
                    sheetName,
                    row: mutation.row,
                    col: mutation.col,
                    oldValue: prior.value,
                    newValue: emptyValue(),
                    oldStringId: prior.stringId,
                    newStringId: undefined,
                  })
                  args.noteSortedLookupLiteralWrite(sortedLookupRequest)
                  formulaChangedCount = markAffectedApproximateLookupDependents(sortedLookupRequest, formulaChangedCount)
                }
              }
              args.state.workbook.cellStore.flags[existingIndex] =
                (args.state.workbook.cellStore.flags[existingIndex] ?? 0) &
                ~(
                  CellFlags.AuthoredBlank |
                  CellFlags.HasFormula |
                  CellFlags.JsOnly |
                  CellFlags.InCycle |
                  CellFlags.SpillChild |
                  CellFlags.PivotOutput
                )
              normalizeHistoryDependencyPlaceholder(existingIndex, source)
              if (!isRestore) {
                pruneCellIfOrphaned(existingIndex)
              }
              changedInputCount = args.markInputChanged(existingIndex, changedInputCount)
              if (trackExplicitChanges) {
                explicitChangedCount = args.markExplicitChanged(existingIndex, explicitChangedCount)
              }
              if (!isRestore && args.state.trackReplicaVersions) {
                setCellEntityVersion(resolveSheetName(sheetId), formatAddress(mutation.row, mutation.col), order!)
              }
              break
            }
            default:
              assertNever(mutation)
          }
        })
      })

      const reboundCount = formulaChangedCount
      formulaChangedCount = args.syncDynamicRanges(formulaChangedCount)
      topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
      args.flushWasmProgramSync()
    }

    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    if (refs.length === 0) {
      if (!isRestore && batch) {
        emitBatch(batch)
      }
      return
    }

    if (topologyChanged) {
      const repaired =
        !hadCycleMembersBefore &&
        formulaChangedCount > 0 &&
        args.repairTopoRanks(args.getChangedFormulaBuffer().subarray(0, formulaChangedCount))
      if (!repaired) {
        args.rebuildTopoRanks()
        args.detectCycles()
        changedInputCount = markCycleMemberInputsChanged(changedInputCount)
      }
    }
    args.prepareRegionQueryIndices()
    const hasActiveFormulas = args.state.formulas.size > 0
    const hasActivePivots = args.state.workbook.listPivots().length > 0
    let recalculated: U32 = new Uint32Array()
    if (hasActiveFormulas || hasActivePivots) {
      formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount)
      const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
      recalculated = args.recalculate(args.composeMutationRoots(changedInputCount, formulaChangedCount), changedInputArray)
      if (postRecalcDirectFormulaIndices.size > 0) {
        const postRecalcChanged: number[] = []
        postRecalcDirectFormulaIndices.forEach((cellIndex) => {
          if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
            return
          }
          const changedCellIndices = args.evaluateDirectFormula(cellIndex)
          postRecalcChanged.push(cellIndex)
          if (changedCellIndices) {
            for (let index = 0; index < changedCellIndices.length; index += 1) {
              postRecalcChanged.push(changedCellIndices[index]!)
            }
          }
        })
        recalculated = mergeChangedCellIndices(recalculated, postRecalcChanged)
      }
      recalculated = args.reconcilePivotOutputs(recalculated, false)
    }
    const changed: U32 = isRestore || !requiresChangedSet ? new Uint32Array() : args.composeEventChanges(recalculated, explicitChangedCount)
    const lastMetrics = {
      ...args.state.getLastMetrics(),
      batchId: args.state.getLastMetrics().batchId + 1,
      changedInputCount: changedInputCount + formulaChangedCount,
      compileMs,
    }
    args.state.setLastMetrics(lastMetrics)
    const invalidation = isRestore ? 'full' : 'cells'
    if (hasGeneralEventListeners || hasWatchedCellListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation,
        changedCellIndices: changed,
        changedCells: hasGeneralEventListeners ? args.captureChangedCells(changed) : [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      }
      if (isRestore) {
        args.state.events.emitAllWatched(event)
        if (!hasTrackedEventListeners) {
          return
        }
      } else {
        args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
      }
    } else if (isRestore && !hasTrackedEventListeners) {
      return
    }
    if (hasTrackedEventListeners) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation,
        changedCellIndices: changed,
        patches: args.captureChangedPatches(changed),
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      })
    }
    if (batch) {
      void args.state.getSyncClientConnection()?.send(batch)
      emitBatch(batch)
    }
  }

  return {
    applyBatch(batch, source, potentialNewCells, preparedCellAddressesByOpIndex) {
      return Effect.try({
        try: () => {
          applyBatchNow(batch, source, potentialNewCells, preparedCellAddressesByOpIndex)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to apply ${source} batch`, cause),
            cause,
          }),
      })
    },
    applyCellMutationsAt(refs, batch, source, potentialNewCells) {
      return Effect.try({
        try: () => {
          applyCellMutationsAtNow(refs, batch, source, potentialNewCells)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to apply ${source} cell mutations`, cause),
            cause,
          }),
      })
    },
    applyDerivedOp(op) {
      return Effect.try({
        try: () => applyDerivedOpNow(op),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to apply derived operation ${op.kind}`, cause),
            cause,
          }),
      })
    },
  }
}
