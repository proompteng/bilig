import { Effect } from 'effect'
import { compileCriteriaMatcher, formatAddress, matchesCompiledCriteria, parseCellAddress } from '@bilig/formula'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import {
  ErrorCode,
  FormulaMode,
  MAX_COLS,
  ValueTag,
  type CellRangeRef,
  type CellValue,
  type EngineChangedCell,
  type EngineEvent,
  type LiteralInput,
  type SelectionState,
} from '@bilig/protocol'
import type { EdgeSlice } from '../../edge-arena.js'
import type {
  EngineCellMutationRef,
  EngineExistingNumericCellMutationRef,
  EngineExistingNumericCellMutationResult,
} from '../../cell-mutations-at.js'
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
import { makeCellKey, normalizeDefinedName, pivotKey, type SheetRecord, type WorkbookPivotRecord } from '../../workbook-store.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import type {
  EngineRuntimeState,
  PreparedCellAddress,
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectLookupDescriptor,
  RuntimeDirectScalarDescriptor,
  RuntimeDirectScalarOperand,
  U32,
} from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'
import type { EnginePatch } from '../../patches/patch-types.js'
import { addEngineCounter } from '../../perf/engine-counters.js'

type MutationSource = 'local' | 'remote' | 'restore' | 'undo' | 'redo'
const GENERAL_CHANGED_CELL_PAYLOAD_LIMIT = 512
const DIRECT_RANGE_POST_RECALC_LIMIT = 16_384
const DIRECT_SCALAR_DELTA_CLOSURE_LIMIT = 4_096
const EMPTY_CHANGED_CELLS = new Uint32Array(0)
const TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY = '__biligTrackedPhysicalSheetId'
const TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY = '__biligTrackedPhysicalSortedSliceSplit'
const ENGINE_OPERATION_TEST_HOOKS_ENABLED = process.env['NODE_ENV'] === 'test'
const ROW_PAIR_LEFT_PLUS_RIGHT = 1
const ROW_PAIR_LEFT_MINUS_RIGHT = 2
const ROW_PAIR_RIGHT_MINUS_LEFT = 3
const ROW_PAIR_LEFT_TIMES_RIGHT = 4
const ROW_PAIR_LEFT_DIV_RIGHT = 5
const ROW_PAIR_RIGHT_DIV_LEFT = 6

interface DirectFormulaMetricCounts {
  wasmFormulaCount: number
  jsFormulaCount: number
}

type StructuralAxisOp = Extract<
  EngineOp,
  {
    kind: 'insertRows' | 'deleteRows' | 'moveRows' | 'insertColumns' | 'deleteColumns' | 'moveColumns'
  }
>

type DerivedOp = Extract<EngineOp, { kind: 'upsertSpillRange' | 'deleteSpillRange' | 'upsertPivotTable' | 'deletePivotTable' }>
type DirectScalarCurrentOperand = { kind: 'number'; value: number } | { kind: 'error'; code: ErrorCode }
type UniformNumericDirectLookup = Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' | 'approximate-uniform-numeric' }>
type UniformLookupTailPatchTarget = Extract<
  RuntimeDirectLookupDescriptor,
  { kind: 'exact-uniform-numeric' | 'approximate-uniform-numeric' }
>

interface LookupNumericColumnWritePlan {
  readonly handled: boolean
  readonly tailPatchTarget?: UniformLookupTailPatchTarget
}

interface ExactLookupImpactEntry {
  readonly formulaCellIndex: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly operandKey: string | undefined
}

interface ExactLookupImpactCache {
  readonly entries: readonly ExactLookupImpactEntry[]
  readonly operandKeys: ReadonlySet<string>
}

type ExactLookupImpactCaches = Map<string, ExactLookupImpactCache>

function directLookupVersionMatches(lookupSheet: SheetRecord | undefined, lookup: UniformNumericDirectLookup): boolean {
  if ((lookupSheet?.structureVersion ?? 0) !== lookup.structureVersion) {
    return false
  }
  const currentColumnVersion = lookupSheet?.columnVersions[lookup.col] ?? 0
  return currentColumnVersion === lookup.columnVersion || currentColumnVersion === lookup.tailPatch?.columnVersion
}

function reverseUint32Array(values: Uint32Array): Uint32Array {
  const reversed = new Uint32Array(values.length)
  for (let index = 0; index < values.length; index += 1) {
    reversed[index] = values[values.length - 1 - index]!
  }
  return reversed
}

class PendingNumericCellValues {
  private readonly values: number[] = []
  private readonly assigned: boolean[] = []

  get(cellIndex: number): number | undefined {
    return this.assigned[cellIndex] === true ? this.values[cellIndex] : undefined
  }

  has(cellIndex: number): boolean {
    return this.assigned[cellIndex] === true
  }

  set(cellIndex: number, value: number): void {
    this.assigned[cellIndex] = true
    this.values[cellIndex] = value
  }
}

class DirectFormulaIndexCollection {
  private readonly cellIndices: number[] = []
  private sharedCellIndices: readonly number[] | U32 | undefined
  private indexByCell: Map<number, number> | undefined
  private deltas: number[] | undefined
  private deltaAssigned: boolean[] | undefined
  private scalarDeltaAssigned: boolean[] | undefined
  private currentResults: DirectScalarCurrentOperand[] | undefined
  private currentResultAssigned: boolean[] | undefined
  private directFormulaCoveredInputCellIndices: number[] | undefined
  private directRangeCoveredInputCellIndices: number[] | undefined
  private deltaCount = 0
  private scalarDeltaCount = 0
  private constantDelta: number | undefined
  private validatedScalarDeltaSize = -1

  get size(): number {
    return this.sharedCellIndices?.length ?? this.cellIndices.length
  }

  add(cellIndex: number): void {
    this.validatedScalarDeltaSize = -1
    this.materializeConstantDeltas()
    if (this.indexByCell) {
      if (this.indexByCell.has(cellIndex)) {
        return
      }
      this.materializeSharedCellIndices()
      this.indexByCell.set(cellIndex, this.cellIndices.length)
      this.cellIndices.push(cellIndex)
      return
    }
    for (let index = 0; index < this.size; index += 1) {
      if (this.getCellIndexAt(index) === cellIndex) {
        return
      }
    }
    this.materializeSharedCellIndices()
    this.cellIndices.push(cellIndex)
    if (this.size > 16) {
      this.materializeIndexByCell()
    }
  }

  has(cellIndex: number): boolean {
    if (!this.indexByCell && this.size > 16) {
      this.materializeIndexByCell()
    }
    if (this.indexByCell) {
      return this.indexByCell.has(cellIndex)
    }
    for (let index = 0; index < this.size; index += 1) {
      if (this.getCellIndexAt(index) === cellIndex) {
        return true
      }
    }
    return false
  }

  addDelta(cellIndex: number, delta: number): void {
    this.addDeltaWithKind(cellIndex, delta, undefined)
  }

  addScalarDelta(cellIndex: number, delta: number): void {
    this.addDeltaWithKind(cellIndex, delta, 'scalar')
  }

  private addDeltaWithKind(cellIndex: number, delta: number, kind: 'scalar' | undefined): void {
    this.validatedScalarDeltaSize = -1
    if (this.constantDelta !== undefined) {
      const existingIndex = this.findIndex(cellIndex)
      if (existingIndex === -1 && Object.is(this.constantDelta, delta)) {
        this.materializeSharedCellIndices()
        if (this.indexByCell) {
          this.indexByCell.set(cellIndex, this.cellIndices.length)
        }
        this.cellIndices.push(cellIndex)
        this.deltaCount += 1
        if (kind === 'scalar') {
          this.scalarDeltaCount += 1
        }
        return
      }
    }
    this.materializeConstantDeltas()
    const index = this.ensureIndex(cellIndex)
    this.deltas ??= []
    this.deltaAssigned ??= []
    this.scalarDeltaAssigned ??= []
    if (!this.deltaAssigned[index]) {
      this.deltaAssigned[index] = true
      this.deltaCount += 1
      this.deltas[index] = delta
      if (kind === 'scalar') {
        this.scalarDeltaAssigned[index] = true
        this.scalarDeltaCount += 1
      }
      return
    }
    this.deltas[index] = (this.deltas[index] ?? 0) + delta
    if (kind !== 'scalar' && this.scalarDeltaAssigned[index]) {
      this.scalarDeltaAssigned[index] = false
      this.scalarDeltaCount -= 1
    }
  }

  appendDeltas(cellIndices: readonly number[] | U32, deltas: readonly number[], kind?: 'scalar'): void {
    if (cellIndices.length === 0) {
      return
    }
    this.validatedScalarDeltaSize = -1
    if (this.size !== 0) {
      if (cellIndices.length > 16 || this.size > 16) {
        this.appendPreparedDeltas(cellIndices, deltas, kind)
        return
      }
      for (let index = 0; index < cellIndices.length; index += 1) {
        this.addDeltaWithKind(cellIndices[index]!, deltas[index]!, kind)
      }
      return
    }
    this.sharedCellIndices = cellIndices
    this.deltas = []
    this.deltaAssigned = []
    this.scalarDeltaAssigned = kind === 'scalar' ? [] : undefined
    for (let index = 0; index < cellIndices.length; index += 1) {
      this.deltas[index] = deltas[index]!
      this.deltaAssigned[index] = true
      if (kind === 'scalar') {
        this.scalarDeltaAssigned![index] = true
      }
    }
    this.deltaCount = cellIndices.length
    this.scalarDeltaCount = kind === 'scalar' ? cellIndices.length : 0
  }

  appendConstantDelta(cellIndices: readonly number[] | U32, delta: number, kind?: 'scalar'): void {
    if (cellIndices.length === 0) {
      return
    }
    this.validatedScalarDeltaSize = -1
    if (this.size !== 0) {
      if (cellIndices.length > 16 || this.size > 16) {
        this.appendPreparedConstantDelta(cellIndices, delta, kind)
        return
      }
      for (let index = 0; index < cellIndices.length; index += 1) {
        this.addDeltaWithKind(cellIndices[index]!, delta, kind)
      }
      return
    }
    this.sharedCellIndices = cellIndices
    this.constantDelta = delta
    this.deltaCount = cellIndices.length
    this.scalarDeltaCount = kind === 'scalar' ? cellIndices.length : 0
  }

  hasDelta(cellIndex: number): boolean {
    const index = this.findIndex(cellIndex)
    if (index !== -1 && this.constantDelta !== undefined) {
      return true
    }
    return index !== -1 && this.deltaAssigned?.[index] === true
  }

  getDelta(cellIndex: number): number | undefined {
    const index = this.findIndex(cellIndex)
    if (index !== -1 && this.constantDelta !== undefined) {
      return this.constantDelta
    }
    if (index === -1 || this.deltaAssigned?.[index] !== true) {
      return undefined
    }
    return this.deltas?.[index]
  }

  getDeltaAt(index: number): number | undefined {
    if (this.constantDelta !== undefined && index >= 0 && index < this.size) {
      return this.constantDelta
    }
    if (this.deltaAssigned?.[index] !== true) {
      return undefined
    }
    return this.deltas?.[index]
  }

  addCurrentResult(cellIndex: number, result: DirectScalarCurrentOperand): void {
    this.validatedScalarDeltaSize = -1
    this.materializeConstantDeltas()
    const index = this.ensureIndex(cellIndex)
    this.currentResults ??= []
    this.currentResultAssigned ??= []
    this.currentResultAssigned[index] = true
    this.currentResults[index] = result
  }

  getCurrentResult(cellIndex: number): DirectScalarCurrentOperand | undefined {
    const index = this.findIndex(cellIndex)
    if (index === -1 || this.currentResultAssigned?.[index] !== true) {
      return undefined
    }
    return this.currentResults?.[index]
  }

  getCurrentResultAt(index: number): DirectScalarCurrentOperand | undefined {
    if (this.currentResultAssigned?.[index] !== true) {
      return undefined
    }
    return this.currentResults?.[index]
  }

  hasCompleteDeltas(): boolean {
    return this.size > 0 && this.deltaCount === this.size
  }

  getConstantScalarDelta(): number | undefined {
    return this.constantDelta !== undefined &&
      this.scalarDeltaCount === this.size &&
      this.deltaCount === this.size &&
      this.currentResultAssigned === undefined
      ? this.constantDelta
      : undefined
  }

  hasCompleteScalarDeltas(): boolean {
    return this.size > 0 && this.deltaCount === this.size && this.scalarDeltaCount === this.size && this.currentResultAssigned === undefined
  }

  markScalarDeltaCellsValidated(): void {
    this.validatedScalarDeltaSize = this.size
  }

  hasValidatedScalarDeltaCells(): boolean {
    return this.validatedScalarDeltaSize === this.size && this.hasCompleteScalarDeltas()
  }

  getScalarDeltaAt(index: number): number | undefined {
    if (this.constantDelta !== undefined && this.scalarDeltaCount === this.size && index >= 0 && index < this.size) {
      return this.constantDelta
    }
    if (this.scalarDeltaAssigned?.[index] !== true) {
      return undefined
    }
    return this.deltas?.[index]
  }

  getCellIndexAt(index: number): number {
    return (this.sharedCellIndices ?? this.cellIndices)[index]!
  }

  markDirectRangeInputCovered(cellIndex: number): void {
    const covered = (this.directRangeCoveredInputCellIndices ??= [])
    for (let index = 0; index < covered.length; index += 1) {
      if (covered[index] === cellIndex) {
        return
      }
    }
    covered.push(cellIndex)
  }

  hasCoveredDirectRangeInput(cellIndex: number): boolean {
    const covered = this.directRangeCoveredInputCellIndices
    if (!covered) {
      return false
    }
    for (let index = 0; index < covered.length; index += 1) {
      if (covered[index] === cellIndex) {
        return true
      }
    }
    return false
  }

  markDirectFormulaInputCovered(cellIndex: number): void {
    const covered = (this.directFormulaCoveredInputCellIndices ??= [])
    for (let index = 0; index < covered.length; index += 1) {
      if (covered[index] === cellIndex) {
        return
      }
    }
    covered.push(cellIndex)
  }

  hasCoveredDirectFormulaInput(cellIndex: number): boolean {
    const covered = this.directFormulaCoveredInputCellIndices
    if (!covered) {
      return false
    }
    for (let index = 0; index < covered.length; index += 1) {
      if (covered[index] === cellIndex) {
        return true
      }
    }
    return false
  }

  forEach(fn: (cellIndex: number) => void): void {
    for (let index = 0; index < this.size; index += 1) {
      fn(this.getCellIndexAt(index))
    }
  }

  forEachIndexed(fn: (cellIndex: number, index: number) => void): void {
    for (let index = 0; index < this.size; index += 1) {
      fn(this.getCellIndexAt(index), index)
    }
  }

  private findIndex(cellIndex: number): number {
    if (!this.indexByCell && this.size > 16) {
      this.materializeIndexByCell()
    }
    const mappedIndex = this.indexByCell?.get(cellIndex)
    if (mappedIndex !== undefined) {
      return mappedIndex
    }
    for (let index = 0; index < this.size; index += 1) {
      if (this.getCellIndexAt(index) === cellIndex) {
        return index
      }
    }
    return -1
  }

  private ensureIndex(cellIndex: number): number {
    if (!this.indexByCell && this.size > 16) {
      this.materializeIndexByCell()
    }
    if (this.indexByCell) {
      const mappedIndex = this.indexByCell.get(cellIndex)
      if (mappedIndex !== undefined) {
        return mappedIndex
      }
      this.materializeSharedCellIndices()
      const index = this.cellIndices.length
      this.indexByCell.set(cellIndex, index)
      this.cellIndices.push(cellIndex)
      return index
    }
    for (let index = 0; index < this.size; index += 1) {
      if (this.getCellIndexAt(index) === cellIndex) {
        return index
      }
    }
    this.materializeSharedCellIndices()
    const index = this.cellIndices.length
    this.cellIndices.push(cellIndex)
    if (this.size > 16) {
      this.materializeIndexByCell()
    }
    return index
  }

  private materializeIndexByCell(): void {
    this.indexByCell = new Map()
    for (let index = 0; index < this.size; index += 1) {
      this.indexByCell.set(this.getCellIndexAt(index), index)
    }
  }

  private appendPreparedDeltas(cellIndices: readonly number[] | U32, deltas: readonly number[], kind: 'scalar' | undefined): void {
    this.prepareForBulkDeltaAppend()
    for (let index = 0; index < cellIndices.length; index += 1) {
      this.addPreparedDelta(cellIndices[index]!, deltas[index]!, kind)
    }
  }

  private appendPreparedConstantDelta(cellIndices: readonly number[] | U32, delta: number, kind: 'scalar' | undefined): void {
    this.prepareForBulkDeltaAppend()
    for (let index = 0; index < cellIndices.length; index += 1) {
      this.addPreparedDelta(cellIndices[index]!, delta, kind)
    }
  }

  private prepareForBulkDeltaAppend(): void {
    this.materializeConstantDeltas()
    this.materializeSharedCellIndices()
    this.materializeIndexByCell()
    this.deltas ??= []
    this.deltaAssigned ??= []
    this.scalarDeltaAssigned ??= []
  }

  private addPreparedDelta(cellIndex: number, delta: number, kind: 'scalar' | undefined): void {
    const mappedIndex = this.indexByCell?.get(cellIndex)
    if (mappedIndex === undefined) {
      const index = this.cellIndices.length
      this.indexByCell!.set(cellIndex, index)
      this.cellIndices.push(cellIndex)
      this.deltas![index] = delta
      this.deltaAssigned![index] = true
      this.deltaCount += 1
      if (kind === 'scalar') {
        this.scalarDeltaAssigned![index] = true
        this.scalarDeltaCount += 1
      }
      return
    }
    if (!this.deltaAssigned![mappedIndex]) {
      this.deltaAssigned![mappedIndex] = true
      this.deltaCount += 1
      this.deltas![mappedIndex] = delta
      if (kind === 'scalar') {
        this.scalarDeltaAssigned![mappedIndex] = true
        this.scalarDeltaCount += 1
      }
      return
    }
    this.deltas![mappedIndex] = (this.deltas![mappedIndex] ?? 0) + delta
    if (kind !== 'scalar' && this.scalarDeltaAssigned![mappedIndex]) {
      this.scalarDeltaAssigned![mappedIndex] = false
      this.scalarDeltaCount -= 1
    }
  }

  private materializeSharedCellIndices(): void {
    const sharedCellIndices = this.sharedCellIndices
    if (!sharedCellIndices) {
      return
    }
    this.cellIndices.length = sharedCellIndices.length
    for (let index = 0; index < sharedCellIndices.length; index += 1) {
      this.cellIndices[index] = sharedCellIndices[index]!
    }
    this.sharedCellIndices = undefined
  }

  private materializeConstantDeltas(): void {
    if (this.constantDelta === undefined) {
      return
    }
    const delta = this.constantDelta
    const scalarDeltaCount = this.scalarDeltaCount
    const size = this.size
    this.constantDelta = undefined
    this.deltas = []
    this.deltaAssigned = []
    this.scalarDeltaAssigned = scalarDeltaCount === size ? [] : undefined
    for (let index = 0; index < size; index += 1) {
      this.deltas[index] = delta
      this.deltaAssigned[index] = true
      if (this.scalarDeltaAssigned) {
        this.scalarDeltaAssigned[index] = true
      }
    }
    this.deltaCount = size
    this.scalarDeltaCount = this.scalarDeltaAssigned ? size : 0
  }
}

export interface EngineOperationService {
  readonly __testHooks: Record<string, unknown>
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
  readonly applyCellMutationsAtNow: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => void
  readonly applyExistingNumericCellMutationAtNow: (
    request: EngineExistingNumericCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
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

function directAggregateNumericContribution(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
    case ValueTag.String:
      return 0
    case ValueTag.Error:
      return undefined
  }
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
  inputCellIndex?: number
}): {
  sheetName: string
  row: number
  col: number
  oldValue: CellValue
  newValue: CellValue
  oldStringId?: number
  newStringId?: number
  inputCellIndex?: number
} {
  return {
    sheetName: request.sheetName,
    row: request.row,
    col: request.col,
    oldValue: request.oldValue,
    newValue: request.newValue,
    ...(request.oldStringId === undefined ? {} : { oldStringId: request.oldStringId }),
    ...(request.newStringId === undefined ? {} : { newStringId: request.newStringId }),
    ...(request.inputCellIndex === undefined ? {} : { inputCellIndex: request.inputCellIndex }),
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

function normalizeExactNumericValue(value: CellValue): number | undefined {
  return value.tag === ValueTag.Number ? (Object.is(value.value, -0) ? 0 : value.value) : undefined
}

function sameExactNumericValue(left: number, right: number): boolean {
  return left === right || Object.is(left, right)
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

function exactLookupLiteralNumericValue(value: unknown): number | undefined {
  return typeof value === 'number' ? (Object.is(value, -0) ? 0 : value) : undefined
}

function canSkipUniformApproximateNumericTailWrite(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>,
  row: number,
  operandNumeric: number,
  oldNumeric: number,
  newNumeric: number,
): boolean {
  if (directLookup.matchMode === 1 && directLookup.step > 0) {
    return row === directLookup.rowEnd && oldNumeric > operandNumeric && newNumeric > operandNumeric && newNumeric >= oldNumeric
  }
  if (directLookup.matchMode === -1 && directLookup.step < 0) {
    return row === directLookup.rowEnd && oldNumeric < operandNumeric && newNumeric < operandNumeric && newNumeric <= oldNumeric
  }
  return false
}

function canSkipUniformApproximateNumericTailWriteFromCurrentResult(
  cellStore: {
    readonly tags: ArrayLike<ValueTag | undefined>
    readonly numbers: ArrayLike<number | undefined>
  },
  formulaCellIndex: number,
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>,
  row: number,
  oldNumeric: number,
  newNumeric: number,
): boolean {
  if (
    !(
      (directLookup.matchMode === 1 && directLookup.step > 0 && row === directLookup.rowEnd && newNumeric >= oldNumeric) ||
      (directLookup.matchMode === -1 && directLookup.step < 0 && row === directLookup.rowEnd && newNumeric <= oldNumeric)
    )
  ) {
    return false
  }
  return (
    (cellStore.tags[formulaCellIndex] ?? ValueTag.Empty) === ValueTag.Number &&
    (cellStore.numbers[formulaCellIndex] ?? 0) < directLookup.length
  )
}

function canSkipUniformExactNumericTailWriteFromCurrentResult(
  cellStore: {
    readonly tags: ArrayLike<ValueTag | undefined>
    readonly numbers: ArrayLike<number | undefined>
  },
  formulaCellIndex: number,
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>,
  row: number,
  oldNumeric: number,
  newNumeric: number,
): boolean {
  if (directLookup.tailPatch !== undefined || row !== directLookup.rowEnd || directLookup.length < 2) {
    return false
  }
  const expectedOldTail = directLookup.start + directLookup.step * (directLookup.length - 1)
  if (!sameExactNumericValue(oldNumeric, expectedOldTail)) {
    return false
  }
  if (!(directLookup.step > 0 ? newNumeric > oldNumeric : newNumeric < oldNumeric)) {
    return false
  }
  return (
    (cellStore.tags[formulaCellIndex] ?? ValueTag.Empty) === ValueTag.Number &&
    (cellStore.numbers[formulaCellIndex] ?? 0) < directLookup.length
  )
}

function directLookupRowBounds(
  directLookup: Extract<
    RuntimeDirectLookupDescriptor,
    { kind: 'exact' | 'exact-uniform-numeric' | 'approximate' | 'approximate-uniform-numeric' }
  >,
): { rowStart: number; rowEnd: number } {
  switch (directLookup.kind) {
    case 'exact':
    case 'approximate':
      return {
        rowStart: directLookup.prepared.rowStart,
        rowEnd: directLookup.prepared.rowEnd,
      }
    case 'exact-uniform-numeric':
    case 'approximate-uniform-numeric':
      return {
        rowStart: directLookup.rowStart,
        rowEnd: directLookup.rowEnd,
      }
  }
  return assertNever(directLookup)
}

function exactUniformLookupCurrentResult(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>,
  lookupValue: number,
): DirectScalarCurrentOperand {
  const numericResult = exactUniformLookupNumericResult(directLookup, lookupValue)
  return numericResult === undefined ? { kind: 'error', code: ErrorCode.NA } : { kind: 'number', value: numericResult }
}

function exactUniformLookupNumericResult(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>,
  lookupValue: number,
): number | undefined {
  const tailPatch = directLookup.tailPatch
  if (tailPatch === undefined && directLookup.step === 1) {
    if (!Number.isInteger(lookupValue)) {
      return undefined
    }
    const position = lookupValue - directLookup.start + 1
    return position >= 1 && position <= directLookup.length ? position : undefined
  }
  if (tailPatch === undefined && directLookup.step === -1) {
    if (!Number.isInteger(lookupValue)) {
      return undefined
    }
    const position = directLookup.start - lookupValue + 1
    return position >= 1 && position <= directLookup.length ? position : undefined
  }
  if (tailPatch !== undefined) {
    if (sameExactNumericValue(lookupValue, tailPatch.newNumeric)) {
      return tailPatch.row - directLookup.rowStart + 1
    }
    if (sameExactNumericValue(lookupValue, tailPatch.oldNumeric)) {
      return undefined
    }
  }
  if (directLookup.step === 1) {
    if (!Number.isInteger(lookupValue)) {
      return undefined
    }
    const position = lookupValue - directLookup.start + 1
    return position >= 1 && position <= directLookup.length ? position : undefined
  }
  if (directLookup.step === -1) {
    if (!Number.isInteger(lookupValue)) {
      return undefined
    }
    const position = directLookup.start - lookupValue + 1
    return position >= 1 && position <= directLookup.length ? position : undefined
  }
  const relative = (lookupValue - directLookup.start) / directLookup.step
  return Number.isInteger(relative) && relative >= 0 && relative < directLookup.length ? relative + 1 : undefined
}

function approximateUniformLookupCurrentResult(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>,
  lookupValue: number,
): DirectScalarCurrentOperand | undefined {
  const numericResult = approximateUniformLookupNumericResult(directLookup, lookupValue)
  if (numericResult !== undefined) {
    return { kind: 'number', value: numericResult }
  }
  if (
    (directLookup.matchMode === 1 && directLookup.step > 0 && lookupValue < directLookup.start) ||
    (directLookup.matchMode === -1 && directLookup.step < 0 && lookupValue > directLookup.start)
  ) {
    return { kind: 'error', code: ErrorCode.NA }
  }
  return undefined
}

function approximateUniformLookupNumericResult(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>,
  lookupValue: number,
): number | undefined {
  const tailPatch = directLookup.tailPatch
  if (tailPatch === undefined && directLookup.matchMode === 1 && directLookup.step === 1) {
    if (lookupValue < directLookup.start) {
      return undefined
    }
    const position = Math.floor(lookupValue - directLookup.start) + 1
    return position >= directLookup.length ? directLookup.length : position
  }
  if (tailPatch === undefined && directLookup.matchMode === -1 && directLookup.step === -1) {
    if (lookupValue > directLookup.start) {
      return undefined
    }
    const position = Math.floor(directLookup.start - lookupValue) + 1
    return position >= directLookup.length ? directLookup.length : position
  }
  const lastValue = directLookup.start + directLookup.step * (directLookup.length - 1)
  if (directLookup.matchMode === 1 && directLookup.step > 0) {
    if (lookupValue < directLookup.start) {
      return undefined
    }
    if (tailPatch !== undefined && tailPatch.row === directLookup.rowEnd && tailPatch.newNumeric > tailPatch.oldNumeric) {
      if (lookupValue >= tailPatch.newNumeric) {
        return directLookup.length
      }
      if (lookupValue >= tailPatch.oldNumeric) {
        return directLookup.length - 1
      }
    }
    if (lookupValue >= lastValue) {
      return directLookup.length
    }
    const position =
      directLookup.step === 1
        ? Math.floor(lookupValue - directLookup.start) + 1
        : Math.floor((lookupValue - directLookup.start) / directLookup.step) + 1
    return Math.min(directLookup.length, Math.max(1, position))
  }
  if (directLookup.matchMode === -1 && directLookup.step < 0) {
    if (lookupValue > directLookup.start) {
      return undefined
    }
    if (tailPatch !== undefined && tailPatch.row === directLookup.rowEnd && tailPatch.newNumeric < tailPatch.oldNumeric) {
      if (lookupValue <= tailPatch.newNumeric) {
        return directLookup.length
      }
      if (lookupValue <= tailPatch.oldNumeric) {
        return directLookup.length - 1
      }
    }
    if (lookupValue <= lastValue) {
      return directLookup.length
    }
    const position =
      directLookup.step === -1
        ? Math.floor(directLookup.start - lookupValue) + 1
        : Math.floor((directLookup.start - lookupValue) / -directLookup.step) + 1
    return Math.min(directLookup.length, Math.max(1, position))
  }
  return undefined
}

function directScalarLiteralNumericValue(value: unknown): number | undefined {
  if (value === null) {
    return 0
  }
  switch (typeof value) {
    case 'number':
      return Object.is(value, -0) ? 0 : value
    case 'boolean':
      return value ? 1 : 0
    case 'string':
    case 'bigint':
    case 'function':
    case 'object':
    case 'symbol':
    case 'undefined':
      return undefined
  }
  return undefined
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

function mergeChangedCellIndices(base: readonly number[] | U32, extras: readonly number[] | U32): U32 {
  if (base.length === 0) {
    return extras instanceof Uint32Array ? extras : Uint32Array.from(extras)
  }
  if (extras.length === 0) {
    return base instanceof Uint32Array ? base : Uint32Array.from(base)
  }
  if (base.length === 1 && extras.length === 1) {
    const baseCellIndex = base[0]!
    const extraCellIndex = extras[0]!
    return baseCellIndex === extraCellIndex ? Uint32Array.of(baseCellIndex) : Uint32Array.of(baseCellIndex, extraCellIndex)
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

function composeSingleDisjointExplicitEventChanges(explicitCellIndex: number, recalculated: U32): U32 {
  if (recalculated.length === 0) {
    return Uint32Array.of(explicitCellIndex)
  }
  const changed = new Uint32Array(recalculated.length + 1)
  changed[0] = explicitCellIndex
  changed.set(recalculated, 1)
  return changed
}

function hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices: DirectFormulaIndexCollection): boolean {
  return postRecalcDirectFormulaIndices.hasCompleteDeltas()
}

function directFormulaChangesAreDisjointFromInputs(
  changedInputArray: U32,
  changedInputCount: number,
  postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
): boolean {
  for (let index = 0; index < changedInputCount; index += 1) {
    if (postRecalcDirectFormulaIndices.has(changedInputArray[index]!)) {
      return false
    }
  }
  return true
}

function countDirectFormulaDeltaSkip(
  formulas: {
    get(cellIndex: number):
      | {
          readonly directAggregate?: unknown
          readonly directCriteria?: unknown
          readonly directScalar?: unknown
        }
      | undefined
  },
  postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  counters: EngineRuntimeState['counters'],
): void {
  let sawAggregate = false
  let sawScalar = false
  postRecalcDirectFormulaIndices.forEach((cellIndex) => {
    const formula = formulas.get(cellIndex)
    sawAggregate ||= formula?.directAggregate !== undefined || formula?.directCriteria !== undefined
    sawScalar ||= formula?.directScalar !== undefined
  })
  if (sawAggregate) {
    addEngineCounter(counters, 'directAggregateDeltaOnlyRecalcSkips')
  }
  if (sawScalar) {
    addEngineCounter(counters, 'directScalarDeltaOnlyRecalcSkips')
  }
}

function canEvaluatePostRecalcDirectFormulasWithoutKernel(
  formulas: {
    get(cellIndex: number):
      | {
          readonly directAggregate?: unknown
          readonly directCriteria?: unknown
          readonly directLookup?: unknown
          readonly directScalar?: unknown
        }
      | undefined
  },
  postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
): boolean {
  if (postRecalcDirectFormulaIndices.size === 0) {
    return false
  }
  let canEvaluate = true
  postRecalcDirectFormulaIndices.forEach((cellIndex) => {
    const formula = formulas.get(cellIndex)
    if (
      formula?.directAggregate === undefined &&
      formula?.directCriteria === undefined &&
      formula?.directLookup === undefined &&
      formula?.directScalar === undefined
    ) {
      canEvaluate = false
    }
  })
  return canEvaluate
}

function lookupImpactCacheKey(sheetId: number, col: number): string {
  return `${sheetId}:${col}`
}

function aggregateColumnDependencyKey(sheetId: number, col: number): number {
  return sheetId * MAX_COLS + col
}

function tagTrustedPhysicalTrackedChanges(changed: U32, sheetId: number, sortedSliceSplit: number): void {
  Reflect.set(changed, TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY, sheetId)
  Reflect.set(changed, TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY, sortedSliceSplit)
}

function makeExistingNumericMutationResult(changedCellIndices: U32, explicitChangedCount: number): EngineExistingNumericCellMutationResult {
  return { changedCellIndices, explicitChangedCount }
}

function makeCompactExistingNumericMutationResult(
  firstChangedCellIndex: number,
  secondChangedCellIndex: number | undefined,
  explicitChangedCount: number,
  secondChangedNumericValue?: number,
  secondChangedPosition?: { readonly row: number; readonly col: number },
): EngineExistingNumericCellMutationResult {
  return secondChangedCellIndex === undefined
    ? { firstChangedCellIndex, changedCellCount: 1, explicitChangedCount }
    : {
        firstChangedCellIndex,
        secondChangedCellIndex,
        changedCellCount: 2,
        explicitChangedCount,
        ...(secondChangedNumericValue === undefined ? {} : { secondChangedNumericValue }),
        ...(secondChangedPosition === undefined
          ? {}
          : { secondChangedRow: secondChangedPosition.row, secondChangedCol: secondChangedPosition.col }),
      }
}

function singleInputAffineDirectScalar(
  directScalar: RuntimeDirectScalarDescriptor,
  inputCellIndex: number,
): { readonly scale: number; readonly offset: number } | null {
  if (directScalar.kind === 'abs') {
    return null
  }
  const leftIsInput = directScalar.left.kind === 'cell' && directScalar.left.cellIndex === inputCellIndex
  const rightIsInput = directScalar.right.kind === 'cell' && directScalar.right.cellIndex === inputCellIndex
  const leftLiteral = directScalar.left.kind === 'literal-number' ? directScalar.left.value : undefined
  const rightLiteral = directScalar.right.kind === 'literal-number' ? directScalar.right.value : undefined
  if (leftIsInput && rightLiteral !== undefined) {
    switch (directScalar.operator) {
      case '+':
        return { scale: 1, offset: rightLiteral }
      case '-':
        return { scale: 1, offset: -rightLiteral }
      case '*':
        return { scale: rightLiteral, offset: 0 }
      case '/':
        return rightLiteral === 0 ? null : { scale: 1 / rightLiteral, offset: 0 }
    }
  }
  if (rightIsInput && leftLiteral !== undefined) {
    switch (directScalar.operator) {
      case '+':
        return { scale: 1, offset: leftLiteral }
      case '-':
        return { scale: -1, offset: leftLiteral }
      case '*':
        return { scale: leftLiteral, offset: 0 }
      case '/':
        return null
    }
  }
  return null
}

function rowPairDirectScalarCode(directScalar: RuntimeDirectScalarDescriptor, leftCellIndex: number, rightCellIndex: number): number {
  if (directScalar.kind === 'abs') {
    return 0
  }
  const leftOperandIsLeft = directScalar.left.kind === 'cell' && directScalar.left.cellIndex === leftCellIndex
  const leftOperandIsRight = directScalar.left.kind === 'cell' && directScalar.left.cellIndex === rightCellIndex
  const rightOperandIsLeft = directScalar.right.kind === 'cell' && directScalar.right.cellIndex === leftCellIndex
  const rightOperandIsRight = directScalar.right.kind === 'cell' && directScalar.right.cellIndex === rightCellIndex
  if (leftOperandIsLeft && rightOperandIsRight) {
    switch (directScalar.operator) {
      case '+':
        return ROW_PAIR_LEFT_PLUS_RIGHT
      case '-':
        return ROW_PAIR_LEFT_MINUS_RIGHT
      case '*':
        return ROW_PAIR_LEFT_TIMES_RIGHT
      case '/':
        return ROW_PAIR_LEFT_DIV_RIGHT
    }
  }
  if (leftOperandIsRight && rightOperandIsLeft) {
    switch (directScalar.operator) {
      case '+':
        return ROW_PAIR_LEFT_PLUS_RIGHT
      case '-':
        return ROW_PAIR_RIGHT_MINUS_LEFT
      case '*':
        return ROW_PAIR_LEFT_TIMES_RIGHT
      case '/':
        return ROW_PAIR_RIGHT_DIV_LEFT
    }
  }
  return 0
}

function evaluateRowPairDirectScalarCode(code: number, leftValue: number, rightValue: number): number | undefined {
  switch (code) {
    case ROW_PAIR_LEFT_PLUS_RIGHT:
      return leftValue + rightValue
    case ROW_PAIR_LEFT_MINUS_RIGHT:
      return leftValue - rightValue
    case ROW_PAIR_RIGHT_MINUS_LEFT:
      return rightValue - leftValue
    case ROW_PAIR_LEFT_TIMES_RIGHT:
      return leftValue * rightValue
    case ROW_PAIR_LEFT_DIV_RIGHT:
      return rightValue === 0 ? undefined : leftValue / rightValue
    case ROW_PAIR_RIGHT_DIV_LEFT:
      return leftValue === 0 ? undefined : rightValue / leftValue
    default:
      return undefined
  }
}

export const operationServiceTestHooks = {
  PendingNumericCellValues,
  DirectFormulaIndexCollection,
  aggregateColumnDependencyKey,
  approximateUniformLookupCurrentResult,
  approximateUniformLookupNumericResult,
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  canSkipUniformApproximateNumericTailWrite,
  canSkipUniformApproximateNumericTailWriteFromCurrentResult,
  canSkipUniformExactNumericTailWriteFromCurrentResult,
  cellRange,
  collectTrackedDependents,
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  directAggregateNumericContribution,
  directCriteriaTouchesPoint,
  directFormulaChangesAreDisjointFromInputs,
  directLookupRowBounds,
  directScalarLiteralNumericValue,
  evaluateRowPairDirectScalarCode,
  exactLookupLiteralNumericValue,
  exactUniformLookupCurrentResult,
  exactUniformLookupNumericResult,
  get ROW_PAIR_LEFT_DIV_RIGHT() {
    return ROW_PAIR_LEFT_DIV_RIGHT
  },
  get ROW_PAIR_LEFT_MINUS_RIGHT() {
    return ROW_PAIR_LEFT_MINUS_RIGHT
  },
  get ROW_PAIR_LEFT_PLUS_RIGHT() {
    return ROW_PAIR_LEFT_PLUS_RIGHT
  },
  get ROW_PAIR_LEFT_TIMES_RIGHT() {
    return ROW_PAIR_LEFT_TIMES_RIGHT
  },
  get ROW_PAIR_RIGHT_DIV_LEFT() {
    return ROW_PAIR_RIGHT_DIV_LEFT
  },
  get ROW_PAIR_RIGHT_MINUS_LEFT() {
    return ROW_PAIR_RIGHT_MINUS_LEFT
  },
  getConstantDirectFormulaDeltas: hasCompleteDirectFormulaDeltas,
  lookupImpactCacheKey,
  makeCompactExistingNumericMutationResult,
  makeExistingNumericMutationResult,
  mergeChangedCellIndices,
  normalizeApproximateNumericValue,
  normalizeApproximateTextValue,
  normalizeExactLookupKey,
  normalizeExactNumericValue,
  rangesIntersect,
  rowPairDirectScalarCode,
  sameExactNumericValue,
  singleInputAffineDirectScalar,
  tagTrustedPhysicalTrackedChanges,
  throwProtectionBlocked,
  withOptionalLookupStringIds,
}

export function createEngineOperationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | 'workbook'
    | 'strings'
    | 'events'
    | 'formulas'
    | 'counters'
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
    readonly reverseCellEdges: Array<EdgeSlice | undefined>
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
  readonly materializeDeferredStructuralFormulaSources: () => void
  readonly removeSheetRuntime: (
    sheetName: string,
    explicitChangedCount: number,
  ) => { changedInputCount: number; formulaChangedCount: number; explicitChangedCount: number }
  readonly applyStructuralAxisOp: (op: StructuralAxisOp) => {
    transaction: StructuralTransaction
    changedCellIndices: number[]
    precomputedChangedInputCellIndices: number[]
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
  readonly hasVolatileFormulas?: () => boolean
  readonly markSpillRootsChanged: (cellIndices: readonly number[], count: number) => number
  readonly markPivotRootsChanged: (cellIndices: readonly number[], count: number) => number
  readonly markExplicitChanged: (cellIndex: number, count: number) => number
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32
  readonly composeEventChanges: (recalculated: U32, explicitChangedCount: number) => U32
  readonly composeDisjointEventChanges: (recalculated: U32, explicitChangedCount: number) => U32
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
  readonly captureChangedPatches: (
    changedCellIndices: readonly number[] | U32,
    request?: {
      invalidation?: 'cells' | 'full'
      invalidatedRanges?: readonly CellRangeRef[]
      invalidatedRows?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
      invalidatedColumns?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    },
  ) => readonly EnginePatch[]
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
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
  readonly evaluateDirectFormula: (cellIndex: number) => readonly number[] | undefined
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => U32
  readonly prepareRegionQueryIndices: () => void
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly getEntityDependents: (entityId: number) => Uint32Array
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly collectFormulaDependents: (entityId: number) => Uint32Array
  readonly hasRegionFormulaSubscriptionsForColumn: (sheetName: string, col: number) => boolean
  readonly hasRegionFormulaSubscriptionsForColumnAt?: (sheetId: number, col: number) => boolean
  readonly collectRegionFormulaDependentsForCell: (sheetName: string, row: number, col: number) => Uint32Array
  readonly collectSingleRegionFormulaDependentForCell: (sheetName: string, row: number, col: number) => number
  readonly collectSingleRegionFormulaDependentForCellAt?: (sheetId: number, row: number, col: number) => number
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
  const singleCellKernelSync = new Uint32Array(1)
  const deferSingleCellKernelSync = (cellIndex: number): void => {
    singleCellKernelSync[0] = cellIndex
    args.deferKernelSync(singleCellKernelSync)
  }
  const makeSingleLiteralSkipMetrics = (): EngineEvent['metrics'] => {
    const previousMetrics = args.state.getLastMetrics()
    return {
      batchId: previousMetrics.batchId + 1,
      changedInputCount: 1,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0,
    }
  }
  const writeLiteralToExistingCell = (cellIndex: number, value: LiteralInput): void => {
    const cellStore = args.state.workbook.cellStore
    const hasSetValueHook = cellStore.onSetValue !== null
    writeLiteralToCellStore(cellStore, cellIndex, value, args.state.strings)
    if (!hasSetValueHook) {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
  }
  const writeNumericLiteralToExistingCell = (cellIndex: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    const flags = cellStore.flags[cellIndex] ?? 0
    const hasSetValueHook = cellStore.onSetValue !== null
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = 0
    cellStore.numbers[cellIndex] = value
    if ((flags & CellFlags.AuthoredBlank) !== 0) {
      cellStore.flags[cellIndex] = flags & ~CellFlags.AuthoredBlank
    }
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    cellStore.onSetValue?.(cellIndex)
    if (!hasSetValueHook) {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
  }
  const writeTrustedExistingNumericLiteralToCell = (cellIndex: number, sheet: SheetRecord, col: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    const flags = cellStore.flags[cellIndex] ?? 0
    cellStore.numbers[cellIndex] = value
    if ((flags & CellFlags.AuthoredBlank) !== 0) {
      cellStore.flags[cellIndex] = flags & ~CellFlags.AuthoredBlank
    }
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    const onSetValue = cellStore.onSetValue
    if (onSetValue) {
      onSetValue(cellIndex)
      return
    }
    sheet.columnVersions[col] = (sheet.columnVersions[col] ?? 0) + 1
  }
  const writeFastPathLiteralToExistingCell = (cellIndex: number, value: LiteralInput): void => {
    if (typeof value === 'number') {
      writeNumericLiteralToExistingCell(cellIndex, value)
      return
    }
    writeLiteralToExistingCell(cellIndex, value)
  }
  const cellsShareVersionColumn = (leftCellIndex: number, rightCellIndex: number): boolean => {
    const workbook = args.state.workbook
    const cellStore = workbook.cellStore
    const leftSheetId = cellStore.sheetIds[leftCellIndex]
    if (leftSheetId === undefined || leftSheetId !== cellStore.sheetIds[rightCellIndex]) {
      return false
    }
    const sheet = workbook.getSheetById(leftSheetId)
    if (!sheet || sheet.structureVersion === 1) {
      const leftCol = cellStore.cols[leftCellIndex]
      return leftCol !== undefined && leftCol === cellStore.cols[rightCellIndex]
    }
    const leftCol = sheet.logical.getCellVisiblePosition(leftCellIndex)?.col ?? cellStore.cols[leftCellIndex]
    const rightCol = sheet.logical.getCellVisiblePosition(rightCellIndex)?.col ?? cellStore.cols[rightCellIndex]
    return leftCol !== undefined && leftCol === rightCol
  }
  const withOptionalColumnVersionBatch = (shouldBatch: boolean, execute: () => void): void => {
    if (shouldBatch) {
      args.state.workbook.withBatchedColumnVersionUpdates(execute)
      return
    }
    execute()
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

  const readApproximateNumericValueForLookup = (cellIndex: number | undefined): number | undefined => {
    if (cellIndex === undefined) {
      return 0
    }
    const cellStore = args.state.workbook.cellStore
    switch ((cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) {
      case ValueTag.Empty:
        return 0
      case ValueTag.Number: {
        const value = cellStore.numbers[cellIndex] ?? 0
        return Object.is(value, -0) ? 0 : value
      }
      case ValueTag.Boolean:
        return (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 1 : 0
      case ValueTag.String:
      case ValueTag.Error:
        return undefined
    }
  }

  const readExactNumericValueForLookup = (cellIndex: number | undefined): number | undefined => {
    if (cellIndex === undefined) {
      return undefined
    }
    const cellStore = args.state.workbook.cellStore
    if (((cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) !== ValueTag.Number) {
      return undefined
    }
    const value = cellStore.numbers[cellIndex] ?? 0
    return Object.is(value, -0) ? 0 : value
  }

  const readCellValueAtForLookup = (sheetName: string, row: number, col: number): { value: CellValue; stringId: number | undefined } => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return { value: emptyValue(), stringId: undefined }
    }
    if (sheet.structureVersion === 1) {
      const cellIndex = sheet.grid.getPhysical(row, col)
      return readCellValueForLookup(cellIndex === -1 ? undefined : cellIndex)
    }
    return readCellValueForLookup(sheet.logical.getVisibleCell(row, col))
  }

  const readDirectCriteriaOperandValue = (operand: RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['criterion']): CellValue => {
    return operand.kind === 'literal' ? operand.value : readCellValueForLookup(operand.cellIndex).value
  }

  const directCriteriaMatchesChangedAggregateRow = (
    directCriteria: RuntimeDirectCriteriaDescriptor,
    aggregateRange: NonNullable<RuntimeDirectCriteriaDescriptor['aggregateRange']>,
    requestRow: number,
  ): boolean | undefined => {
    const rowOffset = requestRow - aggregateRange.rowStart
    for (let index = 0; index < directCriteria.criteriaPairs.length; index += 1) {
      const pair = directCriteria.criteriaPairs[index]!
      const criteria = readDirectCriteriaOperandValue(pair.criterion)
      if (criteria.tag === ValueTag.Error) {
        return undefined
      }
      const candidate = readCellValueAtForLookup(pair.range.sheetName, pair.range.rowStart + rowOffset, pair.range.col).value
      if (!matchesCompiledCriteria(candidate, compileCriteriaMatcher(criteria))) {
        return false
      }
    }
    return true
  }

  const tryDirectCriteriaSumDelta = (
    directCriteria: RuntimeDirectCriteriaDescriptor,
    request: {
      sheetName: string
      row: number
      col: number
      oldValue?: CellValue
      newValue?: CellValue
    },
  ): number | undefined => {
    const aggregateRange = directCriteria.aggregateRange
    if (
      directCriteria.aggregateKind !== 'sum' ||
      aggregateRange === undefined ||
      aggregateRange.sheetName !== request.sheetName ||
      aggregateRange.col !== request.col ||
      request.row < aggregateRange.rowStart ||
      request.row > aggregateRange.rowEnd ||
      request.oldValue === undefined ||
      request.newValue === undefined
    ) {
      return undefined
    }
    const oldContribution = directAggregateNumericContribution(request.oldValue)
    const newContribution = directAggregateNumericContribution(request.newValue)
    if (oldContribution === undefined || newContribution === undefined) {
      return undefined
    }
    const matchesRow = directCriteriaMatchesChangedAggregateRow(directCriteria, aggregateRange, request.row)
    if (matchesRow === undefined) {
      return undefined
    }
    return matchesRow ? newContribution - oldContribution : 0
  }

  const readApproximateNumericValueAtForLookup = (sheetName: string, row: number, col: number): number | undefined => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return 0
    }
    if (sheet.structureVersion === 1) {
      const cellIndex = sheet.grid.getPhysical(row, col)
      return readApproximateNumericValueForLookup(cellIndex === -1 ? undefined : cellIndex)
    }
    return readApproximateNumericValueForLookup(sheet.logical.getVisibleCell(row, col))
  }

  const isLocallySortedNumericWrite = (
    sheetName: string,
    row: number,
    col: number,
    rowStart: number,
    rowEnd: number,
    matchMode: 1 | -1,
    current: number,
  ): boolean => {
    if (row > rowStart) {
      const previous = readApproximateNumericValueAtForLookup(sheetName, row - 1, col)
      if (previous === undefined || (matchMode === 1 ? previous > current : previous < current)) {
        return false
      }
    }
    if (row < rowEnd) {
      const next = readApproximateNumericValueAtForLookup(sheetName, row + 1, col)
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
    current: string,
  ): boolean => {
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

  const planSingleExactLookupNumericColumnWrite = (
    formulaCellIndex: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): LookupNumericColumnWritePlan => {
    const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
    if (directLookup?.kind !== 'exact' && directLookup?.kind !== 'exact-uniform-numeric') {
      return { handled: false }
    }
    if (directLookup.kind === 'exact-uniform-numeric' && directLookup.tailPatch !== undefined) {
      return { handled: false }
    }
    const { rowStart, rowEnd } = directLookupRowBounds(directLookup)
    if (row < rowStart || row > rowEnd) {
      return { handled: true }
    }
    if (
      directLookup.kind === 'exact-uniform-numeric' &&
      canSkipUniformExactNumericTailWriteFromCurrentResult(
        args.state.workbook.cellStore,
        formulaCellIndex,
        directLookup,
        row,
        oldNumeric,
        newNumeric,
      )
    ) {
      return { handled: true, tailPatchTarget: directLookup }
    }
    const operandNumeric = readExactNumericValueForLookup(directLookup.operandCellIndex)
    if (operandNumeric === undefined) {
      return { handled: false }
    }
    return { handled: !sameExactNumericValue(oldNumeric, operandNumeric) && !sameExactNumericValue(newNumeric, operandNumeric) }
  }

  const canSkipSingleExactLookupNumericColumnWrite = (
    formulaCellIndex: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    return planSingleExactLookupNumericColumnWrite(formulaCellIndex, row, oldNumeric, newNumeric).handled
  }

  const planExactLookupNumericColumnWrite = (
    sheetId: number,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): LookupNumericColumnWritePlan => {
    const lookupEntity = makeExactLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    if (singleDependent === -1) {
      return { handled: true }
    }
    if (singleDependent >= 0) {
      return planSingleExactLookupNumericColumnWrite(singleDependent, row, oldNumeric, newNumeric)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipSingleExactLookupNumericColumnWrite(dependents[index]!, row, oldNumeric, newNumeric)) {
        return { handled: false }
      }
    }
    return { handled: true }
  }

  const canSkipExactLookupNumericColumnWrite = (
    sheetId: number,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    const lookupEntity = makeExactLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      return canSkipSingleExactLookupNumericColumnWrite(singleDependent, row, oldNumeric, newNumeric)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipSingleExactLookupNumericColumnWrite(dependents[index]!, row, oldNumeric, newNumeric)) {
        return false
      }
    }
    return true
  }

  const planSingleApproximateLookupNumericColumnWrite = (
    formulaCellIndex: number,
    sheetName: string,
    row: number,
    col: number,
    oldNumeric: number,
    newNumeric: number,
  ): LookupNumericColumnWritePlan => {
    const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
    if (directLookup?.kind !== 'approximate' && directLookup?.kind !== 'approximate-uniform-numeric') {
      return { handled: false }
    }
    if (directLookup.kind === 'approximate-uniform-numeric' && directLookup.tailPatch !== undefined) {
      return { handled: false }
    }
    const { rowStart, rowEnd } = directLookupRowBounds(directLookup)
    if (row < rowStart || row > rowEnd) {
      return { handled: true }
    }
    if (
      directLookup.kind === 'approximate-uniform-numeric' &&
      canSkipUniformApproximateNumericTailWriteFromCurrentResult(
        args.state.workbook.cellStore,
        formulaCellIndex,
        directLookup,
        row,
        oldNumeric,
        newNumeric,
      )
    ) {
      return { handled: true, tailPatchTarget: directLookup }
    }
    const operandNumeric = readApproximateNumericValueForLookup(directLookup.operandCellIndex)
    if (operandNumeric === undefined) {
      return { handled: false }
    }
    const matchMode = directLookup.matchMode
    if (directLookup.kind === 'approximate-uniform-numeric') {
      if (canSkipUniformApproximateNumericTailWrite(directLookup, row, operandNumeric, oldNumeric, newNumeric)) {
        return { handled: true, tailPatchTarget: directLookup }
      }
    }
    if (matchMode === 1) {
      return {
        handled:
          oldNumeric > operandNumeric &&
          newNumeric > operandNumeric &&
          isLocallySortedNumericWrite(sheetName, row, col, rowStart, rowEnd, matchMode, newNumeric),
      }
    }
    return {
      handled:
        oldNumeric < operandNumeric &&
        newNumeric < operandNumeric &&
        isLocallySortedNumericWrite(sheetName, row, col, rowStart, rowEnd, matchMode, newNumeric),
    }
  }

  const canSkipSingleApproximateLookupNumericColumnWrite = (
    formulaCellIndex: number,
    sheetName: string,
    row: number,
    col: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    return planSingleApproximateLookupNumericColumnWrite(formulaCellIndex, sheetName, row, col, oldNumeric, newNumeric).handled
  }

  const planApproximateLookupNumericColumnWrite = (
    sheetId: number,
    sheetName: string,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): LookupNumericColumnWritePlan => {
    const lookupEntity = makeSortedLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    if (singleDependent === -1) {
      return { handled: true }
    }
    if (singleDependent >= 0) {
      return planSingleApproximateLookupNumericColumnWrite(singleDependent, sheetName, row, col, oldNumeric, newNumeric)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipSingleApproximateLookupNumericColumnWrite(dependents[index]!, sheetName, row, col, oldNumeric, newNumeric)) {
        return { handled: false }
      }
    }
    return { handled: true }
  }

  const canSkipApproximateLookupNumericColumnWrite = (
    sheetId: number,
    sheetName: string,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    const lookupEntity = makeSortedLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      return canSkipSingleApproximateLookupNumericColumnWrite(singleDependent, sheetName, row, col, oldNumeric, newNumeric)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipSingleApproximateLookupNumericColumnWrite(dependents[index]!, sheetName, row, col, oldNumeric, newNumeric)) {
        return false
      }
    }
    return true
  }

  const canSkipApproximateLookupNewNumericColumnWrite = (sheetId: number, col: number, row: number): boolean => {
    const lookupEntity = makeSortedLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    const canSkipDependent = (formulaCellIndex: number): boolean => {
      const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
      if (directLookup?.kind !== 'approximate' && directLookup?.kind !== 'approximate-uniform-numeric') {
        return false
      }
      const { rowStart, rowEnd } = directLookupRowBounds(directLookup)
      return row < rowStart || row > rowEnd
    }
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      return canSkipDependent(singleDependent)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipDependent(dependents[index]!)) {
        return false
      }
    }
    return true
  }

  const canPatchUniformLookupTailWrite = (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' | 'approximate-uniform-numeric' }>,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    if (directLookup.tailPatch !== undefined || row !== directLookup.rowEnd || directLookup.length < 2) {
      return false
    }
    const expectedOldTail = directLookup.start + directLookup.step * (directLookup.length - 1)
    if (!sameExactNumericValue(oldNumeric, expectedOldTail)) {
      return false
    }
    return directLookup.step > 0 ? newNumeric > oldNumeric : newNumeric < oldNumeric
  }

  const patchUniformLookupTailWrites = (request: {
    sheetId: number
    col: number
    row: number
    oldNumeric: number
    newNumeric: number
    exact: boolean
    sorted: boolean
  }): { exact: boolean; sorted: boolean } => {
    const patchDependents = (entityId: number, kind: 'exact-uniform-numeric' | 'approximate-uniform-numeric'): boolean => {
      const singleDependent = args.getSingleEntityDependent(entityId)
      if (singleDependent === -1) {
        return true
      }
      const currentSheet = args.state.workbook.getSheetById(request.sheetId)
      const currentColumnVersion = currentSheet?.columnVersions[request.col] ?? 0
      if (singleDependent >= 0) {
        const directLookup = args.state.formulas.get(singleDependent)?.directLookup
        if (
          directLookup?.kind !== kind ||
          !canPatchUniformLookupTailWrite(directLookup, request.row, request.oldNumeric, request.newNumeric)
        ) {
          return false
        }
        directLookup.tailPatch = {
          row: request.row,
          oldNumeric: request.oldNumeric,
          newNumeric: request.newNumeric,
          columnVersion: currentColumnVersion,
        }
        return true
      }
      const dependents = args.getEntityDependents(entityId)
      if (dependents.length === 0) {
        return true
      }
      for (let index = 0; index < dependents.length; index += 1) {
        const directLookup = args.state.formulas.get(dependents[index]!)?.directLookup
        if (
          directLookup?.kind !== kind ||
          !canPatchUniformLookupTailWrite(directLookup, request.row, request.oldNumeric, request.newNumeric)
        ) {
          return false
        }
      }
      for (let index = 0; index < dependents.length; index += 1) {
        const directLookup = args.state.formulas.get(dependents[index]!)?.directLookup
        if (directLookup?.kind === kind) {
          directLookup.tailPatch = {
            row: request.row,
            oldNumeric: request.oldNumeric,
            newNumeric: request.newNumeric,
            columnVersion: currentColumnVersion,
          }
        }
      }
      return true
    }

    return {
      exact: request.exact && patchDependents(makeExactLookupColumnEntity(request.sheetId, request.col), 'exact-uniform-numeric'),
      sorted: request.sorted && patchDependents(makeSortedLookupColumnEntity(request.sheetId, request.col), 'approximate-uniform-numeric'),
    }
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
    const operandNumeric = readApproximateNumericValueForLookup(directLookup.operandCellIndex)
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
          isLocallySortedNumericWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode, newNumeric)
        )
      }
      return (
        oldNumeric < operandNumeric &&
        newNumeric < operandNumeric &&
        isLocallySortedNumericWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode, newNumeric)
      )
    }
    const operand = readCellValueForLookup(directLookup.operandCellIndex)
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
        isLocallySortedTextWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode, newText)
      )
    }
    return (
      oldText < operandText &&
      newText < operandText &&
      isLocallySortedTextWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode, newText)
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
    addEngineCounter(args.state.counters, 'cycleFormulaScans')
    let found = false
    args.state.formulas.forEach((_formula, cellIndex) => {
      if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
        found = true
      }
    })
    return found
  }

  const directScalarNumericValue = (value: CellValue): number | undefined => {
    switch (value.tag) {
      case ValueTag.Number:
        return Object.is(value.value, -0) ? 0 : value.value
      case ValueTag.Boolean:
        return value.value ? 1 : 0
      case ValueTag.Empty:
        return 0
      case ValueTag.Error:
      case ValueTag.String:
        return undefined
    }
  }

  const directScalarCellNumericValue = (cellIndex: number | undefined): number | undefined => {
    if (cellIndex === undefined) {
      return 0
    }
    const cellStore = args.state.workbook.cellStore
    switch ((cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) {
      case ValueTag.Number: {
        const value = cellStore.numbers[cellIndex] ?? 0
        return Object.is(value, -0) ? 0 : value
      }
      case ValueTag.Boolean:
        return (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 1 : 0
      case ValueTag.Empty:
        return 0
      case ValueTag.Error:
      case ValueTag.String:
        return undefined
    }
  }

  const readDirectScalarOperandNumber = (
    operand: RuntimeDirectScalarOperand,
    changedCellIndex: number,
    replacementValue: CellValue,
    touched: { value: boolean },
  ): number | undefined => {
    switch (operand.kind) {
      case 'literal-number':
        return operand.value
      case 'error':
        return undefined
      case 'cell':
        if (operand.cellIndex === changedCellIndex) {
          touched.value = true
          return directScalarNumericValue(replacementValue)
        }
        return directScalarNumericValue(args.state.workbook.cellStore.getValue(operand.cellIndex, (id) => args.state.strings.get(id)))
    }
  }

  const evaluateDirectScalarNumber = (
    directScalar: RuntimeDirectScalarDescriptor,
    changedCellIndex: number,
    replacementValue: CellValue,
    touched: { value: boolean },
  ): number | undefined => {
    if (directScalar.kind === 'abs') {
      const operand = readDirectScalarOperandNumber(directScalar.operand, changedCellIndex, replacementValue, touched)
      return operand === undefined ? undefined : Math.abs(operand)
    }
    const left = readDirectScalarOperandNumber(directScalar.left, changedCellIndex, replacementValue, touched)
    const right = readDirectScalarOperandNumber(directScalar.right, changedCellIndex, replacementValue, touched)
    if (left === undefined || right === undefined) {
      return undefined
    }
    switch (directScalar.operator) {
      case '+':
        return left + right
      case '-':
        return left - right
      case '*':
        return left * right
      case '/':
        return right === 0 ? undefined : left / right
    }
  }

  const tryDirectScalarNumericDelta = (
    directScalar: RuntimeDirectScalarDescriptor,
    changedCellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
  ): number | undefined => {
    const oldChangedNumber = directScalarNumericValue(oldValue)
    const newChangedNumber = directScalarNumericValue(newValue)
    if (oldChangedNumber !== undefined && newChangedNumber !== undefined) {
      if (directScalar.kind === 'abs') {
        return directScalar.operand.kind === 'cell' && directScalar.operand.cellIndex === changedCellIndex
          ? Math.abs(newChangedNumber) - Math.abs(oldChangedNumber)
          : undefined
      }
      const changedDelta = newChangedNumber - oldChangedNumber
      if (directScalar.left.kind === 'cell' && directScalar.left.cellIndex === changedCellIndex) {
        switch (directScalar.operator) {
          case '+':
          case '-':
            return changedDelta
          case '*':
            if (directScalar.right.kind === 'literal-number') {
              return changedDelta * directScalar.right.value
            }
            break
          case '/':
            return directScalar.right.kind === 'literal-number' && directScalar.right.value !== 0
              ? changedDelta / directScalar.right.value
              : undefined
        }
      }
      if (directScalar.right.kind === 'cell' && directScalar.right.cellIndex === changedCellIndex) {
        switch (directScalar.operator) {
          case '+':
            return changedDelta
          case '-':
            return -changedDelta
          case '*':
            if (directScalar.left.kind === 'literal-number') {
              return directScalar.left.value * changedDelta
            }
            break
          case '/':
            return directScalar.left.kind === 'literal-number' && oldChangedNumber !== 0 && newChangedNumber !== 0
              ? directScalar.left.value / newChangedNumber - directScalar.left.value / oldChangedNumber
              : undefined
        }
      }
    }
    const oldTouched = { value: false }
    const oldResult = evaluateDirectScalarNumber(directScalar, changedCellIndex, oldValue, oldTouched)
    if (!oldTouched.value || oldResult === undefined) {
      return undefined
    }
    const newTouched = { value: false }
    const newResult = evaluateDirectScalarNumber(directScalar, changedCellIndex, newValue, newTouched)
    if (!newTouched.value || newResult === undefined) {
      return undefined
    }
    return newResult - oldResult
  }

  const readDirectScalarOperandNumberWithReplacement = (
    operand: RuntimeDirectScalarOperand,
    changedCellIndex: number,
    replacementNumber: number,
    touched: { value: boolean },
  ): number | undefined => {
    switch (operand.kind) {
      case 'literal-number':
        return operand.value
      case 'error':
        return undefined
      case 'cell':
        if (operand.cellIndex === changedCellIndex) {
          touched.value = true
          return replacementNumber
        }
        return directScalarNumericValue(args.state.workbook.cellStore.getValue(operand.cellIndex, (id) => args.state.strings.get(id)))
    }
  }

  const evaluateDirectScalarNumberWithReplacement = (
    directScalar: RuntimeDirectScalarDescriptor,
    changedCellIndex: number,
    replacementNumber: number,
    touched: { value: boolean },
  ): number | undefined => {
    if (directScalar.kind === 'abs') {
      const operand = readDirectScalarOperandNumberWithReplacement(directScalar.operand, changedCellIndex, replacementNumber, touched)
      return operand === undefined ? undefined : Math.abs(operand)
    }
    const left = readDirectScalarOperandNumberWithReplacement(directScalar.left, changedCellIndex, replacementNumber, touched)
    const right = readDirectScalarOperandNumberWithReplacement(directScalar.right, changedCellIndex, replacementNumber, touched)
    if (left === undefined || right === undefined) {
      return undefined
    }
    switch (directScalar.operator) {
      case '+':
        return left + right
      case '-':
        return left - right
      case '*':
        return left * right
      case '/':
        return right === 0 ? undefined : left / right
    }
  }

  const tryDirectScalarNumericDeltaFromNumbers = (
    directScalar: RuntimeDirectScalarDescriptor,
    changedCellIndex: number,
    oldChangedNumber: number,
    newChangedNumber: number,
  ): number | undefined => {
    if (directScalar.kind === 'abs') {
      return directScalar.operand.kind === 'cell' && directScalar.operand.cellIndex === changedCellIndex
        ? Math.abs(newChangedNumber) - Math.abs(oldChangedNumber)
        : undefined
    }
    const changedDelta = newChangedNumber - oldChangedNumber
    if (directScalar.left.kind === 'cell' && directScalar.left.cellIndex === changedCellIndex) {
      switch (directScalar.operator) {
        case '+':
        case '-':
          return changedDelta
        case '*':
          if (directScalar.right.kind === 'literal-number') {
            return changedDelta * directScalar.right.value
          }
          break
        case '/':
          return directScalar.right.kind === 'literal-number' && directScalar.right.value !== 0
            ? changedDelta / directScalar.right.value
            : undefined
      }
    }
    if (directScalar.right.kind === 'cell' && directScalar.right.cellIndex === changedCellIndex) {
      switch (directScalar.operator) {
        case '+':
          return changedDelta
        case '-':
          return -changedDelta
        case '*':
          if (directScalar.left.kind === 'literal-number') {
            return directScalar.left.value * changedDelta
          }
          break
        case '/':
          return directScalar.left.kind === 'literal-number' && oldChangedNumber !== 0 && newChangedNumber !== 0
            ? directScalar.left.value / newChangedNumber - directScalar.left.value / oldChangedNumber
            : undefined
      }
    }
    const oldTouched = { value: false }
    const oldResult = evaluateDirectScalarNumberWithReplacement(directScalar, changedCellIndex, oldChangedNumber, oldTouched)
    if (!oldTouched.value || oldResult === undefined) {
      return undefined
    }
    const newTouched = { value: false }
    const newResult = evaluateDirectScalarNumberWithReplacement(directScalar, changedCellIndex, newChangedNumber, newTouched)
    if (!newTouched.value || newResult === undefined) {
      return undefined
    }
    return newResult - oldResult
  }

  const lookupSheetForUniformLookup = (directLookup: UniformNumericDirectLookup, lookupSheetHint?: SheetRecord): SheetRecord | undefined =>
    lookupSheetHint?.id === directLookup.sheetId ? lookupSheetHint : args.state.workbook.getSheetById(directLookup.sheetId)

  const tryDirectUniformLookupCurrentResult = (formulaCellIndex: number): DirectScalarCurrentOperand | undefined => {
    const formula = args.state.formulas.get(formulaCellIndex)
    const directLookup = formula?.directLookup
    if (directLookup === undefined) {
      return undefined
    }
    const cellStore = args.state.workbook.cellStore
    if (directLookup.kind === 'exact-uniform-numeric') {
      const lookupSheet = args.state.workbook.getSheetById(directLookup.sheetId)
      if (!directLookupVersionMatches(lookupSheet, directLookup)) {
        return undefined
      }
      const tag = cellStore.tags[directLookup.operandCellIndex]
      if (tag === ValueTag.Error) {
        return undefined
      }
      if (tag !== ValueTag.Number) {
        return { kind: 'error', code: ErrorCode.NA }
      }
      const lookupValue = Object.is(cellStore.numbers[directLookup.operandCellIndex] ?? 0, -0)
        ? 0
        : (cellStore.numbers[directLookup.operandCellIndex] ?? 0)
      return exactUniformLookupCurrentResult(directLookup, lookupValue)
    }
    if (directLookup.kind !== 'approximate-uniform-numeric') {
      return undefined
    }
    const lookupSheet = args.state.workbook.getSheetById(directLookup.sheetId)
    if (!directLookupVersionMatches(lookupSheet, directLookup)) {
      return undefined
    }
    const tag = cellStore.tags[directLookup.operandCellIndex]
    let lookupValue = 0
    switch (tag) {
      case undefined:
      case ValueTag.Empty:
        lookupValue = 0
        break
      case ValueTag.Number:
        lookupValue = Object.is(cellStore.numbers[directLookup.operandCellIndex] ?? 0, -0)
          ? 0
          : (cellStore.numbers[directLookup.operandCellIndex] ?? 0)
        break
      case ValueTag.Boolean:
        lookupValue = (cellStore.numbers[directLookup.operandCellIndex] ?? 0) !== 0 ? 1 : 0
        break
      case ValueTag.Error:
      case ValueTag.String:
        return undefined
    }
    return approximateUniformLookupCurrentResult(directLookup, lookupValue)
  }

  const tryDirectUniformLookupCurrentResultFromNumeric = (
    formulaCellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    lookupSheetHint?: SheetRecord,
  ): DirectScalarCurrentOperand | undefined => {
    const formula = args.state.formulas.get(formulaCellIndex)
    const directLookup = formula?.directLookup
    if (directLookup === undefined) {
      return undefined
    }
    if (directLookup.kind === 'exact-uniform-numeric') {
      if (exactLookupValue === undefined) {
        return undefined
      }
      const lookupSheet = lookupSheetForUniformLookup(directLookup, lookupSheetHint)
      if (!directLookupVersionMatches(lookupSheet, directLookup)) {
        return undefined
      }
      return exactUniformLookupCurrentResult(directLookup, exactLookupValue)
    }
    if (directLookup.kind !== 'approximate-uniform-numeric' || approximateLookupValue === undefined) {
      return undefined
    }
    const lookupSheet = lookupSheetForUniformLookup(directLookup, lookupSheetHint)
    if (!directLookupVersionMatches(lookupSheet, directLookup)) {
      return undefined
    }
    return approximateUniformLookupCurrentResult(directLookup, approximateLookupValue)
  }

  const tryDirectUniformLookupNumericResultFromDescriptor = (
    directLookup: RuntimeDirectLookupDescriptor | undefined,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    lookupSheetHint?: SheetRecord,
  ): number | undefined => {
    if (directLookup?.kind === 'exact-uniform-numeric') {
      if (exactLookupValue === undefined) {
        return undefined
      }
      const lookupSheet = lookupSheetForUniformLookup(directLookup, lookupSheetHint)
      return directLookupVersionMatches(lookupSheet, directLookup)
        ? exactUniformLookupNumericResult(directLookup, exactLookupValue)
        : undefined
    }
    if (directLookup?.kind === 'approximate-uniform-numeric') {
      if (approximateLookupValue === undefined) {
        return undefined
      }
      const lookupSheet = lookupSheetForUniformLookup(directLookup, lookupSheetHint)
      return directLookupVersionMatches(lookupSheet, directLookup)
        ? approximateUniformLookupNumericResult(directLookup, approximateLookupValue)
        : undefined
    }
    return undefined
  }

  const canEvaluateDirectUniformLookupCurrentResultFromNumeric = (
    formulaCellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
  ): boolean => {
    const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
    if (directLookup?.kind === 'exact-uniform-numeric') {
      const lookupSheet = args.state.workbook.getSheetById(directLookup.sheetId)
      if (!directLookupVersionMatches(lookupSheet, directLookup)) {
        return false
      }
      return exactLookupValue !== undefined
    }
    if (directLookup?.kind !== 'approximate-uniform-numeric') {
      return false
    }
    const lookupSheet = args.state.workbook.getSheetById(directLookup.sheetId)
    if (!directLookupVersionMatches(lookupSheet, directLookup)) {
      return false
    }
    return approximateLookupValue !== undefined
  }

  const directScalarCurrentResultMatchesCell = (cellIndex: number, result: DirectScalarCurrentOperand): boolean => {
    const cellStore = args.state.workbook.cellStore
    const currentTag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    if (result.kind === 'number') {
      return currentTag === ValueTag.Number && Object.is(cellStore.numbers[cellIndex] ?? 0, result.value)
    }
    return currentTag === ValueTag.Error && ((cellStore.errors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None) === result.code
  }

  const directScalarNumericResultMatchesCell = (cellIndex: number, result: number): boolean => {
    const cellStore = args.state.workbook.cellStore
    return cellStore.tags[cellIndex] === ValueTag.Number && Object.is(cellStore.numbers[cellIndex] ?? 0, result)
  }

  const addDirectLookupCurrentResultIfChanged = (
    formulaCellIndex: number,
    result: DirectScalarCurrentOperand,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ): boolean => {
    if (!directScalarCurrentResultMatchesCell(formulaCellIndex, result)) {
      postRecalcDirectFormulaIndices.addCurrentResult(formulaCellIndex, result)
      return true
    }
    return false
  }

  const canUseDirectFormulaPostRecalc = (formulaCellIndex: number): boolean => {
    const formula = args.state.formulas.get(formulaCellIndex)
    return (
      formula !== undefined &&
      (formula.directScalar === undefined || !formula.compiled.deps.some((dependency) => dependency.includes('!'))) &&
      (formula.directLookup !== undefined ||
        formula.directAggregate !== undefined ||
        formula.directScalar !== undefined ||
        formula.directCriteria !== undefined) &&
      hasNoCellDependents(formulaCellIndex)
    )
  }

  const markPostRecalcDirectLookupCurrentDependentsFromNumeric = (
    cellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ): boolean => {
    const singleDependent = args.getSingleEntityDependent(makeCellEntity(cellIndex))
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      if (!canUseDirectFormulaPostRecalc(singleDependent)) {
        return false
      }
      const directLookupResult = tryDirectUniformLookupCurrentResultFromNumeric(singleDependent, exactLookupValue, approximateLookupValue)
      if (directLookupResult === undefined) {
        return false
      }
      addDirectLookupCurrentResultIfChanged(singleDependent, directLookupResult, postRecalcDirectFormulaIndices)
      return true
    }

    const dependents = args.getEntityDependents(makeCellEntity(cellIndex))
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (!canUseDirectFormulaPostRecalc(formulaCellIndex)) {
        return false
      }
      if (!canEvaluateDirectUniformLookupCurrentResultFromNumeric(formulaCellIndex, exactLookupValue, approximateLookupValue)) {
        return false
      }
    }
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      const directLookupResult = tryDirectUniformLookupCurrentResultFromNumeric(formulaCellIndex, exactLookupValue, approximateLookupValue)
      if (directLookupResult === undefined) {
        return false
      }
      addDirectLookupCurrentResultIfChanged(formulaCellIndex, directLookupResult, postRecalcDirectFormulaIndices)
    }
    return true
  }

  const markPostRecalcDirectFormulaDependents = (
    cellIndex: number,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
    oldValue?: CellValue,
    newValue?: CellValue,
  ): boolean => {
    const singleDependent = args.getSingleEntityDependent(makeCellEntity(cellIndex))
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      if (!canUseDirectFormulaPostRecalc(singleDependent)) {
        return false
      }
      if (oldValue === undefined || newValue === undefined) {
        postRecalcDirectFormulaIndices.add(singleDependent)
        return true
      }
      const directScalar = args.state.formulas.get(singleDependent)?.directScalar
      if (directScalar !== undefined) {
        const delta = tryDirectScalarNumericDelta(directScalar, cellIndex, oldValue, newValue)
        if (delta !== undefined) {
          postRecalcDirectFormulaIndices.addScalarDelta(singleDependent, delta)
          return true
        }
      }
      const directLookupResult = tryDirectUniformLookupCurrentResult(singleDependent)
      if (directLookupResult !== undefined) {
        if (!addDirectLookupCurrentResultIfChanged(singleDependent, directLookupResult, postRecalcDirectFormulaIndices)) {
          postRecalcDirectFormulaIndices.markDirectFormulaInputCovered(cellIndex)
        }
        return true
      }
      postRecalcDirectFormulaIndices.add(singleDependent)
      return true
    }
    const dependents = args.getEntityDependents(makeCellEntity(cellIndex))
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canUseDirectFormulaPostRecalc(dependents[index]!)) {
        return false
      }
    }
    if (oldValue !== undefined && newValue !== undefined && dependents.length > 16) {
      let canUseBulkScalarDeltas = true
      let commonDelta: number | undefined
      let allDeltasMatch = true
      for (let index = 0; index < dependents.length; index += 1) {
        const formulaCellIndex = dependents[index]!
        const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
        const delta = directScalar === undefined ? undefined : tryDirectScalarNumericDelta(directScalar, cellIndex, oldValue, newValue)
        if (delta === undefined) {
          canUseBulkScalarDeltas = false
          break
        }
        if (commonDelta === undefined) {
          commonDelta = delta
        } else if (!Object.is(commonDelta, delta)) {
          allDeltasMatch = false
        }
      }
      if (canUseBulkScalarDeltas) {
        if (allDeltasMatch && commonDelta !== undefined) {
          postRecalcDirectFormulaIndices.appendConstantDelta(dependents, commonDelta, 'scalar')
        } else {
          const deltaCellIndices: number[] = []
          const deltas: number[] = []
          for (let index = 0; index < dependents.length; index += 1) {
            const formulaCellIndex = dependents[index]!
            const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
            const delta = directScalar === undefined ? undefined : tryDirectScalarNumericDelta(directScalar, cellIndex, oldValue, newValue)
            if (delta === undefined) {
              canUseBulkScalarDeltas = false
              break
            }
            deltaCellIndices.push(formulaCellIndex)
            deltas.push(delta)
          }
          if (!canUseBulkScalarDeltas) {
            return false
          }
          postRecalcDirectFormulaIndices.appendDeltas(deltaCellIndices, deltas, 'scalar')
        }
        return true
      }
    }
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
      if (oldValue === undefined || newValue === undefined) {
        postRecalcDirectFormulaIndices.add(formulaCellIndex)
        continue
      }
      if (directScalar !== undefined) {
        const delta = tryDirectScalarNumericDelta(directScalar, cellIndex, oldValue, newValue)
        if (delta !== undefined) {
          postRecalcDirectFormulaIndices.addScalarDelta(formulaCellIndex, delta)
          continue
        }
      }
      const directLookupResult = tryDirectUniformLookupCurrentResult(formulaCellIndex)
      if (directLookupResult !== undefined) {
        addDirectLookupCurrentResultIfChanged(formulaCellIndex, directLookupResult, postRecalcDirectFormulaIndices)
        continue
      }
      postRecalcDirectFormulaIndices.add(formulaCellIndex)
    }
    return true
  }

  const markPostRecalcDirectScalarNumericDependents = (
    cellIndex: number,
    oldNumber: number,
    newNumber: number,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
    exactLookupValue?: number,
    approximateLookupValue?: number,
  ): boolean => {
    const singleDependent = args.getSingleEntityDependent(makeCellEntity(cellIndex))
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      if (!canUseDirectFormulaPostRecalc(singleDependent)) {
        return false
      }
      const directScalar = args.state.formulas.get(singleDependent)?.directScalar
      const delta =
        directScalar === undefined ? undefined : tryDirectScalarNumericDeltaFromNumbers(directScalar, cellIndex, oldNumber, newNumber)
      if (delta === undefined) {
        const directLookupResult = tryDirectUniformLookupCurrentResultFromNumeric(singleDependent, exactLookupValue, approximateLookupValue)
        if (directLookupResult === undefined) {
          return false
        }
        if (!addDirectLookupCurrentResultIfChanged(singleDependent, directLookupResult, postRecalcDirectFormulaIndices)) {
          postRecalcDirectFormulaIndices.markDirectFormulaInputCovered(cellIndex)
        }
        return true
      }
      postRecalcDirectFormulaIndices.addScalarDelta(singleDependent, delta)
      return true
    }

    const dependents = args.getEntityDependents(makeCellEntity(cellIndex))
    if (dependents.length === 0) {
      return true
    }
    let commonDelta: number | undefined
    let allDeltasMatch = true
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (!canUseDirectFormulaPostRecalc(formulaCellIndex)) {
        return false
      }
      const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
      const delta =
        directScalar === undefined ? undefined : tryDirectScalarNumericDeltaFromNumbers(directScalar, cellIndex, oldNumber, newNumber)
      if (delta === undefined) {
        return false
      }
      if (commonDelta === undefined) {
        commonDelta = delta
      } else if (!Object.is(commonDelta, delta)) {
        allDeltasMatch = false
        break
      }
    }
    if (allDeltasMatch && commonDelta !== undefined) {
      postRecalcDirectFormulaIndices.appendConstantDelta(dependents, commonDelta, 'scalar')
      return true
    }
    const deltas: number[] = []
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (!canUseDirectFormulaPostRecalc(formulaCellIndex)) {
        return false
      }
      const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
      const delta =
        directScalar === undefined ? undefined : tryDirectScalarNumericDeltaFromNumbers(directScalar, cellIndex, oldNumber, newNumber)
      if (delta === undefined) {
        return false
      }
      deltas[index] = delta
    }
    postRecalcDirectFormulaIndices.appendDeltas(dependents, deltas, 'scalar')
    return true
  }

  const tryMarkDirectScalarLinearDeltaClosure = (
    rootCellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ): boolean => {
    const oldRootNumber = directScalarNumericValue(oldValue)
    const newRootNumber = directScalarNumericValue(newValue)
    if (oldRootNumber === undefined || newRootNumber === undefined) {
      return false
    }
    let currentCellIndex = rootCellIndex
    let oldNumber = oldRootNumber
    let newNumber = newRootNumber
    let closureCount = 0
    const cellIndices: number[] = []
    let deltas: number[] | undefined
    let commonDelta: number | undefined
    let canUseValidatedTerminalWrites = true
    for (;;) {
      if (closureCount > DIRECT_SCALAR_DELTA_CLOSURE_LIMIT) {
        return false
      }
      const formulaCellIndex = args.getSingleEntityDependent(makeCellEntity(currentCellIndex))
      if (formulaCellIndex === -1) {
        break
      }
      if (formulaCellIndex < 0) {
        return false
      }
      if (formulaCellIndex === rootCellIndex) {
        return false
      }
      const formula = args.state.formulas.get(formulaCellIndex)
      if (
        !formula ||
        formula.directScalar === undefined ||
        formula.compiled.volatile ||
        formula.compiled.producesSpill ||
        ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
      ) {
        return false
      }
      const formulaDelta = tryDirectScalarNumericDeltaFromNumbers(formula.directScalar, currentCellIndex, oldNumber, newNumber)
      if (formulaDelta === undefined) {
        return false
      }
      const formulaOldNumber = directScalarCellNumericValue(formulaCellIndex)
      if (formulaOldNumber === undefined) {
        return false
      }
      if (canUseValidatedTerminalWrites && !canSkipDirectFormulaColumnVersion(formulaCellIndex)) {
        canUseValidatedTerminalWrites = false
      }
      if (commonDelta === undefined) {
        commonDelta = formulaDelta
      } else if (!Object.is(commonDelta, formulaDelta) && deltas === undefined) {
        deltas = []
        for (let index = 0; index < cellIndices.length; index += 1) {
          deltas[index] = commonDelta
        }
      }
      cellIndices.push(formulaCellIndex)
      if (deltas) {
        deltas.push(formulaDelta)
      }
      currentCellIndex = formulaCellIndex
      oldNumber = formulaOldNumber
      newNumber = formulaOldNumber + formulaDelta
      closureCount += 1
    }
    if (cellIndices.length === 0) {
      return false
    }
    if (deltas) {
      postRecalcDirectFormulaIndices.appendDeltas(cellIndices, deltas, 'scalar')
    } else if (commonDelta !== undefined) {
      postRecalcDirectFormulaIndices.appendConstantDelta(cellIndices, commonDelta, 'scalar')
    }
    if (canUseValidatedTerminalWrites) {
      postRecalcDirectFormulaIndices.markScalarDeltaCellsValidated()
    }
    return true
  }

  const markDirectScalarDeltaClosure = (
    rootCellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ): void => {
    const rootDependent = args.getSingleEntityDependent(makeCellEntity(rootCellIndex))
    if (rootDependent < 0 || postRecalcDirectFormulaIndices.hasDelta(rootDependent)) {
      return
    }
    if (args.getSingleEntityDependent(makeCellEntity(rootDependent)) === -1) {
      return
    }
    if (tryMarkDirectScalarLinearDeltaClosure(rootCellIndex, oldValue, newValue, postRecalcDirectFormulaIndices)) {
      return
    }
    const pending: Array<{ cellIndex: number; oldValue: CellValue; newValue: CellValue }> = [
      { cellIndex: rootCellIndex, oldValue, newValue },
    ]
    const closureDeltas = new Map<number, number>()
    const visited = new Set<number>([rootCellIndex])
    for (let cursor = 0; cursor < pending.length; cursor += 1) {
      if (closureDeltas.size > DIRECT_SCALAR_DELTA_CLOSURE_LIMIT) {
        return
      }
      const current = pending[cursor]!
      const dependents = args.getEntityDependents(makeCellEntity(current.cellIndex))
      for (let index = 0; index < dependents.length; index += 1) {
        const formulaCellIndex = dependents[index]!
        if (visited.has(formulaCellIndex)) {
          return
        }
        const formula = args.state.formulas.get(formulaCellIndex)
        if (
          !formula ||
          formula.directScalar === undefined ||
          formula.compiled.volatile ||
          formula.compiled.producesSpill ||
          ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
        ) {
          return
        }
        const formulaDelta = tryDirectScalarNumericDelta(formula.directScalar, current.cellIndex, current.oldValue, current.newValue)
        if (formulaDelta === undefined) {
          return
        }
        const formulaOldValue = args.state.workbook.cellStore.getValue(formulaCellIndex, (id) => args.state.strings.get(id))
        if (formulaOldValue.tag !== ValueTag.Number) {
          return
        }
        const accumulatedDelta = (closureDeltas.get(formulaCellIndex) ?? 0) + formulaDelta
        closureDeltas.set(formulaCellIndex, accumulatedDelta)
        visited.add(formulaCellIndex)
        pending.push({
          cellIndex: formulaCellIndex,
          oldValue: formulaOldValue,
          newValue: { tag: ValueTag.Number, value: formulaOldValue.value + accumulatedDelta },
        })
      }
    }
    closureDeltas.forEach((delta, formulaCellIndex) => {
      postRecalcDirectFormulaIndices.addScalarDelta(formulaCellIndex, delta)
    })
  }

  const canSkipDirtyTraversalForChangedInputs = (
    changedInputCellIndices: U32,
    changedInputCount: number,
    postRecalcDirectFormulaIndices?: DirectFormulaIndexCollection,
    options: {
      readonly lookupHandledInputCellIndices?: readonly number[]
    } = {},
  ): boolean => {
    const lookupInputCovered = (cellIndex: number): boolean => {
      const covered = options.lookupHandledInputCellIndices
      if (covered === undefined) {
        return false
      }
      for (let index = 0; index < covered.length; index += 1) {
        if (covered[index] === cellIndex) {
          return true
        }
      }
      return false
    }
    const lookupDependentsArePostRecalcDirect = (sheetId: number, col: number): boolean => {
      if (postRecalcDirectFormulaIndices === undefined) {
        return false
      }
      const exactLookupDependents = hasTrackedExactLookupDependents(sheetId, col)
        ? args.getEntityDependents(makeExactLookupColumnEntity(sheetId, col))
        : EMPTY_CHANGED_CELLS
      for (let dependentIndex = 0; dependentIndex < exactLookupDependents.length; dependentIndex += 1) {
        if (!postRecalcDirectFormulaIndices.has(exactLookupDependents[dependentIndex]!)) {
          return false
        }
      }
      const sortedLookupDependents = hasTrackedSortedLookupDependents(sheetId, col)
        ? args.getEntityDependents(makeSortedLookupColumnEntity(sheetId, col))
        : EMPTY_CHANGED_CELLS
      for (let dependentIndex = 0; dependentIndex < sortedLookupDependents.length; dependentIndex += 1) {
        if (!postRecalcDirectFormulaIndices.has(sortedLookupDependents[dependentIndex]!)) {
          return false
        }
      }
      return true
    }
    const dependentsArePostRecalcDirect = (dependents: U32): boolean => {
      if (postRecalcDirectFormulaIndices === undefined) {
        return false
      }
      for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
        const dependent = dependents[dependentIndex]!
        if (isRangeEntity(dependent)) {
          return false
        }
        if (!postRecalcDirectFormulaIndices.has(dependent)) {
          return false
        }
      }
      return true
    }
    const rangeDependentsArePostRecalcDirect = (cellIndex: number, requireTrackedRangeDependents = false): boolean => {
      if (postRecalcDirectFormulaIndices === undefined) {
        return false
      }
      if (postRecalcDirectFormulaIndices.hasCoveredDirectRangeInput(cellIndex)) {
        return true
      }
      const cellStore = args.state.workbook.cellStore
      const sheetId = cellStore.sheetIds[cellIndex]
      if (sheetId === undefined) {
        return true
      }
      const sheetName = args.state.workbook.getSheetNameById(sheetId)
      if (!sheetName) {
        return false
      }
      const sheet = args.state.workbook.getSheetById(sheetId)
      let row: number | undefined
      let col: number | undefined
      if (!sheet || sheet.structureVersion === 1) {
        row = cellStore.rows[cellIndex]
        col = cellStore.cols[cellIndex]
      } else {
        const position = sheet.logical.getCellVisiblePosition(cellIndex)
        row = position?.row
        col = position?.col
      }
      if (row === undefined || col === undefined) {
        return false
      }
      if (
        (hasTrackedExactLookupDependents(sheetId, col) || hasTrackedSortedLookupDependents(sheetId, col)) &&
        !lookupInputCovered(cellIndex) &&
        !lookupDependentsArePostRecalcDirect(sheetId, col)
      ) {
        return false
      }
      if (!hasTrackedDirectRangeDependents(sheetId, col)) {
        return !requireTrackedRangeDependents
      }
      const regionDependents = args.collectRegionFormulaDependentsForCell(sheetName, row, col)
      for (let dependentIndex = 0; dependentIndex < regionDependents.length; dependentIndex += 1) {
        if (!postRecalcDirectFormulaIndices.has(regionDependents[dependentIndex]!)) {
          return false
        }
      }
      const directRangeDependents = collectAffectedDirectRangeDependents({ sheetName, row, col })
      for (let dependentIndex = 0; dependentIndex < directRangeDependents.length; dependentIndex += 1) {
        if (!postRecalcDirectFormulaIndices.has(directRangeDependents[dependentIndex]!)) {
          return false
        }
      }
      return true
    }
    for (let index = 0; index < changedInputCount; index += 1) {
      const cellIndex = changedInputCellIndices[index]!
      if (postRecalcDirectFormulaIndices?.hasCoveredDirectFormulaInput(cellIndex) === true) {
        if (!rangeDependentsArePostRecalcDirect(cellIndex)) {
          return false
        }
        continue
      }
      const singleDependent = args.getSingleEntityDependent(makeCellEntity(cellIndex))
      if (singleDependent === -1) {
        if (!rangeDependentsArePostRecalcDirect(cellIndex)) {
          return false
        }
        continue
      }
      if (singleDependent >= 0) {
        if (isRangeEntity(singleDependent)) {
          if (!rangeDependentsArePostRecalcDirect(cellIndex, true)) {
            return false
          }
          continue
        }
        if (postRecalcDirectFormulaIndices?.has(singleDependent) === true) {
          if (!rangeDependentsArePostRecalcDirect(cellIndex)) {
            return false
          }
          continue
        }
        return false
      }
      const dependents = args.getEntityDependents(makeCellEntity(cellIndex))
      if (dependents.length > 0 && !dependentsArePostRecalcDirect(dependents)) {
        return false
      }
      if (!rangeDependentsArePostRecalcDirect(cellIndex)) {
        return false
      }
    }
    return true
  }

  const changedInputsNeedRegionQueryIndices = (
    changedInputCellIndices: U32,
    changedInputCount: number,
    postRecalcDirectFormulaIndices?: DirectFormulaIndexCollection,
  ): boolean => {
    const cellStore = args.state.workbook.cellStore
    for (let index = 0; index < changedInputCount; index += 1) {
      const cellIndex = changedInputCellIndices[index]!
      if (postRecalcDirectFormulaIndices?.hasCoveredDirectRangeInput(cellIndex) === true) {
        continue
      }
      const sheetId = cellStore.sheetIds[cellIndex]
      if (sheetId === undefined) {
        continue
      }
      const sheet = args.state.workbook.getSheetById(sheetId)
      const col = !sheet || sheet.structureVersion === 1 ? cellStore.cols[cellIndex] : sheet.logical.getCellVisiblePosition(cellIndex)?.col
      if (col !== undefined && hasTrackedDirectRangeDependents(sheetId, col)) {
        return true
      }
    }
    return false
  }

  const hasTrackedExactLookupDependents = (sheetId: number, col: number): boolean => {
    const exactLookupEdges = args.reverseState.reverseExactLookupColumnEdges
    if (exactLookupEdges.size === 0) {
      return false
    }
    const slice = exactLookupEdges.get(entityPayload(makeExactLookupColumnEntity(sheetId, col)))
    return slice !== undefined && slice.len > 0
  }

  const hasTrackedSortedLookupDependents = (sheetId: number, col: number): boolean => {
    const sortedLookupEdges = args.reverseState.reverseSortedLookupColumnEdges
    if (sortedLookupEdges.size === 0) {
      return false
    }
    const slice = sortedLookupEdges.get(entityPayload(makeSortedLookupColumnEntity(sheetId, col)))
    return slice !== undefined && slice.len > 0
  }

  const hasTrackedDirectRangeDependents = (sheetId: number, col: number): boolean => {
    let hasRegionSubscriptions = args.hasRegionFormulaSubscriptionsForColumnAt?.(sheetId, col)
    if (hasRegionSubscriptions === undefined) {
      const sheetName = args.state.workbook.getSheetNameById(sheetId)
      hasRegionSubscriptions = sheetName ? args.hasRegionFormulaSubscriptionsForColumn(sheetName, col) : false
    }
    return (
      hasRegionSubscriptions ||
      (args.reverseState.reverseAggregateColumnEdges.get(aggregateColumnDependencyKey(sheetId, col))?.size ?? 0) > 0
    )
  }

  const hasTrackedColumnDependents = (sheetId: number, col: number): boolean =>
    hasTrackedExactLookupDependents(sheetId, col) ||
    hasTrackedSortedLookupDependents(sheetId, col) ||
    hasTrackedDirectRangeDependents(sheetId, col)

  const hasNoCellDependents = (cellIndex: number): boolean => {
    const slice = args.reverseState.reverseCellEdges[cellIndex]
    return slice === undefined || slice.len === 0 || slice.ptr < 0
  }

  const canSkipTerminalFormulaColumnVersion = (cellIndex: number): boolean => {
    const cellStore = args.state.workbook.cellStore
    const sheetId = cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(sheetId)
    const col =
      !sheet || sheet.structureVersion === 1
        ? cellStore.cols[cellIndex]
        : (sheet.logical.getCellVisiblePosition(cellIndex)?.col ?? cellStore.cols[cellIndex])
    return col !== undefined && !hasTrackedColumnDependents(sheetId, col)
  }

  const canSkipDirectFormulaColumnVersion = (cellIndex: number): boolean => {
    const cellStore = args.state.workbook.cellStore
    const sheetId = cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(sheetId)
    const col =
      !sheet || sheet.structureVersion === 1
        ? cellStore.cols[cellIndex]
        : (sheet.logical.getCellVisiblePosition(cellIndex)?.col ?? cellStore.cols[cellIndex])
    return col !== undefined && !hasTrackedColumnDependents(sheetId, col)
  }

  const canTrustPhysicalTrackedChangeSplit = (changed: U32, sheetId: number, split: number): boolean => {
    if (split <= 0 || split >= changed.length) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(sheetId)
    if (!sheet || sheet.structureVersion !== 1) {
      return false
    }
    const cellStore = args.state.workbook.cellStore
    const validateSlice = (start: number, end: number): boolean => {
      let previousRow = -1
      let previousCol = -1
      for (let index = start; index < end; index += 1) {
        const cellIndex = changed[index]!
        if (cellStore.sheetIds[cellIndex] !== sheetId) {
          return false
        }
        const row = cellStore.rows[cellIndex]
        const col = cellStore.cols[cellIndex]
        if (row === undefined || col === undefined || row < previousRow || (row === previousRow && col <= previousCol)) {
          return false
        }
        previousRow = row
        previousCol = col
      }
      return true
    }
    return validateSlice(0, split) && validateSlice(split, changed.length)
  }

  const getExactLookupImpactCache = (sheetId: number, col: number, caches: ExactLookupImpactCaches): ExactLookupImpactCache => {
    const cacheKey = lookupImpactCacheKey(sheetId, col)
    const cached = caches.get(cacheKey)
    if (cached !== undefined) {
      return cached
    }
    const dependents = args.getEntityDependents(makeExactLookupColumnEntity(sheetId, col))
    const entries: ExactLookupImpactEntry[] = []
    const operandKeys = new Set<string>()
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
      /* v8 ignore next -- defensive guard for stale exact-lookup reverse edges. */
      if (directLookup?.kind !== 'exact' && directLookup?.kind !== 'exact-uniform-numeric') {
        continue
      }
      const rowStart = directLookup.kind === 'exact' ? directLookup.prepared.rowStart : directLookup.rowStart
      const rowEnd = directLookup.kind === 'exact' ? directLookup.prepared.rowEnd : directLookup.rowEnd
      const operand = readCellValueForLookup(directLookup.operandCellIndex)
      const operandKey = normalizeExactLookupKey(operand.value, (id) => args.state.strings.get(id), operand.stringId)
      if (operandKey !== undefined) {
        operandKeys.add(operandKey)
      }
      entries.push({
        formulaCellIndex,
        rowStart,
        rowEnd,
        operandKey,
      })
    }
    const cache = { entries, operandKeys }
    caches.set(cacheKey, cache)
    return cache
  }

  const markSingleNumericExactLookupImpact = (
    sheetId: number,
    request: {
      sheetName: string
      row: number
      col: number
      oldValue: CellValue
      newValue: CellValue
    },
    formulaChangedCount: number,
  ): number | undefined => {
    const formulaCellIndex = args.getSingleEntityDependent(makeExactLookupColumnEntity(sheetId, request.col))
    if (formulaCellIndex === -1) {
      return formulaChangedCount
    }
    if (formulaCellIndex < 0) {
      return undefined
    }
    const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
    if (directLookup?.kind !== 'exact' && directLookup?.kind !== 'exact-uniform-numeric') {
      return undefined
    }
    const rowStart = directLookup.kind === 'exact' ? directLookup.prepared.rowStart : directLookup.rowStart
    const rowEnd = directLookup.kind === 'exact' ? directLookup.prepared.rowEnd : directLookup.rowEnd
    if (request.row < rowStart || request.row > rowEnd) {
      return formulaChangedCount
    }
    const oldNumeric = normalizeExactNumericValue(request.oldValue)
    const newNumeric = normalizeExactNumericValue(request.newValue)
    if (oldNumeric === undefined || newNumeric === undefined) {
      return undefined
    }
    const operandNumeric = normalizeExactNumericValue(readCellValueForLookup(directLookup.operandCellIndex).value)
    if (operandNumeric === undefined) {
      return undefined
    }
    if (!sameExactNumericValue(oldNumeric, operandNumeric) && !sameExactNumericValue(newNumeric, operandNumeric)) {
      return formulaChangedCount
    }
    return args.markFormulaChanged(formulaCellIndex, formulaChangedCount)
  }

  const cellTouchesPivotSource = (sheetName: string, row: number, col: number): boolean => {
    if (!args.state.workbook.hasPivots()) {
      return false
    }
    return args.state.workbook.listPivots().some((pivot) => {
      if (pivot.source.sheetName !== sheetName) {
        return false
      }
      const start = parseCellAddress(pivot.source.startAddress, pivot.source.sheetName)
      const end = parseCellAddress(pivot.source.endAddress, pivot.source.sheetName)
      return row >= start.row && row <= end.row && col >= start.col && col <= end.col
    })
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
    caches: ExactLookupImpactCaches,
  ): number => {
    const sheet = args.state.workbook.getSheet(request.sheetName)
    /* v8 ignore next -- defensive guard for stale lookup writes after sheet deletion. */
    if (!sheet) {
      return formulaChangedCount
    }
    const singleNumericImpact = markSingleNumericExactLookupImpact(sheet.id, request, formulaChangedCount)
    if (singleNumericImpact !== undefined) {
      return singleNumericImpact
    }
    const oldKey = normalizeExactLookupKey(request.oldValue, (id) => args.state.strings.get(id), request.oldStringId)
    const newKey = normalizeExactLookupKey(request.newValue, (id) => args.state.strings.get(id), request.newStringId)
    /* v8 ignore next -- error values cannot affect exact lookup matches. */
    if (oldKey === undefined && newKey === undefined) {
      return formulaChangedCount
    }
    const cache = getExactLookupImpactCache(sheet.id, request.col, caches)
    if (
      cache.entries.length === 0 ||
      ((oldKey === undefined || !cache.operandKeys.has(oldKey)) && (newKey === undefined || !cache.operandKeys.has(newKey)))
    ) {
      return formulaChangedCount
    }
    for (let index = 0; index < cache.entries.length; index += 1) {
      const entry = cache.entries[index]!
      /* v8 ignore next -- cached ranges are normally aligned with the lookup owner. */
      if (request.row < entry.rowStart || request.row > entry.rowEnd) {
        continue
      }
      /* v8 ignore next -- operand keys are prefiltered before scanning entries. */
      if (entry.operandKey === undefined || (entry.operandKey !== oldKey && entry.operandKey !== newKey)) {
        continue
      }
      formulaChangedCount = args.markFormulaChanged(entry.formulaCellIndex, formulaChangedCount)
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
    const dependents = args.getEntityDependents(makeSortedLookupColumnEntity(sheet.id, request.col))
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

  const noteExactLookupLiteralWriteWhenDirty = (
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
    caches: ExactLookupImpactCaches,
  ): number => {
    const nextFormulaChangedCount = markAffectedExactLookupDependents(request, formulaChangedCount, caches)
    args.noteExactLookupLiteralWrite(request)
    return nextFormulaChangedCount
  }

  const noteSortedLookupLiteralWriteWhenDirty = (
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
    const nextFormulaChangedCount = markAffectedApproximateLookupDependents(request, formulaChangedCount)
    args.noteSortedLookupLiteralWrite(request)
    return nextFormulaChangedCount
  }

  const collectAffectedDirectRangeDependents = (request: { sheetName: string; row: number; col: number }): number[] => {
    const sheetId = args.state.workbook.getSheet(request.sheetName)?.id
    const dependents = args.collectRegionFormulaDependentsForCell(request.sheetName, request.row, request.col)
    const affected: number[] = []
    const seen = new Set<number>()
    const consider = (formulaCellIndex: number): void => {
      if (seen.has(formulaCellIndex)) {
        return
      }
      seen.add(formulaCellIndex)
      const formula = args.state.formulas.get(formulaCellIndex)
      const directAggregate = formula?.directAggregate
      if (
        directAggregate &&
        directAggregate.sheetName === request.sheetName &&
        directAggregate.col === request.col &&
        request.row >= directAggregate.rowStart &&
        request.row <= directAggregate.rowEnd
      ) {
        affected.push(formulaCellIndex)
        return
      }
      const directCriteria = formula?.directCriteria
      if (directCriteria && directCriteriaTouchesPoint(directCriteria, request)) {
        affected.push(formulaCellIndex)
      }
    }
    for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
      consider(dependents[dependentIndex]!)
    }
    if (sheetId !== undefined) {
      args.reverseState.reverseAggregateColumnEdges.get(aggregateColumnDependencyKey(sheetId, request.col))?.forEach(consider)
    }
    return affected
  }

  const collectSingleAffectedDirectRangeDependent = (request: {
    sheetName: string
    row: number
    col: number
    sheetId?: number
  }): number => {
    const formulaCellIndex =
      request.sheetId !== undefined && args.collectSingleRegionFormulaDependentForCellAt
        ? args.collectSingleRegionFormulaDependentForCellAt(request.sheetId, request.row, request.col)
        : args.collectSingleRegionFormulaDependentForCell(request.sheetName, request.row, request.col)
    if (formulaCellIndex === -2) {
      return formulaCellIndex
    }
    if (formulaCellIndex >= 0) {
      const formula = args.state.formulas.get(formulaCellIndex)
      const directAggregate = formula?.directAggregate
      if (
        directAggregate &&
        directAggregate.sheetName === request.sheetName &&
        directAggregate.col === request.col &&
        request.row >= directAggregate.rowStart &&
        request.row <= directAggregate.rowEnd
      ) {
        return formulaCellIndex
      }
      const directCriteria = formula?.directCriteria
      if (directCriteria && directCriteriaTouchesPoint(directCriteria, request)) {
        return formulaCellIndex
      }
    }
    const sheetId = request.sheetId ?? args.state.workbook.getSheet(request.sheetName)?.id
    if (sheetId === undefined) {
      return -1
    }
    const aggregateDependents = args.reverseState.reverseAggregateColumnEdges.get(aggregateColumnDependencyKey(sheetId, request.col))
    if (!aggregateDependents || aggregateDependents.size === 0) {
      return -1
    }
    let singleAggregateDependent = -1
    for (const candidate of aggregateDependents) {
      const candidateFormula = args.state.formulas.get(candidate)
      const candidateDirectCriteria = candidateFormula?.directCriteria
      const candidateDirectAggregate = candidateFormula?.directAggregate
      const touches =
        (candidateDirectCriteria !== undefined && directCriteriaTouchesPoint(candidateDirectCriteria, request)) ||
        (candidateDirectAggregate !== undefined &&
          candidateDirectAggregate.sheetName === request.sheetName &&
          candidateDirectAggregate.col === request.col &&
          request.row >= candidateDirectAggregate.rowStart &&
          request.row <= candidateDirectAggregate.rowEnd)
      if (!touches) {
        continue
      }
      if (singleAggregateDependent !== -1 && singleAggregateDependent !== candidate) {
        return -2
      }
      singleAggregateDependent = candidate
    }
    return singleAggregateDependent
  }

  const canApplyDirectAggregateLiteralDeltaForRequest = (
    formulaCellIndex: number,
    request: { sheetName: string; row: number; col: number },
  ): boolean => {
    const formula = args.state.formulas.get(formulaCellIndex)
    const directAggregate = formula?.directAggregate
    return (
      formula !== undefined &&
      directAggregate?.aggregateKind === 'sum' &&
      formula.dependencyIndices.length === 0 &&
      directAggregate.sheetName === request.sheetName &&
      directAggregate.col === request.col &&
      request.row >= directAggregate.rowStart &&
      request.row <= directAggregate.rowEnd &&
      hasNoCellDependents(formulaCellIndex)
    )
  }

  const canApplyTrustedRangeDirectAggregateLiteralDelta = (formulaCellIndex: number): boolean => {
    const formula = args.state.formulas.get(formulaCellIndex)
    return (
      formula !== undefined &&
      formula.directAggregate?.aggregateKind === 'sum' &&
      formula.dependencyIndices.length === 0 &&
      hasNoCellDependents(formulaCellIndex)
    )
  }

  const collectSingleApplicableDirectAggregateDependent = (request: {
    sheetName: string
    row: number
    col: number
    sheetId?: number
  }): number => {
    const formulaCellIndex =
      request.sheetId !== undefined && args.collectSingleRegionFormulaDependentForCellAt
        ? args.collectSingleRegionFormulaDependentForCellAt(request.sheetId, request.row, request.col)
        : args.collectSingleRegionFormulaDependentForCell(request.sheetName, request.row, request.col)
    if (formulaCellIndex === -2) {
      return formulaCellIndex
    }
    if (formulaCellIndex >= 0) {
      return canApplyDirectAggregateLiteralDeltaForRequest(formulaCellIndex, request) ? formulaCellIndex : -2
    }
    const sheetId = request.sheetId ?? args.state.workbook.getSheet(request.sheetName)?.id
    if (sheetId === undefined) {
      return -1
    }
    const aggregateDependents = args.reverseState.reverseAggregateColumnEdges.get(aggregateColumnDependencyKey(sheetId, request.col))
    if (!aggregateDependents || aggregateDependents.size === 0) {
      return -1
    }
    let singleAggregateDependent = -1
    for (const candidate of aggregateDependents) {
      if (!canApplyDirectAggregateLiteralDeltaForRequest(candidate, request)) {
        continue
      }
      if (singleAggregateDependent !== -1 && singleAggregateDependent !== candidate) {
        return -2
      }
      singleAggregateDependent = candidate
    }
    return singleAggregateDependent
  }

  const markAffectedDirectRangeDependents = (
    request: {
      sheetName: string
      row: number
      col: number
      oldValue?: CellValue
      newValue?: CellValue
      inputCellIndex?: number
    },
    formulaChangedCount: number,
    postRecalcDirectFormulaIndices?: DirectFormulaIndexCollection,
  ): number => {
    const singleAffected = collectSingleAffectedDirectRangeDependent(request)
    const oldContribution = request.oldValue ? directAggregateNumericContribution(request.oldValue) : undefined
    const newContribution = request.newValue ? directAggregateNumericContribution(request.newValue) : undefined
    const contributionDelta = oldContribution === undefined || newContribution === undefined ? undefined : newContribution - oldContribution
    if (singleAffected === -1) {
      return formulaChangedCount
    }
    if (singleAffected >= 0) {
      const formula = args.state.formulas.get(singleAffected)
      const canUsePostRecalc =
        postRecalcDirectFormulaIndices !== undefined &&
        (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) &&
        args.getSingleEntityDependent(makeCellEntity(singleAffected)) === -1
      if (canUsePostRecalc) {
        if (
          contributionDelta !== undefined &&
          formula?.directAggregate?.aggregateKind === 'sum' &&
          formula.dependencyIndices.length === 0
        ) {
          postRecalcDirectFormulaIndices.addDelta(singleAffected, contributionDelta)
        } else if (formula?.directCriteria !== undefined) {
          const criteriaDelta = tryDirectCriteriaSumDelta(formula.directCriteria, request)
          if (criteriaDelta !== undefined) {
            postRecalcDirectFormulaIndices.addDelta(singleAffected, criteriaDelta)
          } else {
            postRecalcDirectFormulaIndices.add(singleAffected)
          }
        } else {
          postRecalcDirectFormulaIndices.add(singleAffected)
        }
        if (request.inputCellIndex !== undefined) {
          postRecalcDirectFormulaIndices.markDirectRangeInputCovered(request.inputCellIndex)
        }
        return formulaChangedCount
      }
      if (postRecalcDirectFormulaIndices && (formula?.dependencyIndices.length ?? 0) > 0) {
        postRecalcDirectFormulaIndices.add(singleAffected)
      }
      return args.markFormulaChanged(singleAffected, formulaChangedCount)
    }
    const dependents = collectAffectedDirectRangeDependents(request)
    const canUsePostRecalc =
      postRecalcDirectFormulaIndices !== undefined &&
      dependents.length > 0 &&
      dependents.length <= DIRECT_RANGE_POST_RECALC_LIMIT &&
      dependents.every((formulaCellIndex) => {
        const formula = args.state.formulas.get(formulaCellIndex)
        return (
          (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) &&
          args.getSingleEntityDependent(makeCellEntity(formulaCellIndex)) === -1
        )
      })
    const canUseDeltaPostRecalc =
      canUsePostRecalc &&
      contributionDelta !== undefined &&
      dependents.every((formulaCellIndex) => {
        const formula = args.state.formulas.get(formulaCellIndex)
        return formula?.directAggregate?.aggregateKind === 'sum' && formula.dependencyIndices.length === 0
      })
    let canUseCriteriaDeltaPostRecalc = false
    let criteriaDelta: number | undefined
    if (canUsePostRecalc && !canUseDeltaPostRecalc) {
      canUseCriteriaDeltaPostRecalc = true
      for (let index = 0; index < dependents.length; index += 1) {
        const formula = args.state.formulas.get(dependents[index]!)
        if (formula?.directCriteria === undefined) {
          canUseCriteriaDeltaPostRecalc = false
          break
        }
        const nextDelta = tryDirectCriteriaSumDelta(formula.directCriteria, request)
        if (nextDelta === undefined) {
          canUseCriteriaDeltaPostRecalc = false
          break
        }
        if (criteriaDelta === undefined) {
          criteriaDelta = nextDelta
        } else if (!Object.is(criteriaDelta, nextDelta)) {
          canUseCriteriaDeltaPostRecalc = false
          break
        }
      }
    }
    if (canUsePostRecalc && canUseDeltaPostRecalc) {
      postRecalcDirectFormulaIndices.appendConstantDelta(dependents, contributionDelta)
      if (request.inputCellIndex !== undefined) {
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(request.inputCellIndex)
      }
      return formulaChangedCount
    }
    if (canUsePostRecalc && canUseCriteriaDeltaPostRecalc && criteriaDelta !== undefined) {
      postRecalcDirectFormulaIndices.appendConstantDelta(dependents, criteriaDelta)
      if (request.inputCellIndex !== undefined) {
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(request.inputCellIndex)
      }
      return formulaChangedCount
    }
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (canUsePostRecalc) {
        postRecalcDirectFormulaIndices.add(formulaCellIndex)
        continue
      }
      if (postRecalcDirectFormulaIndices && (args.state.formulas.get(formulaCellIndex)?.dependencyIndices.length ?? 0) > 0) {
        postRecalcDirectFormulaIndices.add(formulaCellIndex)
      }
      formulaChangedCount = args.markFormulaChanged(formulaCellIndex, formulaChangedCount)
    }
    if (canUsePostRecalc && request.inputCellIndex !== undefined) {
      postRecalcDirectFormulaIndices.markDirectRangeInputCovered(request.inputCellIndex)
    }
    return formulaChangedCount
  }

  const applyDirectFormulaNumericDelta = (cellIndex: number, delta: number): boolean => {
    const cellStore = args.state.workbook.cellStore
    if (cellStore.tags[cellIndex] !== ValueTag.Number) {
      return false
    }
    const beforeNumber = cellStore.numbers[cellIndex] ?? 0
    const nextNumber = beforeNumber + delta
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.numbers[cellIndex] = nextNumber
    cellStore.stringIds[cellIndex] = 0
    cellStore.errors[cellIndex] = 0
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    if (cellStore.onSetValue) {
      cellStore.onSetValue(cellIndex)
    } else if (!Object.is(beforeNumber, nextNumber)) {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
    return true
  }

  const applyTerminalDirectFormulaNumericDelta = (cellIndex: number, delta: number): boolean => {
    return applyTerminalDirectFormulaNumericDeltaAndReturn(cellIndex, delta) !== undefined
  }

  const applyTerminalDirectFormulaNumericDeltaAndReturn = (cellIndex: number, delta: number): number | undefined => {
    const cellStore = args.state.workbook.cellStore
    if (cellStore.tags[cellIndex] !== ValueTag.Number) {
      return undefined
    }
    const nextNumber = (cellStore.numbers[cellIndex] ?? 0) + delta
    const flags = cellStore.flags[cellIndex] ?? 0
    if ((flags & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
      cellStore.flags[cellIndex] = flags & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    }
    cellStore.numbers[cellIndex] = nextNumber
    if ((cellStore.stringIds[cellIndex] ?? 0) !== 0) {
      cellStore.stringIds[cellIndex] = 0
    }
    if ((cellStore.errors[cellIndex] ?? 0) !== 0) {
      cellStore.errors[cellIndex] = 0
    }
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    return nextNumber
  }

  const tryApplyDirectFormulaDeltas = (collection: DirectFormulaIndexCollection, captureChanged = true): U32 | undefined => {
    if (!collection.hasCompleteDeltas()) {
      return undefined
    }
    const cellStore = args.state.workbook.cellStore
    const changed = captureChanged ? new Uint32Array(collection.size) : EMPTY_CHANGED_CELLS
    let directAggregateDeltaApplicationCount = 0
    let directScalarDeltaApplicationCount = 0
    let canUseTerminalFormulaWrites = true
    for (let index = 0; index < collection.size; index += 1) {
      const cellIndex = collection.getCellIndexAt(index)
      if (
        ((cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0 ||
        cellStore.tags[cellIndex] !== ValueTag.Number ||
        collection.getDeltaAt(index) === undefined
      ) {
        return undefined
      }
      const formula = args.state.formulas.get(cellIndex)
      if (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) {
        directAggregateDeltaApplicationCount += 1
      }
      if (formula?.directScalar !== undefined) {
        directScalarDeltaApplicationCount += 1
      }
      if (canUseTerminalFormulaWrites && !canSkipTerminalFormulaColumnVersion(cellIndex)) {
        canUseTerminalFormulaWrites = false
      }
      if (captureChanged) {
        changed[index] = cellIndex
      }
    }
    const applyDeltas = (): void => {
      for (let index = 0; index < collection.size; index += 1) {
        const cellIndex = captureChanged ? changed[index]! : collection.getCellIndexAt(index)
        const applied = canUseTerminalFormulaWrites
          ? applyTerminalDirectFormulaNumericDelta(cellIndex, collection.getDeltaAt(index)!)
          : applyDirectFormulaNumericDelta(cellIndex, collection.getDeltaAt(index)!)
        if (!applied) {
          throw new Error('Failed to apply direct formula delta')
        }
      }
    }
    if (canUseTerminalFormulaWrites) {
      applyDeltas()
    } else {
      args.state.workbook.withBatchedColumnVersionUpdates(applyDeltas)
    }
    if (directAggregateDeltaApplicationCount > 0) {
      addEngineCounter(args.state.counters, 'directAggregateDeltaApplications', directAggregateDeltaApplicationCount)
    }
    if (directScalarDeltaApplicationCount > 0) {
      addEngineCounter(args.state.counters, 'directScalarDeltaApplications', directScalarDeltaApplicationCount)
    }
    return changed
  }

  const tryApplyDirectScalarDeltas = (collection: DirectFormulaIndexCollection, captureChanged = true): U32 | undefined => {
    const constantDelta = collection.getConstantScalarDelta()
    if (constantDelta === undefined && !collection.hasCompleteScalarDeltas()) {
      return undefined
    }
    const cellStore = args.state.workbook.cellStore
    const changed = captureChanged ? new Uint32Array(collection.size) : EMPTY_CHANGED_CELLS
    const hasValidatedTerminalWrites = collection.hasValidatedScalarDeltaCells()
    let canUseTerminalFormulaWrites = hasValidatedTerminalWrites
    if (!hasValidatedTerminalWrites) {
      canUseTerminalFormulaWrites = true
      for (let index = 0; index < collection.size; index += 1) {
        const cellIndex = collection.getCellIndexAt(index)
        if (((cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0 || cellStore.tags[cellIndex] !== ValueTag.Number) {
          return undefined
        }
        if (canUseTerminalFormulaWrites && !canSkipDirectFormulaColumnVersion(cellIndex)) {
          canUseTerminalFormulaWrites = false
        }
        if (captureChanged) {
          changed[index] = cellIndex
        }
      }
    }
    const applyDeltas = (): void => {
      for (let index = 0; index < collection.size; index += 1) {
        const cellIndex = hasValidatedTerminalWrites || !captureChanged ? collection.getCellIndexAt(index) : changed[index]!
        if (hasValidatedTerminalWrites && captureChanged) {
          changed[index] = cellIndex
        }
        const delta = constantDelta ?? collection.getScalarDeltaAt(index)
        if (delta === undefined) {
          throw new Error('Missing direct scalar delta')
        }
        if (canUseTerminalFormulaWrites) {
          if (!applyTerminalDirectFormulaNumericDelta(cellIndex, delta)) {
            throw new Error('Failed to apply direct scalar delta')
          }
        } else {
          const beforeNumber = cellStore.numbers[cellIndex] ?? 0
          const nextNumber = beforeNumber + delta
          cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
          cellStore.numbers[cellIndex] = nextNumber
          cellStore.stringIds[cellIndex] = 0
          cellStore.errors[cellIndex] = 0
          cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
          if (cellStore.onSetValue) {
            cellStore.onSetValue(cellIndex)
          } else if (!Object.is(beforeNumber, nextNumber)) {
            args.state.workbook.notifyCellValueWritten(cellIndex)
          }
        }
      }
    }
    if (canUseTerminalFormulaWrites) {
      applyDeltas()
    } else {
      args.state.workbook.withBatchedColumnVersionUpdates(applyDeltas)
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', collection.size)
    return changed
  }

  const tryApplySinglePostRecalcDirectFormula = (
    collection: DirectFormulaIndexCollection,
    didRunRecalc: boolean,
    counts: DirectFormulaMetricCounts,
    captureChanged = true,
  ): U32 | undefined => {
    if (didRunRecalc || collection.size !== 1) {
      return undefined
    }
    const cellIndex = collection.getCellIndexAt(0)
    if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
      return undefined
    }
    const currentResult = collection.getCurrentResultAt(0)
    if (currentResult !== undefined) {
      return applyDirectFormulaCurrentResult(cellIndex, currentResult)
        ? captureChanged
          ? Uint32Array.of(cellIndex)
          : EMPTY_CHANGED_CELLS
        : undefined
    }
    const delta = collection.getDeltaAt(0)
    if (delta !== undefined) {
      if (!applyDirectFormulaNumericDelta(cellIndex, delta)) {
        return undefined
      }
      const formula = args.state.formulas.get(cellIndex)
      if (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) {
        addEngineCounter(args.state.counters, 'directAggregateDeltaApplications')
      }
      if (formula?.directScalar !== undefined) {
        addEngineCounter(args.state.counters, 'directScalarDeltaApplications')
      }
      return captureChanged ? Uint32Array.of(cellIndex) : EMPTY_CHANGED_CELLS
    }
    const formula = args.state.formulas.get(cellIndex)
    if (
      formula?.directScalar !== undefined &&
      !formula.compiled.producesSpill &&
      applyDirectScalarCurrentValue(cellIndex, formula.directScalar)
    ) {
      countPostRecalcDirectFormulaMetric(cellIndex, counts)
      return captureChanged ? Uint32Array.of(cellIndex) : EMPTY_CHANGED_CELLS
    }
    return undefined
  }

  const canApplyDirectAggregateLiteralDelta = (formulaCellIndex: number): boolean => {
    const formula = args.state.formulas.get(formulaCellIndex)
    return (
      formula?.directAggregate?.aggregateKind === 'sum' && formula.dependencyIndices.length === 0 && hasNoCellDependents(formulaCellIndex)
    )
  }

  const tryApplySingleDirectAggregateLiteralMutationFastPath = (request: {
    existingIndex: number
    sheetId?: number
    sheetName: string
    row: number
    col: number
    value: LiteralInput
    delta: number
    emitTracked: boolean
    singleRangeEntityDependent?: number
  }): EngineExistingNumericCellMutationResult | null => {
    let singleAffected = -2
    if (request.singleRangeEntityDependent !== undefined) {
      const rangeDependent = args.getSingleEntityDependent(request.singleRangeEntityDependent)
      if (rangeDependent < -1) {
        return null
      }
      if (rangeDependent >= 0) {
        singleAffected = rangeDependent
      }
    }
    if (singleAffected >= 0 && !canApplyDirectAggregateLiteralDeltaForRequest(singleAffected, request)) {
      return null
    }
    if (singleAffected < -1) {
      singleAffected = collectSingleApplicableDirectAggregateDependent({
        sheetName: request.sheetName,
        ...(request.sheetId === undefined ? {} : { sheetId: request.sheetId }),
        row: request.row,
        col: request.col,
      })
    }
    if (singleAffected === -1) {
      if (typeof request.value === 'number') {
        writeNumericLiteralToExistingCell(request.existingIndex, request.value)
      } else {
        writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
      }
      deferSingleCellKernelSync(request.existingIndex)
      const lastMetrics = makeSingleLiteralSkipMetrics()
      args.state.setLastMetrics(lastMetrics)
      if (request.emitTracked) {
        const changed = Uint32Array.of(request.existingIndex)
        args.state.events.emitTracked({
          kind: 'batch',
          invalidation: 'cells',
          changedCellIndices: changed,
          invalidatedRanges: [],
          invalidatedRows: [],
          invalidatedColumns: [],
          metrics: lastMetrics,
          explicitChangedCount: 1,
        })
        return makeExistingNumericMutationResult(changed, 1)
      }
      return makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
    }

    let singleAggregateCellIndex = -1
    let affected: readonly number[] | undefined
    if (singleAffected >= 0) {
      singleAggregateCellIndex = singleAffected
    } else {
      const collected = collectAffectedDirectRangeDependents({
        sheetName: request.sheetName,
        row: request.row,
        col: request.col,
      })
      if (collected.length === 0 || collected.length > DIRECT_RANGE_POST_RECALC_LIMIT) {
        return null
      }
      for (let index = 0; index < collected.length; index += 1) {
        if (!canApplyDirectAggregateLiteralDelta(collected[index]!)) {
          return null
        }
      }
      affected = collected
    }
    const affectedCount = singleAggregateCellIndex >= 0 ? 1 : (affected?.length ?? 0)
    const sharesSingleAggregateVersionColumn =
      singleAggregateCellIndex >= 0 &&
      request.sheetId !== undefined &&
      args.state.workbook.cellStore.sheetIds[singleAggregateCellIndex] === request.sheetId &&
      args.state.workbook.cellStore.cols[singleAggregateCellIndex] === request.col
    const shouldBatchColumnVersions =
      affectedCount > 1 ||
      (singleAggregateCellIndex >= 0 &&
        (request.sheetId === undefined
          ? cellsShareVersionColumn(request.existingIndex, singleAggregateCellIndex)
          : sharesSingleAggregateVersionColumn))

    let singleAggregateNumericValue: number | undefined
    if (!shouldBatchColumnVersions && singleAggregateCellIndex >= 0) {
      if (typeof request.value === 'number') {
        writeNumericLiteralToExistingCell(request.existingIndex, request.value)
      } else {
        writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
      }
      singleAggregateNumericValue = applyTerminalDirectFormulaNumericDeltaAndReturn(singleAggregateCellIndex, request.delta)
      if (singleAggregateNumericValue === undefined) {
        throw new Error('Failed to apply direct aggregate delta')
      }
    } else {
      withOptionalColumnVersionBatch(shouldBatchColumnVersions, () => {
        if (typeof request.value === 'number') {
          writeNumericLiteralToExistingCell(request.existingIndex, request.value)
        } else {
          writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
        }
        if (singleAggregateCellIndex >= 0) {
          if (!applyDirectFormulaNumericDelta(singleAggregateCellIndex, request.delta)) {
            throw new Error('Failed to apply direct aggregate delta')
          }
        } else {
          for (let index = 0; index < affected!.length; index += 1) {
            if (!applyDirectFormulaNumericDelta(affected![index]!, request.delta)) {
              throw new Error('Failed to apply direct aggregate delta')
            }
          }
        }
      })
    }
    addEngineCounter(args.state.counters, 'directAggregateDeltaApplications', affectedCount)
    addEngineCounter(args.state.counters, 'directAggregateDeltaOnlyRecalcSkips')
    deferSingleCellKernelSync(request.existingIndex)
    const lastMetrics = makeSingleLiteralSkipMetrics()
    args.state.setLastMetrics(lastMetrics)
    if (request.emitTracked) {
      const changed =
        singleAggregateCellIndex >= 0
          ? Uint32Array.of(request.existingIndex, singleAggregateCellIndex)
          : affectedCount === 0
            ? Uint32Array.of(request.existingIndex)
            : composeSingleDisjointExplicitEventChanges(request.existingIndex, Uint32Array.from(affected!))
      if (singleAggregateCellIndex >= 0 && changed.length > 4 && request.sheetId !== undefined) {
        tagTrustedPhysicalTrackedChanges(changed, request.sheetId, 1)
      }
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount: 1,
      })
      return makeExistingNumericMutationResult(changed, 1)
    }
    if (singleAggregateCellIndex >= 0) {
      return makeCompactExistingNumericMutationResult(request.existingIndex, singleAggregateCellIndex, 1, singleAggregateNumericValue)
    }
    if (affectedCount === 0) {
      return makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
    }
    const changed = composeSingleDisjointExplicitEventChanges(request.existingIndex, Uint32Array.from(affected!))
    if (singleAggregateCellIndex >= 0 && changed.length > 4 && request.sheetId !== undefined) {
      tagTrustedPhysicalTrackedChanges(changed, request.sheetId, 1)
    }
    return makeExistingNumericMutationResult(changed, 1)
  }

  const tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation = (request: {
    existingIndex: number
    rangeEntityDependent: number
    sheet: SheetRecord
    sheetId: number
    col: number
    value: number
    delta: number
    hasExactLookupDependents: boolean
    hasSortedLookupDependents: boolean
  }): EngineExistingNumericCellMutationResult | null => {
    if (request.hasExactLookupDependents || request.hasSortedLookupDependents) {
      return null
    }
    const formulaCellIndex = args.getSingleEntityDependent(request.rangeEntityDependent)
    if (formulaCellIndex < 0 || !canApplyTrustedRangeDirectAggregateLiteralDelta(formulaCellIndex)) {
      return null
    }
    writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.sheet, request.col, request.value)
    const aggregateNumericValue = applyTerminalDirectFormulaNumericDeltaAndReturn(formulaCellIndex, request.delta)
    if (aggregateNumericValue === undefined) {
      throw new Error('Failed to apply direct aggregate delta')
    }
    args.state.counters.directAggregateDeltaApplications += 1
    args.state.counters.directAggregateDeltaOnlyRecalcSkips += 1
    deferSingleCellKernelSync(request.existingIndex)
    args.state.setLastMetrics(makeSingleLiteralSkipMetrics())
    const cellStore = args.state.workbook.cellStore
    return makeCompactExistingNumericMutationResult(request.existingIndex, formulaCellIndex, 1, aggregateNumericValue, {
      row: cellStore.rows[formulaCellIndex] ?? 0,
      col: cellStore.cols[formulaCellIndex] ?? 0,
    })
  }

  const tryApplyTrustedDirectScalarClosureExistingNumericMutation = (request: {
    existingIndex: number
    sheet: SheetRecord
    sheetId: number
    col: number
    value: number
    oldNumber: number
    hasTrackedEventListeners: boolean
  }): EngineExistingNumericCellMutationResult | null => {
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()
    const oldValue: CellValue = { tag: ValueTag.Number, value: request.oldNumber }
    const newValue: CellValue = { tag: ValueTag.Number, value: request.value }
    if (!tryMarkDirectScalarLinearDeltaClosure(request.existingIndex, oldValue, newValue, postRecalcDirectFormulaIndices)) {
      return null
    }
    if (!hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices)) {
      return null
    }
    countDirectFormulaDeltaSkip(args.state.formulas, postRecalcDirectFormulaIndices, args.state.counters)
    writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.sheet, request.col, request.value)
    const directChanged = tryApplyDirectScalarDeltas(postRecalcDirectFormulaIndices, true)
    if (directChanged === undefined) {
      throw new Error('Failed to apply direct scalar closure delta')
    }
    deferSingleCellKernelSync(request.existingIndex)
    const lastMetrics = makeSingleLiteralSkipMetrics()
    args.state.setLastMetrics(lastMetrics)
    const changed = composeSingleDisjointExplicitEventChanges(request.existingIndex, directChanged)
    if (changed.length > 4 && canTrustPhysicalTrackedChangeSplit(changed, request.sheetId, 1)) {
      tagTrustedPhysicalTrackedChanges(changed, request.sheetId, 1)
    }
    if (request.hasTrackedEventListeners) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount: 1,
      })
    }
    return makeExistingNumericMutationResult(changed, 1)
  }

  const tryApplySingleDirectScalarLiteralMutationWithoutEvents = (request: {
    existingIndex: number
    value: LiteralInput
    oldNumber: number
    newNumber: number
  }): boolean => {
    const dependencyEntity = makeCellEntity(request.existingIndex)
    const singleDependent = args.getSingleEntityDependent(dependencyEntity)
    if (singleDependent === -1) {
      return false
    }

    let singleFormulaCellIndex = -1
    let dependents: U32 | undefined
    if (singleDependent >= 0) {
      singleFormulaCellIndex = singleDependent
    } else {
      dependents = args.getEntityDependents(dependencyEntity)
      if (dependents.length === 0) {
        return false
      }
    }

    let commonDelta = 0
    let hasCommonDelta = false
    const validateDependent = (formulaCellIndex: number): boolean => {
      if (!canUseDirectFormulaPostRecalc(formulaCellIndex)) {
        return false
      }
      const formula = args.state.formulas.get(formulaCellIndex)
      const delta =
        formula?.directScalar === undefined
          ? undefined
          : tryDirectScalarNumericDeltaFromNumbers(formula.directScalar, request.existingIndex, request.oldNumber, request.newNumber)
      if (
        delta === undefined ||
        ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0 ||
        args.state.workbook.cellStore.tags[formulaCellIndex] !== ValueTag.Number
      ) {
        return false
      }
      if (!hasCommonDelta) {
        commonDelta = delta
        hasCommonDelta = true
        return true
      }
      return Object.is(commonDelta, delta)
    }

    if (singleFormulaCellIndex >= 0) {
      if (!validateDependent(singleFormulaCellIndex)) {
        return false
      }
    } else {
      for (let index = 0; index < dependents!.length; index += 1) {
        if (!validateDependent(dependents![index]!)) {
          return false
        }
      }
    }

    args.state.workbook.withBatchedColumnVersionUpdates(() => {
      writeLiteralToCellStore(args.state.workbook.cellStore, request.existingIndex, request.value, args.state.strings)
      args.state.workbook.notifyCellValueWritten(request.existingIndex)
      if (singleFormulaCellIndex >= 0) {
        if (!applyDirectFormulaNumericDelta(singleFormulaCellIndex, commonDelta)) {
          throw new Error('Failed to apply direct scalar delta')
        }
      } else {
        for (let index = 0; index < dependents!.length; index += 1) {
          if (!applyDirectFormulaNumericDelta(dependents![index]!, commonDelta)) {
            throw new Error('Failed to apply direct scalar delta')
          }
        }
      }
    })
    const applicationCount = singleFormulaCellIndex >= 0 ? 1 : dependents!.length
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', applicationCount)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    deferSingleCellKernelSync(request.existingIndex)
    args.state.setLastMetrics(makeSingleLiteralSkipMetrics())
    return true
  }

  const tryApplySingleDirectLookupOperandMutationFastPath = (request: {
    existingIndex: number
    formulaCellIndex: number
    value: LiteralInput
    exactLookupValue: number | undefined
    approximateLookupValue: number | undefined
    emitTracked: boolean
    lookupSheetHint?: SheetRecord | undefined
    trustedInputSheet?: SheetRecord | undefined
    trustedInputCol?: number | undefined
  }): EngineExistingNumericCellMutationResult | null => {
    const formulaCellIndex = request.formulaCellIndex
    if (formulaCellIndex < 0 || !hasNoCellDependents(formulaCellIndex)) {
      return null
    }
    const formula = args.state.formulas.get(formulaCellIndex)
    const directLookup = formula?.directLookup
    const numericResult = tryDirectUniformLookupNumericResultFromDescriptor(
      directLookup,
      request.exactLookupValue,
      request.approximateLookupValue,
      request.lookupSheetHint,
    )
    if (numericResult !== undefined) {
      const resultChanged = !directScalarNumericResultMatchesCell(formulaCellIndex, numericResult)
      const writeInput = (): void => {
        if (typeof request.value === 'number' && request.trustedInputSheet !== undefined && request.trustedInputCol !== undefined) {
          writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.trustedInputSheet, request.trustedInputCol, request.value)
        } else if (typeof request.value === 'number') {
          writeNumericLiteralToExistingCell(request.existingIndex, request.value)
        } else {
          writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
        }
      }
      const apply = (): void => {
        writeInput()
        if (resultChanged) {
          applyTerminalDirectFormulaNumericResult(formulaCellIndex, numericResult)
        }
      }
      if (resultChanged && cellsShareVersionColumn(request.existingIndex, formulaCellIndex)) {
        withOptionalColumnVersionBatch(true, apply)
      } else {
        apply()
      }
      addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
      deferSingleCellKernelSync(request.existingIndex)
      const lastMetrics = makeSingleLiteralSkipMetrics()
      args.state.setLastMetrics(lastMetrics)
      if (request.emitTracked) {
        const changedCellIndices = resultChanged
          ? Uint32Array.of(request.existingIndex, formulaCellIndex)
          : Uint32Array.of(request.existingIndex)
        args.state.events.emitTracked({
          kind: 'batch',
          invalidation: 'cells',
          changedCellIndices,
          invalidatedRanges: [],
          invalidatedRows: [],
          invalidatedColumns: [],
          metrics: lastMetrics,
          explicitChangedCount: 1,
        })
      }
      return makeCompactExistingNumericMutationResult(
        request.existingIndex,
        resultChanged ? formulaCellIndex : undefined,
        1,
        resultChanged ? numericResult : undefined,
      )
    }
    let result: DirectScalarCurrentOperand | undefined
    if (directLookup?.kind === 'exact-uniform-numeric') {
      const lookupSheet = lookupSheetForUniformLookup(directLookup, request.lookupSheetHint)
      if (request.exactLookupValue !== undefined && directLookupVersionMatches(lookupSheet, directLookup)) {
        result = exactUniformLookupCurrentResult(directLookup, request.exactLookupValue)
      }
    } else if (directLookup?.kind === 'approximate-uniform-numeric') {
      const lookupSheet = lookupSheetForUniformLookup(directLookup, request.lookupSheetHint)
      if (request.approximateLookupValue !== undefined && directLookupVersionMatches(lookupSheet, directLookup)) {
        result = approximateUniformLookupCurrentResult(directLookup, request.approximateLookupValue)
      }
    }
    if (result === undefined) {
      return null
    }
    const resultChanged = !directScalarCurrentResultMatchesCell(formulaCellIndex, result)
    const writeInput = (): void => {
      if (typeof request.value === 'number' && request.trustedInputSheet !== undefined && request.trustedInputCol !== undefined) {
        writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.trustedInputSheet, request.trustedInputCol, request.value)
      } else if (typeof request.value === 'number') {
        writeNumericLiteralToExistingCell(request.existingIndex, request.value)
      } else {
        writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
      }
    }
    const apply = (): void => {
      writeInput()
      if (resultChanged && !applyDirectFormulaCurrentResult(formulaCellIndex, result)) {
        throw new Error('Failed to apply direct lookup result')
      }
    }
    if (resultChanged && cellsShareVersionColumn(request.existingIndex, formulaCellIndex)) {
      withOptionalColumnVersionBatch(true, apply)
    } else {
      apply()
    }
    addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
    deferSingleCellKernelSync(request.existingIndex)
    const lastMetrics = makeSingleLiteralSkipMetrics()
    args.state.setLastMetrics(lastMetrics)
    if (request.emitTracked) {
      const changedCellIndices = resultChanged
        ? Uint32Array.of(request.existingIndex, formulaCellIndex)
        : Uint32Array.of(request.existingIndex)
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount: 1,
      })
    }
    return makeCompactExistingNumericMutationResult(request.existingIndex, resultChanged ? formulaCellIndex : undefined, 1)
  }

  const tryApplySingleDirectFormulaLiteralMutationWithoutEvents = (request: {
    existingIndex: number
    formulaCellIndex: number
    value: LiteralInput
    oldNumber: number
    newNumber: number
    exactLookupValue: number | undefined
    approximateLookupValue: number | undefined
  }): boolean => {
    const formulaCellIndex = request.formulaCellIndex
    if (formulaCellIndex < 0) {
      return false
    }
    const formula = args.state.formulas.get(formulaCellIndex)
    if (!formula || ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
      return false
    }
    if (formula.directLookup !== undefined) {
      if (!hasNoCellDependents(formulaCellIndex)) {
        return false
      }
      const numericResult = tryDirectUniformLookupNumericResultFromDescriptor(
        formula.directLookup,
        request.exactLookupValue,
        request.approximateLookupValue,
      )
      if (numericResult !== undefined) {
        const resultChanged = !directScalarNumericResultMatchesCell(formulaCellIndex, numericResult)
        withOptionalColumnVersionBatch(resultChanged && cellsShareVersionColumn(request.existingIndex, formulaCellIndex), () => {
          writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
          if (resultChanged) {
            applyTerminalDirectFormulaNumericResult(formulaCellIndex, numericResult)
          }
        })
        addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
        deferSingleCellKernelSync(request.existingIndex)
        args.state.setLastMetrics(makeSingleLiteralSkipMetrics())
        return true
      }
      const result = tryDirectUniformLookupCurrentResultFromNumeric(
        formulaCellIndex,
        request.exactLookupValue,
        request.approximateLookupValue,
      )
      if (result === undefined) {
        return false
      }
      const resultChanged = !directScalarCurrentResultMatchesCell(formulaCellIndex, result)
      withOptionalColumnVersionBatch(resultChanged && cellsShareVersionColumn(request.existingIndex, formulaCellIndex), () => {
        writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
        if (resultChanged && !applyDirectFormulaCurrentResult(formulaCellIndex, result)) {
          throw new Error('Failed to apply direct lookup result')
        }
      })
      addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
      deferSingleCellKernelSync(request.existingIndex)
      args.state.setLastMetrics(makeSingleLiteralSkipMetrics())
      return true
    }
    if (!canUseDirectFormulaPostRecalc(formulaCellIndex)) {
      return false
    }
    if (formula.directScalar === undefined || args.state.workbook.cellStore.tags[formulaCellIndex] !== ValueTag.Number) {
      return false
    }
    const delta = tryDirectScalarNumericDeltaFromNumbers(formula.directScalar, request.existingIndex, request.oldNumber, request.newNumber)
    if (delta === undefined) {
      return false
    }
    withOptionalColumnVersionBatch(cellsShareVersionColumn(request.existingIndex, formulaCellIndex), () => {
      writeLiteralToCellStore(args.state.workbook.cellStore, request.existingIndex, request.value, args.state.strings)
      args.state.workbook.notifyCellValueWritten(request.existingIndex)
      if (!applyDirectFormulaNumericDelta(formulaCellIndex, delta)) {
        throw new Error('Failed to apply direct scalar delta')
      }
    })
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications')
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    deferSingleCellKernelSync(request.existingIndex)
    args.state.setLastMetrics(makeSingleLiteralSkipMetrics())
    return true
  }

  const tryApplySingleKernelSyncOnlyLiteralMutationFastPath = (request: {
    existingIndex: number
    value: LiteralInput
    emitTracked: boolean
    afterWrite?: () => void
  }): boolean => {
    writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
    request.afterWrite?.()
    addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
    deferSingleCellKernelSync(request.existingIndex)
    const lastMetrics = makeSingleLiteralSkipMetrics()
    args.state.setLastMetrics(lastMetrics)
    if (request.emitTracked) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: Uint32Array.of(request.existingIndex),
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount: 1,
      })
    }
    return true
  }

  const applyDirectFormulaCurrentResult = (cellIndex: number, result: DirectScalarCurrentOperand): boolean => {
    const cellStore = args.state.workbook.cellStore
    const beforeTag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    const beforeNumber = cellStore.numbers[cellIndex] ?? 0
    const beforeError = (cellStore.errors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    cellStore.stringIds[cellIndex] = 0
    if (result.kind === 'number') {
      cellStore.tags[cellIndex] = ValueTag.Number
      cellStore.numbers[cellIndex] = result.value
      cellStore.errors[cellIndex] = ErrorCode.None
      if (cellStore.onSetValue) {
        cellStore.onSetValue(cellIndex)
      } else if (beforeTag !== ValueTag.Number || !Object.is(beforeNumber, result.value)) {
        args.state.workbook.notifyCellValueWritten(cellIndex)
      }
      return true
    }
    cellStore.tags[cellIndex] = ValueTag.Error
    cellStore.numbers[cellIndex] = 0
    cellStore.errors[cellIndex] = result.code
    if (cellStore.onSetValue) {
      cellStore.onSetValue(cellIndex)
    } else if (beforeTag !== ValueTag.Error || beforeError !== result.code) {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
    return true
  }

  const applyDirectFormulaNumericResult = (cellIndex: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    cellStore.stringIds[cellIndex] = 0
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.numbers[cellIndex] = value
    cellStore.errors[cellIndex] = ErrorCode.None
    if (cellStore.onSetValue) {
      cellStore.onSetValue(cellIndex)
    } else {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
  }

  const applyTerminalDirectFormulaNumericResult = (cellIndex: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    cellStore.stringIds[cellIndex] = 0
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.numbers[cellIndex] = value
    cellStore.errors[cellIndex] = ErrorCode.None
  }

  const writeNumericLiteralToCellStore = (cellIndex: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    cellStore.stringIds[cellIndex] = 0
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.numbers[cellIndex] = value
    cellStore.errors[cellIndex] = ErrorCode.None
  }

  const readDirectScalarCurrentOperand = (operand: RuntimeDirectScalarOperand): DirectScalarCurrentOperand | undefined => {
    switch (operand.kind) {
      case 'literal-number':
        return { kind: 'number', value: operand.value }
      case 'error':
        return { kind: 'error', code: operand.code }
      case 'cell': {
        const cellStore = args.state.workbook.cellStore
        const tag = (cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
        switch (tag) {
          case ValueTag.Number:
            return { kind: 'number', value: cellStore.numbers[operand.cellIndex] ?? 0 }
          case ValueTag.Boolean:
            return { kind: 'number', value: (cellStore.numbers[operand.cellIndex] ?? 0) !== 0 ? 1 : 0 }
          case ValueTag.Empty:
            return { kind: 'number', value: 0 }
          case ValueTag.Error:
            return { kind: 'error', code: (cellStore.errors[operand.cellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
          case ValueTag.String:
            return { kind: 'error', code: ErrorCode.Value }
        }
      }
    }
  }

  const evaluateDirectScalarCurrentValue = (directScalar: RuntimeDirectScalarDescriptor): DirectScalarCurrentOperand | undefined => {
    if (directScalar.kind === 'abs') {
      const operand = readDirectScalarCurrentOperand(directScalar.operand)
      return operand?.kind === 'number' ? { kind: 'number', value: Math.abs(operand.value) } : operand
    }
    const left = readDirectScalarCurrentOperand(directScalar.left)
    const right = readDirectScalarCurrentOperand(directScalar.right)
    if (!left || !right) {
      return undefined
    }
    if (left.kind === 'error') {
      return left
    }
    if (right.kind === 'error') {
      return right
    }
    switch (directScalar.operator) {
      case '+':
        return { kind: 'number', value: left.value + right.value }
      case '-':
        return { kind: 'number', value: left.value - right.value }
      case '*':
        return { kind: 'number', value: left.value * right.value }
      case '/':
        return right.value === 0 ? { kind: 'error', code: ErrorCode.Div0 } : { kind: 'number', value: left.value / right.value }
    }
  }

  const applyDirectScalarCurrentValue = (cellIndex: number, directScalar: RuntimeDirectScalarDescriptor): boolean => {
    const result = evaluateDirectScalarCurrentValue(directScalar)
    if (!result) {
      return false
    }
    return applyDirectFormulaCurrentResult(cellIndex, result)
  }

  const tryApplyFormulaReplacementAsDirectScalarDeltaRoot = (request: {
    cellIndex: number
    oldNumber: number | undefined
    changedTopology: boolean
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
    postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts
  }): boolean => {
    if (request.changedTopology || request.oldNumber === undefined) {
      return false
    }
    const formula = args.state.formulas.get(request.cellIndex)
    if (
      !formula ||
      formula.directScalar === undefined ||
      formula.compiled.volatile ||
      formula.compiled.producesSpill ||
      ((args.state.workbook.cellStore.flags[request.cellIndex] ?? 0) & CellFlags.InCycle) !== 0
    ) {
      return false
    }
    const result = evaluateDirectScalarCurrentValue(formula.directScalar)
    if (result?.kind !== 'number') {
      return false
    }
    const dependent = args.getSingleEntityDependent(makeCellEntity(request.cellIndex))
    if (
      dependent !== -1 &&
      !tryMarkDirectScalarLinearDeltaClosure(
        request.cellIndex,
        { tag: ValueTag.Number, value: request.oldNumber },
        { tag: ValueTag.Number, value: result.value },
        request.postRecalcDirectFormulaIndices,
      )
    ) {
      return false
    }
    if (!applyDirectFormulaCurrentResult(request.cellIndex, result)) {
      return false
    }
    countPostRecalcDirectFormulaMetric(request.cellIndex, request.postRecalcDirectFormulaMetrics)
    return true
  }

  const tryEvaluateDirectScalarWithPendingNumbers = (
    directScalar: RuntimeDirectScalarDescriptor,
    pendingNumbers: PendingNumericCellValues,
  ): DirectScalarCurrentOperand | undefined => {
    const readOperand = (operand: RuntimeDirectScalarOperand): DirectScalarCurrentOperand | undefined => {
      switch (operand.kind) {
        case 'literal-number':
          return { kind: 'number', value: operand.value }
        case 'error':
          return { kind: 'error', code: operand.code }
        case 'cell': {
          const pending = pendingNumbers.get(operand.cellIndex)
          if (pending !== undefined) {
            return { kind: 'number', value: pending }
          }
          const cellStore = args.state.workbook.cellStore
          const tag = (cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
          switch (tag) {
            case ValueTag.Number:
              return { kind: 'number', value: cellStore.numbers[operand.cellIndex] ?? 0 }
            case ValueTag.Boolean:
              return { kind: 'number', value: (cellStore.numbers[operand.cellIndex] ?? 0) !== 0 ? 1 : 0 }
            case ValueTag.Empty:
              return { kind: 'number', value: 0 }
            case ValueTag.Error:
              return { kind: 'error', code: (cellStore.errors[operand.cellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
            case ValueTag.String:
              return { kind: 'error', code: ErrorCode.Value }
          }
        }
      }
    }

    if (directScalar.kind === 'abs') {
      const operand = readOperand(directScalar.operand)
      return operand?.kind === 'number' ? { kind: 'number', value: Math.abs(operand.value) } : operand
    }
    const left = readOperand(directScalar.left)
    const right = readOperand(directScalar.right)
    if (!left || !right) {
      return undefined
    }
    if (left.kind === 'error') {
      return left
    }
    if (right.kind === 'error') {
      return right
    }
    switch (directScalar.operator) {
      case '+':
        return { kind: 'number', value: left.value + right.value }
      case '-':
        return { kind: 'number', value: left.value - right.value }
      case '*':
        return { kind: 'number', value: left.value * right.value }
      case '/':
        return right.value === 0 ? { kind: 'error', code: ErrorCode.Div0 } : { kind: 'number', value: left.value / right.value }
    }
  }

  const tryEvaluateDirectScalarNumericWithPendingNumbers = (
    directScalar: RuntimeDirectScalarDescriptor,
    pendingNumbers: PendingNumericCellValues,
  ): number | undefined => {
    const readOperand = (operand: RuntimeDirectScalarOperand): number | undefined => {
      switch (operand.kind) {
        case 'literal-number':
          return operand.value
        case 'error':
          return undefined
        case 'cell': {
          const pending = pendingNumbers.get(operand.cellIndex)
          if (pending !== undefined) {
            return pending
          }
          const cellStore = args.state.workbook.cellStore
          switch ((cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) {
            case ValueTag.Number:
              return cellStore.numbers[operand.cellIndex] ?? 0
            case ValueTag.Boolean:
              return (cellStore.numbers[operand.cellIndex] ?? 0) !== 0 ? 1 : 0
            case ValueTag.Empty:
              return 0
            case ValueTag.Error:
            case ValueTag.String:
              return undefined
          }
        }
      }
    }

    if (directScalar.kind === 'abs') {
      const operand = readOperand(directScalar.operand)
      return operand === undefined ? undefined : Math.abs(operand)
    }
    const left = readOperand(directScalar.left)
    const right = readOperand(directScalar.right)
    if (left === undefined || right === undefined) {
      return undefined
    }
    switch (directScalar.operator) {
      case '+':
        return left + right
      case '-':
        return left - right
      case '*':
        return left * right
      case '/':
        return right === 0 ? undefined : left / right
    }
  }

  const tryApplyDenseSingleColumnDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (firstRef === undefined || refs.length < 32) {
      return false
    }
    const firstMutation = firstRef.mutation
    if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
      return false
    }
    const secondMutation = refs[1]?.mutation
    if (secondMutation?.kind !== 'setCellValue') {
      return false
    }
    const rowOrder = secondMutation.row > firstMutation.row ? 1 : secondMutation.row < firstMutation.row ? -1 : 0
    if (rowOrder === 0) {
      return false
    }
    const firstSheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !firstSheet ||
      firstSheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col)
    ) {
      return false
    }
    const inputCellIndices = new Uint32Array(refs.length)
    const formulaCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = new Float64Array(refs.length)
    const formulaNumericResults = new Float64Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    const readDirectScalarOperandWithSingleReplacement = (
      operand: RuntimeDirectScalarOperand,
      inputCellIndex: number,
      inputValue: number,
    ): number | undefined => {
      switch (operand.kind) {
        case 'literal-number':
          return operand.value
        case 'error':
          return undefined
        case 'cell': {
          if (operand.cellIndex === inputCellIndex) {
            return inputValue
          }
          switch ((cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) {
            case ValueTag.Number:
              return cellStore.numbers[operand.cellIndex] ?? 0
            case ValueTag.Boolean:
              return (cellStore.numbers[operand.cellIndex] ?? 0) !== 0 ? 1 : 0
            case ValueTag.Empty:
              return 0
            case ValueTag.Error:
            case ValueTag.String:
              return undefined
          }
        }
      }
    }
    const evaluateDirectScalarWithSingleReplacement = (
      directScalar: RuntimeDirectScalarDescriptor,
      inputCellIndex: number,
      inputValue: number,
    ): number | undefined => {
      if (directScalar.kind === 'abs') {
        const operand = readDirectScalarOperandWithSingleReplacement(directScalar.operand, inputCellIndex, inputValue)
        return operand === undefined ? undefined : Math.abs(operand)
      }
      const left = readDirectScalarOperandWithSingleReplacement(directScalar.left, inputCellIndex, inputValue)
      const right = readDirectScalarOperandWithSingleReplacement(directScalar.right, inputCellIndex, inputValue)
      if (left === undefined || right === undefined) {
        return undefined
      }
      switch (directScalar.operator) {
        case '+':
          return left + right
        case '-':
          return left - right
        case '*':
          return left * right
        case '/':
          return right === 0 ? undefined : left / right
      }
    }
    let previousRow = firstMutation.row
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (refIndex > 0) {
        if ((rowOrder > 0 && mutation.row <= previousRow) || (rowOrder < 0 && mutation.row >= previousRow)) {
          return false
        }
        previousRow = mutation.row
      }
      if (
        ref.sheetId !== firstRef.sheetId ||
        mutation.kind !== 'setCellValue' ||
        mutation.col !== firstMutation.col ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0)
      ) {
        return false
      }
      const existingIndex =
        ref.cellIndex !== undefined &&
        args.state.workbook.cellStore.sheetIds[ref.cellIndex] === ref.sheetId &&
        args.state.workbook.cellStore.rows[ref.cellIndex] === mutation.row &&
        args.state.workbook.cellStore.cols[ref.cellIndex] === mutation.col
          ? ref.cellIndex
          : firstSheet.grid.getPhysical(mutation.row, mutation.col)
      if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex)) {
        return false
      }
      const singleDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
      if (singleDependent < 0 || !canUseDirectFormulaPostRecalc(singleDependent) || !canSkipDirectFormulaColumnVersion(singleDependent)) {
        return false
      }
      const formula = args.state.formulas.get(singleDependent)
      const result =
        formula?.directScalar === undefined
          ? undefined
          : evaluateDirectScalarWithSingleReplacement(formula.directScalar, existingIndex, mutation.value)
      if (result === undefined) {
        return false
      }
      const outputIndex = rowOrder < 0 ? refs.length - 1 - refIndex : refIndex
      inputCellIndices[outputIndex] = existingIndex
      formulaCellIndices[outputIndex] = singleDependent
      inputNumericValues[outputIndex] = mutation.value
      formulaNumericResults[outputIndex] = result
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    const reservedNewCells = potentialNewCells ?? 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length * 2 + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        writeNumericLiteralToCellStore(cellIndex, inputNumericValues[index]!)
        changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
        if (requiresChangedSet) {
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        }
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    const formulaChanged = requiresChangedSet ? new Uint32Array(refs.length) : EMPTY_CHANGED_CELLS
    for (let index = 0; index < refs.length; index += 1) {
      const formulaCellIndex = formulaCellIndices[index]!
      applyTerminalDirectFormulaNumericResult(formulaCellIndex, formulaNumericResults[index]!)
      if (requiresChangedSet) {
        formulaChanged[index] = formulaCellIndex
      }
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', refs.length)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaChanged, explicitChangedCount) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, firstRef.sheetId, explicitChangedCount)
    }
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      batchId: previousMetrics.batchId + 1,
      changedInputCount,
      compileMs: 0,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasGeneralEventListeners || hasWatchedCellListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        changedCells: hasGeneralEventListeners ? args.captureChangedCells(changed) : [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      }
      args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
    }
    if (hasTrackedEventListeners) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
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
    return true
  }

  const tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (firstRef === undefined || refs.length < 32 || (potentialNewCells ?? 0) !== 0) {
      return false
    }
    const firstMutation = firstRef.mutation
    if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
      return false
    }
    const secondMutation = refs[1]?.mutation
    if (secondMutation?.kind !== 'setCellValue') {
      return false
    }
    const rowOrder = secondMutation.row > firstMutation.row ? 1 : secondMutation.row < firstMutation.row ? -1 : 0
    if (rowOrder === 0) {
      return false
    }
    const firstSheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !firstSheet ||
      firstSheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col)
    ) {
      return false
    }

    const inputCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = new Float64Array(refs.length)
    const formulaCellIndices = new Uint32Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    let previousRow = firstMutation.row
    let previousFormulaRow = -1
    let previousFormulaCol = -1
    let affineScale: number | undefined
    let affineOffset: number | undefined
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (refIndex > 0) {
        if ((rowOrder > 0 && mutation.row <= previousRow) || (rowOrder < 0 && mutation.row >= previousRow)) {
          return false
        }
        previousRow = mutation.row
      }
      if (
        ref.sheetId !== firstRef.sheetId ||
        mutation.kind !== 'setCellValue' ||
        mutation.col !== firstMutation.col ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0) ||
        ref.cellIndex === undefined ||
        cellStore.sheetIds[ref.cellIndex] !== ref.sheetId ||
        cellStore.rows[ref.cellIndex] !== mutation.row ||
        cellStore.cols[ref.cellIndex] !== mutation.col
      ) {
        return false
      }
      const existingIndex = ref.cellIndex
      if (!canFastPathLiteralOverwrite(existingIndex)) {
        return false
      }
      const singleDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
      if (singleDependent < 0 || !canUseDirectFormulaPostRecalc(singleDependent) || !canSkipDirectFormulaColumnVersion(singleDependent)) {
        return false
      }
      const formula = args.state.formulas.get(singleDependent)
      if (
        !formula ||
        formula.directScalar === undefined ||
        cellStore.sheetIds[singleDependent] !== firstRef.sheetId ||
        cellStore.rows[singleDependent] !== mutation.row
      ) {
        return false
      }
      const formulaRow = cellStore.rows[singleDependent] ?? 0
      const formulaCol = cellStore.cols[singleDependent] ?? 0
      if (
        refIndex > 0 &&
        ((rowOrder > 0 && (formulaRow < previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol <= previousFormulaCol))) ||
          (rowOrder < 0 && (formulaRow > previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol >= previousFormulaCol))))
      ) {
        return false
      }
      const affine = singleInputAffineDirectScalar(formula.directScalar, existingIndex)
      if (affine === null) {
        return false
      }
      if (affineScale === undefined) {
        affineScale = affine.scale
        affineOffset = affine.offset
      } else if (!Object.is(affineScale, affine.scale) || !Object.is(affineOffset, affine.offset)) {
        return false
      }
      const outputIndex = rowOrder < 0 ? refs.length - 1 - refIndex : refIndex
      inputCellIndices[outputIndex] = existingIndex
      inputNumericValues[outputIndex] = mutation.value
      formulaCellIndices[outputIndex] = singleDependent
      previousFormulaRow = formulaRow
      previousFormulaCol = formulaCol
    }
    if (affineScale === undefined || affineOffset === undefined) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length * 2 + 1)
    args.resetMaterializedCellScratch(0)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        writeNumericLiteralToCellStore(cellIndex, inputNumericValues[index]!)
        changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
        if (requiresChangedSet) {
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        }
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    for (let index = 0; index < refs.length; index += 1) {
      applyTerminalDirectFormulaNumericResult(formulaCellIndices[index]!, inputNumericValues[index]! * affineScale + affineOffset)
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', refs.length)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaCellIndices, explicitChangedCount) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, firstRef.sheetId, explicitChangedCount)
    }
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      batchId: previousMetrics.batchId + 1,
      changedInputCount,
      compileMs: 0,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasGeneralEventListeners || hasWatchedCellListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        changedCells: hasGeneralEventListeners ? args.captureChangedCells(changed) : [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      }
      args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
    }
    if (hasTrackedEventListeners) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
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
    return true
  }

  const tryApplyDenseRowPairSimpleDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    if (refs.length < 32 || refs.length % 2 !== 0 || (potentialNewCells ?? 0) !== 0) {
      return false
    }
    const firstRef = refs[0]!
    const secondRef = refs[1]!
    const firstMutation = firstRef.mutation
    const secondMutation = secondRef.mutation
    if (
      firstMutation.kind !== 'setCellValue' ||
      secondMutation.kind !== 'setCellValue' ||
      firstRef.sheetId !== secondRef.sheetId ||
      firstMutation.row !== secondMutation.row ||
      firstMutation.col >= secondMutation.col ||
      typeof firstMutation.value !== 'number' ||
      typeof secondMutation.value !== 'number' ||
      Object.is(firstMutation.value, -0) ||
      Object.is(secondMutation.value, -0)
    ) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedExactLookupDependents(firstRef.sheetId, secondMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, secondMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, secondMutation.col)
    ) {
      return false
    }

    const formulaCellIndices = new Uint32Array(refs.length)
    const formulaCodes = new Uint8Array(refs.length)
    const mutationValues = new Float64Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    let formulaCount = 0
    let previousRow = firstMutation.row - 1
    let previousFormulaRow = -1
    let previousFormulaCol = -1
    for (let refIndex = 0; refIndex < refs.length; refIndex += 2) {
      const leftRef = refs[refIndex]!
      const rightRef = refs[refIndex + 1]!
      const leftMutation = leftRef.mutation
      const rightMutation = rightRef.mutation
      if (
        leftRef.sheetId !== firstRef.sheetId ||
        rightRef.sheetId !== firstRef.sheetId ||
        leftMutation.kind !== 'setCellValue' ||
        rightMutation.kind !== 'setCellValue' ||
        leftMutation.row !== rightMutation.row ||
        leftMutation.row <= previousRow ||
        leftMutation.col !== firstMutation.col ||
        rightMutation.col !== secondMutation.col ||
        typeof leftMutation.value !== 'number' ||
        typeof rightMutation.value !== 'number' ||
        Object.is(leftMutation.value, -0) ||
        Object.is(rightMutation.value, -0) ||
        leftRef.cellIndex === undefined ||
        rightRef.cellIndex === undefined ||
        cellStore.sheetIds[leftRef.cellIndex] !== leftRef.sheetId ||
        cellStore.rows[leftRef.cellIndex] !== leftMutation.row ||
        cellStore.cols[leftRef.cellIndex] !== leftMutation.col ||
        cellStore.sheetIds[rightRef.cellIndex] !== rightRef.sheetId ||
        cellStore.rows[rightRef.cellIndex] !== rightMutation.row ||
        cellStore.cols[rightRef.cellIndex] !== rightMutation.col
      ) {
        return false
      }
      previousRow = leftMutation.row
      const leftValue = leftMutation.value
      const rightValue = rightMutation.value
      mutationValues[refIndex] = leftValue
      mutationValues[refIndex + 1] = rightValue
      const leftIndex = leftRef.cellIndex
      const rightIndex = rightRef.cellIndex
      if (!canFastPathLiteralOverwrite(leftIndex) || !canFastPathLiteralOverwrite(rightIndex)) {
        return false
      }
      const rowFormulaStart = formulaCount
      const considerDependent = (formulaCellIndex: number): boolean => {
        for (let index = rowFormulaStart; index < formulaCount; index += 1) {
          if (formulaCellIndices[index] === formulaCellIndex) {
            return true
          }
        }
        const formula = args.state.formulas.get(formulaCellIndex)
        if (
          !formula ||
          formula.directScalar === undefined ||
          cellStore.sheetIds[formulaCellIndex] !== firstRef.sheetId ||
          cellStore.rows[formulaCellIndex] !== leftMutation.row ||
          !canUseDirectFormulaPostRecalc(formulaCellIndex) ||
          !canSkipDirectFormulaColumnVersion(formulaCellIndex)
        ) {
          return false
        }
        const formulaRow = cellStore.rows[formulaCellIndex] ?? 0
        const formulaCol = cellStore.cols[formulaCellIndex] ?? 0
        if (formulaRow < previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol <= previousFormulaCol)) {
          return false
        }
        const code = rowPairDirectScalarCode(formula.directScalar, leftIndex, rightIndex)
        if (
          code === 0 ||
          evaluateRowPairDirectScalarCode(code, leftValue, rightValue) === undefined ||
          formulaCount >= formulaCellIndices.length
        ) {
          return false
        }
        formulaCellIndices[formulaCount] = formulaCellIndex
        formulaCodes[formulaCount] = code
        formulaCount += 1
        previousFormulaRow = formulaRow
        previousFormulaCol = formulaCol
        return true
      }
      const leftDependents = args.getEntityDependents(makeCellEntity(leftIndex))
      const rightDependents = args.getEntityDependents(makeCellEntity(rightIndex))
      if (leftDependents.length === 0 || rightDependents.length === 0) {
        return false
      }
      for (let index = 0; index < leftDependents.length; index += 1) {
        if (!considerDependent(leftDependents[index]!)) {
          return false
        }
      }
      for (let index = 0; index < rightDependents.length; index += 1) {
        if (!considerDependent(rightDependents[index]!)) {
          return false
        }
      }
      if (formulaCount !== rowFormulaStart + 2) {
        return false
      }
    }
    if (formulaCount !== refs.length) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length * 2 + 1)
    args.resetMaterializedCellScratch(0)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const ref = refs[index]!
        const cellIndex = ref.cellIndex!
        writeNumericLiteralToCellStore(cellIndex, mutationValues[index]!)
        changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
        if (requiresChangedSet) {
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        }
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col, secondMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    for (let refIndex = 0; refIndex < refs.length; refIndex += 2) {
      const leftValue = mutationValues[refIndex]!
      const rightValue = mutationValues[refIndex + 1]!
      for (let formulaIndex = refIndex; formulaIndex < refIndex + 2; formulaIndex += 1) {
        const result = evaluateRowPairDirectScalarCode(formulaCodes[formulaIndex]!, leftValue, rightValue)
        if (result === undefined) {
          throw new Error('Failed to apply direct row-pair scalar result')
        }
        applyTerminalDirectFormulaNumericResult(formulaCellIndices[formulaIndex]!, result)
      }
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', formulaCount)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaCellIndices, explicitChangedCount) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, firstRef.sheetId, explicitChangedCount)
    }
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      batchId: previousMetrics.batchId + 1,
      changedInputCount,
      compileMs: 0,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasGeneralEventListeners || hasWatchedCellListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        changedCells: hasGeneralEventListeners ? args.captureChangedCells(changed) : [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      }
      args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
    }
    if (hasTrackedEventListeners) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
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
    return true
  }

  const tryApplyDenseRowPairDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    if (refs.length < 32 || refs.length % 2 !== 0) {
      return false
    }
    const firstRef = refs[0]!
    const secondRef = refs[1]!
    const firstMutation = firstRef.mutation
    const secondMutation = secondRef.mutation
    if (
      firstMutation.kind !== 'setCellValue' ||
      secondMutation.kind !== 'setCellValue' ||
      firstRef.sheetId !== secondRef.sheetId ||
      firstMutation.row !== secondMutation.row ||
      firstMutation.col >= secondMutation.col ||
      typeof firstMutation.value !== 'number' ||
      typeof secondMutation.value !== 'number' ||
      Object.is(firstMutation.value, -0) ||
      Object.is(secondMutation.value, -0)
    ) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedExactLookupDependents(firstRef.sheetId, secondMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, secondMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, secondMutation.col)
    ) {
      return false
    }
    const inputCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = new Float64Array(refs.length)
    const formulaCellIndices = new Uint32Array(refs.length * 2)
    const formulaNumericResults = new Float64Array(refs.length * 2)
    const cellStore = args.state.workbook.cellStore
    let formulaCount = 0
    let previousRow = firstMutation.row - 1
    let previousFormulaRow = -1
    let previousFormulaCol = -1

    const readDirectScalarOperandWithRowPair = (
      operand: RuntimeDirectScalarOperand,
      leftCellIndex: number,
      leftValue: number,
      rightCellIndex: number,
      rightValue: number,
    ): number | undefined => {
      switch (operand.kind) {
        case 'literal-number':
          return operand.value
        case 'error':
          return undefined
        case 'cell': {
          if (operand.cellIndex === leftCellIndex) {
            return leftValue
          }
          if (operand.cellIndex === rightCellIndex) {
            return rightValue
          }
          switch ((cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) {
            case ValueTag.Number:
              return cellStore.numbers[operand.cellIndex] ?? 0
            case ValueTag.Boolean:
              return (cellStore.numbers[operand.cellIndex] ?? 0) !== 0 ? 1 : 0
            case ValueTag.Empty:
              return 0
            case ValueTag.Error:
            case ValueTag.String:
              return undefined
          }
        }
      }
    }

    const evaluateDirectScalarWithRowPair = (
      directScalar: RuntimeDirectScalarDescriptor,
      leftCellIndex: number,
      leftValue: number,
      rightCellIndex: number,
      rightValue: number,
    ): number | undefined => {
      if (directScalar.kind === 'abs') {
        const operand = readDirectScalarOperandWithRowPair(directScalar.operand, leftCellIndex, leftValue, rightCellIndex, rightValue)
        return operand === undefined ? undefined : Math.abs(operand)
      }
      const left = readDirectScalarOperandWithRowPair(directScalar.left, leftCellIndex, leftValue, rightCellIndex, rightValue)
      const right = readDirectScalarOperandWithRowPair(directScalar.right, leftCellIndex, leftValue, rightCellIndex, rightValue)
      if (left === undefined || right === undefined) {
        return undefined
      }
      switch (directScalar.operator) {
        case '+':
          return left + right
        case '-':
          return left - right
        case '*':
          return left * right
        case '/':
          return right === 0 ? undefined : left / right
      }
    }

    for (let refIndex = 0; refIndex < refs.length; refIndex += 2) {
      const leftRef = refs[refIndex]!
      const rightRef = refs[refIndex + 1]!
      const leftMutation = leftRef.mutation
      const rightMutation = rightRef.mutation
      if (
        leftRef.sheetId !== firstRef.sheetId ||
        rightRef.sheetId !== firstRef.sheetId ||
        leftMutation.kind !== 'setCellValue' ||
        rightMutation.kind !== 'setCellValue' ||
        leftMutation.row !== rightMutation.row ||
        leftMutation.row <= previousRow ||
        leftMutation.col !== firstMutation.col ||
        rightMutation.col !== secondMutation.col ||
        typeof leftMutation.value !== 'number' ||
        typeof rightMutation.value !== 'number' ||
        Object.is(leftMutation.value, -0) ||
        Object.is(rightMutation.value, -0)
      ) {
        return false
      }
      const leftValue = leftMutation.value
      const rightValue = rightMutation.value
      previousRow = leftMutation.row
      const leftIndex =
        leftRef.cellIndex !== undefined &&
        cellStore.sheetIds[leftRef.cellIndex] === leftRef.sheetId &&
        cellStore.rows[leftRef.cellIndex] === leftMutation.row &&
        cellStore.cols[leftRef.cellIndex] === leftMutation.col
          ? leftRef.cellIndex
          : sheet.grid.getPhysical(leftMutation.row, leftMutation.col)
      const rightIndex =
        rightRef.cellIndex !== undefined &&
        cellStore.sheetIds[rightRef.cellIndex] === rightRef.sheetId &&
        cellStore.rows[rightRef.cellIndex] === rightMutation.row &&
        cellStore.cols[rightRef.cellIndex] === rightMutation.col
          ? rightRef.cellIndex
          : sheet.grid.getPhysical(rightMutation.row, rightMutation.col)
      if (leftIndex === -1 || rightIndex === -1 || !canFastPathLiteralOverwrite(leftIndex) || !canFastPathLiteralOverwrite(rightIndex)) {
        return false
      }
      const rowFormulaStart = formulaCount
      const considerDependent = (formulaCellIndex: number): boolean => {
        for (let index = rowFormulaStart; index < formulaCount; index += 1) {
          if (formulaCellIndices[index] === formulaCellIndex) {
            return true
          }
        }
        const formula = args.state.formulas.get(formulaCellIndex)
        if (
          !formula ||
          formula.directScalar === undefined ||
          cellStore.sheetIds[formulaCellIndex] !== firstRef.sheetId ||
          cellStore.rows[formulaCellIndex] !== leftMutation.row ||
          !canUseDirectFormulaPostRecalc(formulaCellIndex) ||
          !canSkipDirectFormulaColumnVersion(formulaCellIndex)
        ) {
          return false
        }
        const formulaRow = cellStore.rows[formulaCellIndex] ?? 0
        const formulaCol = cellStore.cols[formulaCellIndex] ?? 0
        if (formulaRow < previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol <= previousFormulaCol)) {
          return false
        }
        const result = evaluateDirectScalarWithRowPair(formula.directScalar, leftIndex, leftValue, rightIndex, rightValue)
        if (result === undefined || formulaCount >= formulaCellIndices.length) {
          return false
        }
        formulaCellIndices[formulaCount] = formulaCellIndex
        formulaNumericResults[formulaCount] = result
        formulaCount += 1
        previousFormulaRow = formulaRow
        previousFormulaCol = formulaCol
        return true
      }
      const leftDependents = args.getEntityDependents(makeCellEntity(leftIndex))
      const rightDependents = args.getEntityDependents(makeCellEntity(rightIndex))
      if (leftDependents.length === 0 || rightDependents.length === 0) {
        return false
      }
      for (let index = 0; index < leftDependents.length; index += 1) {
        if (!considerDependent(leftDependents[index]!)) {
          return false
        }
      }
      for (let index = 0; index < rightDependents.length; index += 1) {
        if (!considerDependent(rightDependents[index]!)) {
          return false
        }
      }
      inputCellIndices[refIndex] = leftIndex
      inputCellIndices[refIndex + 1] = rightIndex
      inputNumericValues[refIndex] = leftValue
      inputNumericValues[refIndex + 1] = rightValue
    }
    if (formulaCount === 0) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    const reservedNewCells = potentialNewCells ?? 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + formulaCount + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        writeNumericLiteralToCellStore(cellIndex, inputNumericValues[index]!)
        changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
        if (requiresChangedSet) {
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        }
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col, secondMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    const formulaChanged = requiresChangedSet ? new Uint32Array(formulaCount) : EMPTY_CHANGED_CELLS
    for (let index = 0; index < formulaCount; index += 1) {
      const formulaCellIndex = formulaCellIndices[index]!
      applyTerminalDirectFormulaNumericResult(formulaCellIndex, formulaNumericResults[index]!)
      if (requiresChangedSet) {
        formulaChanged[index] = formulaCellIndex
      }
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', formulaCount)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaChanged, explicitChangedCount) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, firstRef.sheetId, explicitChangedCount)
    }
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      batchId: previousMetrics.batchId + 1,
      changedInputCount,
      compileMs: 0,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasGeneralEventListeners || hasWatchedCellListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        changedCells: hasGeneralEventListeners ? args.captureChangedCells(changed) : [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      }
      args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
    }
    if (hasTrackedEventListeners) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
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
    return true
  }

  const tryApplyLookupOnlyNumericColumnLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (firstRef === undefined || refs.length < 32) {
      return false
    }
    const firstMutation = firstRef.mutation
    if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (!sheet || sheet.structureVersion !== 1 || hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col)) {
      return false
    }
    const hasExactLookupDependents = hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col)
    const hasSortedLookupDependents = hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col)
    if (!hasExactLookupDependents && !hasSortedLookupDependents) {
      return false
    }

    const inputCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = new Float64Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    let ascending = true
    let descending = true
    let previousRow = -1
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (
        ref.sheetId !== firstRef.sheetId ||
        mutation.kind !== 'setCellValue' ||
        mutation.col !== firstMutation.col ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0)
      ) {
        return false
      }
      if (refIndex > 0) {
        ascending &&= mutation.row > previousRow
        descending &&= mutation.row < previousRow
      }
      previousRow = mutation.row
      const existingIndex =
        ref.cellIndex !== undefined &&
        cellStore.sheetIds[ref.cellIndex] === ref.sheetId &&
        cellStore.rows[ref.cellIndex] === mutation.row &&
        cellStore.cols[ref.cellIndex] === mutation.col
          ? ref.cellIndex
          : sheet.grid.getPhysical(mutation.row, mutation.col)
      if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex)) {
        return false
      }
      const oldNumber = directScalarCellNumericValue(existingIndex)
      if (oldNumber === undefined) {
        return false
      }
      if (
        hasExactLookupDependents &&
        !planExactLookupNumericColumnWrite(firstRef.sheetId, firstMutation.col, mutation.row, oldNumber, mutation.value).handled
      ) {
        return false
      }
      if (
        hasSortedLookupDependents &&
        !planApproximateLookupNumericColumnWrite(firstRef.sheetId, sheet.name, firstMutation.col, mutation.row, oldNumber, mutation.value)
          .handled
      ) {
        return false
      }
      inputCellIndices[refIndex] = existingIndex
      inputNumericValues[refIndex] = mutation.value
    }
    if (!ascending && !descending) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    const reservedNewCells = potentialNewCells ?? 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        writeNumericLiteralToCellStore(cellIndex, inputNumericValues[index]!)
        changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
        if (requiresChangedSet) {
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        }
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (hasExactLookupDependents) {
      args.invalidateExactLookupColumn({ sheetName: sheet.name, col: firstMutation.col })
    }
    if (hasSortedLookupDependents) {
      args.invalidateSortedLookupColumn({ sheetName: sheet.name, col: firstMutation.col })
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    const changed = requiresChangedSet ? (ascending ? inputCellIndices : reverseUint32Array(inputCellIndices)) : EMPTY_CHANGED_CELLS
    addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      batchId: previousMetrics.batchId + 1,
      changedInputCount,
      compileMs: 0,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasGeneralEventListeners || hasWatchedCellListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        changedCells: hasGeneralEventListeners ? args.captureChangedCells(changed) : [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      }
      args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
    }
    if (hasTrackedEventListeners) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
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
    return true
  }

  const tryApplyCoalescedDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ): boolean => {
    if (
      source === 'restore' ||
      (source !== 'local' && batch !== null) ||
      refs.length < 32 ||
      args.state.formulas.size === 0 ||
      args.state.workbook.hasPivots()
    ) {
      return false
    }
    if (args.hasVolatileFormulas?.()) {
      return false
    }
    if (tryApplyLookupOnlyNumericColumnLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyDenseSingleColumnDirectScalarLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyDenseRowPairSimpleDirectScalarLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyDenseRowPairDirectScalarLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }

    const pendingNumbers = new PendingNumericCellValues()
    const inputCellIndices: number[] = []
    const inputNumericValues = new Float64Array(refs.length)
    const formulaCellIndices: number[] = []
    const formulaSeen = new Set<number>()
    let canUseNumericInputWrites = true
    let trustedPhysicalSheetId: number | undefined
    let canTrustPhysicalTrackedSlices = true
    let previousInputRow = -1
    let previousInputCol = -1
    let previousFormulaRow = -1
    let previousFormulaCol = -1
    const notePhysicalSliceCell = (sheetId: number, row: number, col: number, previous: 'input' | 'formula'): void => {
      if (!canTrustPhysicalTrackedSlices) {
        return
      }
      if (trustedPhysicalSheetId === undefined) {
        trustedPhysicalSheetId = sheetId
      } else if (trustedPhysicalSheetId !== sheetId) {
        canTrustPhysicalTrackedSlices = false
        return
      }
      if (previous === 'input') {
        if (row < previousInputRow || (row === previousInputRow && col < previousInputCol)) {
          canTrustPhysicalTrackedSlices = false
          return
        }
        previousInputRow = row
        previousInputCol = col
        return
      }
      if (row < previousFormulaRow || (row === previousFormulaRow && col < previousFormulaCol)) {
        canTrustPhysicalTrackedSlices = false
        return
      }
      previousFormulaRow = row
      previousFormulaCol = col
    }
    const trackedColumnDependencyFlagsBySheet = new Map<number, Map<number, boolean>>()
    const hasTrackedColumnDependencies = (sheetId: number, col: number): boolean => {
      let flagsByColumn = trackedColumnDependencyFlagsBySheet.get(sheetId)
      if (flagsByColumn === undefined) {
        flagsByColumn = new Map()
        trackedColumnDependencyFlagsBySheet.set(sheetId, flagsByColumn)
      }
      const cached = flagsByColumn.get(col)
      if (cached !== undefined) {
        return cached
      }
      const hasDependencies =
        hasTrackedExactLookupDependents(sheetId, col) ||
        hasTrackedSortedLookupDependents(sheetId, col) ||
        hasTrackedDirectRangeDependents(sheetId, col)
      flagsByColumn.set(col, hasDependencies)
      return hasDependencies
    }
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (mutation.kind !== 'setCellValue') {
        return false
      }
      if (typeof mutation.value !== 'number' || Object.is(mutation.value, -0)) {
        canUseNumericInputWrites = false
      }
      const nextNumber = directScalarLiteralNumericValue(mutation.value)
      if (nextNumber === undefined) {
        return false
      }
      inputNumericValues[refIndex] = nextNumber
      const sheet = args.state.workbook.getSheetById(ref.sheetId)
      if (!sheet || sheet.structureVersion !== 1) {
        return false
      }
      const candidate = ref.cellIndex
      const existingIndex =
        candidate !== undefined &&
        args.state.workbook.cellStore.sheetIds[candidate] === ref.sheetId &&
        args.state.workbook.cellStore.rows[candidate] === mutation.row &&
        args.state.workbook.cellStore.cols[candidate] === mutation.col
          ? candidate
          : sheet.grid.getPhysical(mutation.row, mutation.col)
      if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex) || pendingNumbers.has(existingIndex)) {
        return false
      }
      notePhysicalSliceCell(ref.sheetId, mutation.row, mutation.col, 'input')
      if (hasTrackedColumnDependencies(ref.sheetId, mutation.col)) {
        return false
      }
      const dependents = args.getEntityDependents(makeCellEntity(existingIndex))
      for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
        const formulaCellIndex = dependents[dependentIndex]!
        const formula = args.state.formulas.get(formulaCellIndex)
        if (
          !formula ||
          formula.directScalar === undefined ||
          !canUseDirectFormulaPostRecalc(formulaCellIndex) ||
          ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
        ) {
          return false
        }
        if (!formulaSeen.has(formulaCellIndex)) {
          formulaSeen.add(formulaCellIndex)
          formulaCellIndices.push(formulaCellIndex)
          const cellStore = args.state.workbook.cellStore
          const formulaSheetId = cellStore.sheetIds[formulaCellIndex]
          const formulaSheet = formulaSheetId === undefined ? undefined : args.state.workbook.getSheetById(formulaSheetId)
          if (formulaSheetId === undefined || (formulaSheet && formulaSheet.structureVersion !== 1)) {
            canTrustPhysicalTrackedSlices = false
          } else {
            notePhysicalSliceCell(formulaSheetId, cellStore.rows[formulaCellIndex] ?? 0, cellStore.cols[formulaCellIndex] ?? 0, 'formula')
          }
        }
      }
      pendingNumbers.set(existingIndex, nextNumber)
      inputCellIndices.push(existingIndex)
    }
    if (inputCellIndices.length === 0 || formulaCellIndices.length === 0) {
      return false
    }
    const formulaNumericResults = new Float64Array(formulaCellIndices.length)
    let canUseNumericFormulaResults = true
    for (let index = 0; index < formulaCellIndices.length; index += 1) {
      const formula = args.state.formulas.get(formulaCellIndices[index]!)
      const result = formula?.directScalar
        ? tryEvaluateDirectScalarNumericWithPendingNumbers(formula.directScalar, pendingNumbers)
        : undefined
      if (result === undefined) {
        canUseNumericFormulaResults = false
        break
      }
      formulaNumericResults[index] = result
    }
    let formulaResults: DirectScalarCurrentOperand[] | undefined
    if (!canUseNumericFormulaResults) {
      const evaluatedFormulaResults: DirectScalarCurrentOperand[] = []
      for (let index = 0; index < formulaCellIndices.length; index += 1) {
        const formula = args.state.formulas.get(formulaCellIndices[index]!)
        const result = formula?.directScalar ? tryEvaluateDirectScalarWithPendingNumbers(formula.directScalar, pendingNumbers) : undefined
        if (result === undefined) {
          return false
        }
        evaluatedFormulaResults[index] = result
      }
      formulaResults = evaluatedFormulaResults
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    const reservedNewCells = potentialNewCells ?? 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + formulaCellIndices.length + refs.length + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        for (let index = 0; index < refs.length; index += 1) {
          const cellIndex = inputCellIndices[index]!
          if (canUseNumericInputWrites) {
            writeNumericLiteralToCellStore(cellIndex, inputNumericValues[index]!)
          } else {
            const mutation = refs[index]!.mutation
            if (mutation.kind !== 'setCellValue') {
              throw new Error('Expected coalesced direct scalar batch to contain only literal writes')
            }
            writeLiteralToCellStore(args.state.workbook.cellStore, cellIndex, mutation.value, args.state.strings)
          }
          args.state.workbook.notifyCellValueWritten(cellIndex)
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          if (requiresChangedSet) {
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
          }
        }
      })
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    const formulaChanged = requiresChangedSet ? new Uint32Array(formulaCellIndices.length) : EMPTY_CHANGED_CELLS
    args.state.workbook.withBatchedColumnVersionUpdates(() => {
      for (let index = 0; index < formulaCellIndices.length; index += 1) {
        const cellIndex = formulaCellIndices[index]!
        if (canUseNumericFormulaResults) {
          applyDirectFormulaNumericResult(cellIndex, formulaNumericResults[index]!)
        } else if (!applyDirectFormulaCurrentResult(cellIndex, formulaResults![index]!)) {
          throw new Error('Failed to apply direct scalar batch result')
        }
        if (requiresChangedSet) {
          formulaChanged[index] = cellIndex
        }
      }
    })
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', formulaCellIndices.length)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')

    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaChanged, explicitChangedCount) : EMPTY_CHANGED_CELLS
    if (
      hasTrackedEventListeners &&
      requiresChangedSet &&
      canTrustPhysicalTrackedSlices &&
      trustedPhysicalSheetId !== undefined &&
      explicitChangedCount > 0 &&
      explicitChangedCount < changed.length
    ) {
      tagTrustedPhysicalTrackedChanges(changed, trustedPhysicalSheetId, explicitChangedCount)
    }
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      batchId: previousMetrics.batchId + 1,
      changedInputCount,
      compileMs: 0,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasGeneralEventListeners || hasWatchedCellListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        changedCells: hasGeneralEventListeners ? args.captureChangedCells(changed) : [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      }
      args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
    }
    if (hasTrackedEventListeners) {
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
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
    return true
  }

  const tryApplySingleExistingDirectLiteralMutation = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
  ): boolean => {
    if (
      source !== 'local' ||
      batch !== null ||
      refs.length !== 1 ||
      args.state.workbook.hasPivots() ||
      args.state.events.hasListeners() ||
      args.state.events.hasCellListeners()
    ) {
      return false
    }
    if (args.hasVolatileFormulas?.()) {
      return false
    }
    const ref = refs[0]!
    const mutation = ref.mutation
    if (mutation.kind !== 'setCellValue' || mutation.value === null) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(ref.sheetId)
    if (!sheet || sheet.structureVersion !== 1) {
      return false
    }
    const existingIndex =
      ref.cellIndex !== undefined &&
      args.state.workbook.cellStore.sheetIds[ref.cellIndex] === ref.sheetId &&
      args.state.workbook.cellStore.rows[ref.cellIndex] === mutation.row &&
      args.state.workbook.cellStore.cols[ref.cellIndex] === mutation.col
        ? ref.cellIndex
        : sheet.grid.getPhysical(mutation.row, mutation.col)
    const sheetName = sheet.name
    const hasExactLookupDependents = hasTrackedExactLookupDependents(ref.sheetId, mutation.col)
    const hasSortedLookupDependents = hasTrackedSortedLookupDependents(ref.sheetId, mutation.col)
    const hasAggregateDependents = hasTrackedDirectRangeDependents(ref.sheetId, mutation.col)
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    if (existingIndex === -1) {
      if (
        args.state.trackReplicaVersions ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0) ||
        hasExactLookupDependents ||
        hasAggregateDependents ||
        (hasSortedLookupDependents && !canSkipApproximateLookupNewNumericColumnWrite(ref.sheetId, mutation.col, mutation.row))
      ) {
        return false
      }
      const cellIndex = args.state.workbook.ensureCellAt(ref.sheetId, mutation.row, mutation.col).cellIndex
      writeNumericLiteralToExistingCell(cellIndex, mutation.value)
      deferSingleCellKernelSync(cellIndex)
      const lastMetrics = makeSingleLiteralSkipMetrics()
      args.state.setLastMetrics(lastMetrics)
      addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
      if (hasTrackedEventListeners) {
        args.state.events.emitTracked({
          kind: 'batch',
          invalidation: 'cells',
          changedCellIndices: Uint32Array.of(cellIndex),
          invalidatedRanges: [],
          invalidatedRows: [],
          invalidatedColumns: [],
          metrics: lastMetrics,
          explicitChangedCount: 1,
        })
      }
      return true
    }
    if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex)) {
      return false
    }
    const oldNumber = directScalarCellNumericValue(existingIndex)
    const newNumber = directScalarLiteralNumericValue(mutation.value)
    if (oldNumber === undefined || newNumber === undefined) {
      return false
    }

    const singleExistingCellDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
    if (
      hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      (singleExistingCellDependent === -1 || isRangeEntity(singleExistingCellDependent)) &&
      tryApplySingleDirectAggregateLiteralMutationFastPath({
        existingIndex,
        sheetId: ref.sheetId,
        sheetName,
        row: mutation.row,
        col: mutation.col,
        value: mutation.value,
        delta: newNumber - oldNumber,
        emitTracked: hasTrackedEventListeners,
        ...(isRangeEntity(singleExistingCellDependent) ? { singleRangeEntityDependent: singleExistingCellDependent } : {}),
      })
    ) {
      return true
    }
    const existingTag = (args.state.workbook.cellStore.tags[existingIndex] as ValueTag | undefined) ?? ValueTag.Empty
    const mutationIsNumber = typeof mutation.value === 'number'
    const directLookupExactMutationNumber = mutationIsNumber ? newNumber : undefined
    const directLookupApproximateMutationNumber = newNumber
    const oldExactLookupNumber = hasExactLookupDependents && existingTag === ValueTag.Number ? oldNumber : undefined
    const newExactLookupNumber = hasExactLookupDependents && mutationIsNumber ? newNumber : undefined
    const oldApproximateLookupNumber = hasSortedLookupDependents ? oldNumber : undefined
    const newApproximateLookupNumber = hasSortedLookupDependents ? newNumber : undefined
    const exactLookupWritePlan =
      hasExactLookupDependents && oldExactLookupNumber !== undefined && newExactLookupNumber !== undefined
        ? planExactLookupNumericColumnWrite(ref.sheetId, mutation.col, mutation.row, oldExactLookupNumber, newExactLookupNumber)
        : { handled: false }
    const sortedLookupWritePlan =
      hasSortedLookupDependents && oldApproximateLookupNumber !== undefined && newApproximateLookupNumber !== undefined
        ? planApproximateLookupNumericColumnWrite(
            ref.sheetId,
            sheetName,
            mutation.col,
            mutation.row,
            oldApproximateLookupNumber,
            newApproximateLookupNumber,
          )
        : { handled: false }
    const exactLookupDependentsHandled = hasExactLookupDependents && exactLookupWritePlan.handled
    const sortedLookupDependentsHandled = hasSortedLookupDependents && sortedLookupWritePlan.handled
    if ((hasExactLookupDependents && !exactLookupDependentsHandled) || (hasSortedLookupDependents && !sortedLookupDependentsHandled)) {
      return false
    }

    const lookupDependentsHandled =
      (hasExactLookupDependents && exactLookupDependentsHandled) || (hasSortedLookupDependents && sortedLookupDependentsHandled)
    const canUseNumericLookupWriteFastPath = lookupDependentsHandled && existingTag === ValueTag.Number && mutationIsNumber
    if (!hasAggregateDependents && (hasExactLookupDependents || hasSortedLookupDependents) && singleExistingCellDependent === -1) {
      if (canUseNumericLookupWriteFastPath) {
        writeNumericLiteralToExistingCell(existingIndex, newNumber)
        const currentColumnVersion = sheet.columnVersions[mutation.col] ?? 0
        if (exactLookupWritePlan.tailPatchTarget !== undefined) {
          exactLookupWritePlan.tailPatchTarget.tailPatch = {
            row: mutation.row,
            oldNumeric: oldNumber,
            newNumeric: newNumber,
            columnVersion: currentColumnVersion,
          }
        }
        if (sortedLookupWritePlan.tailPatchTarget !== undefined) {
          sortedLookupWritePlan.tailPatchTarget.tailPatch = {
            row: mutation.row,
            oldNumeric: oldNumber,
            newNumeric: newNumber,
            columnVersion: currentColumnVersion,
          }
        }
        const needsExactPatch =
          hasExactLookupDependents && exactLookupDependentsHandled && exactLookupWritePlan.tailPatchTarget === undefined
        const needsSortedPatch =
          hasSortedLookupDependents && sortedLookupDependentsHandled && sortedLookupWritePlan.tailPatchTarget === undefined
        const patchedLookupOwners =
          needsExactPatch || needsSortedPatch
            ? patchUniformLookupTailWrites({
                sheetId: ref.sheetId,
                col: mutation.col,
                row: mutation.row,
                oldNumeric: oldNumber,
                newNumeric: newNumber,
                exact: needsExactPatch,
                sorted: needsSortedPatch,
              })
            : { exact: true, sorted: true }
        if (needsExactPatch && !patchedLookupOwners.exact) {
          args.invalidateExactLookupColumn({ sheetName, col: mutation.col })
        }
        if (needsSortedPatch && !patchedLookupOwners.sorted) {
          args.invalidateSortedLookupColumn({ sheetName, col: mutation.col })
        }
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
        deferSingleCellKernelSync(existingIndex)
        const lastMetrics = makeSingleLiteralSkipMetrics()
        args.state.setLastMetrics(lastMetrics)
        if (hasTrackedEventListeners) {
          args.state.events.emitTracked({
            kind: 'batch',
            invalidation: 'cells',
            changedCellIndices: Uint32Array.of(existingIndex),
            invalidatedRanges: [],
            invalidatedRows: [],
            invalidatedColumns: [],
            metrics: lastMetrics,
            explicitChangedCount: 1,
          })
        }
        return true
      }
      if (
        tryApplySingleKernelSyncOnlyLiteralMutationFastPath({
          existingIndex,
          value: mutation.value,
          emitTracked: hasTrackedEventListeners,
        })
      ) {
        return true
      }
    }

    if (!hasTrackedEventListeners && !hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents) {
      if (
        tryApplySingleDirectFormulaLiteralMutationWithoutEvents({
          existingIndex,
          formulaCellIndex: singleExistingCellDependent,
          value: mutation.value,
          oldNumber,
          newNumber,
          exactLookupValue: directLookupExactMutationNumber,
          approximateLookupValue: directLookupApproximateMutationNumber,
        }) ||
        tryApplySingleDirectScalarLiteralMutationWithoutEvents({
          existingIndex,
          value: mutation.value,
          oldNumber,
          newNumber,
        })
      ) {
        return true
      }
    }
    if (!hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents) {
      if (
        tryApplySingleDirectLookupOperandMutationFastPath({
          existingIndex,
          formulaCellIndex: singleExistingCellDependent,
          value: mutation.value,
          exactLookupValue: directLookupExactMutationNumber,
          approximateLookupValue: directLookupApproximateMutationNumber,
          emitTracked: hasTrackedEventListeners,
          lookupSheetHint: sheet,
        })
      ) {
        return true
      }
    }
    const oldValue: CellValue = { tag: ValueTag.Number, value: oldNumber }
    const newValue: CellValue = { tag: ValueTag.Number, value: newNumber }
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()
    let directDependentsHandled = markPostRecalcDirectScalarNumericDependents(
      existingIndex,
      oldNumber,
      newNumber,
      postRecalcDirectFormulaIndices,
      directLookupExactMutationNumber,
      directLookupApproximateMutationNumber,
    )
    if (
      !directDependentsHandled &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      !hasAggregateDependents &&
      tryMarkDirectScalarLinearDeltaClosure(existingIndex, oldValue, newValue, postRecalcDirectFormulaIndices)
    ) {
      directDependentsHandled = true
    }
    if (
      hasAggregateDependents &&
      directDependentsHandled &&
      postRecalcDirectFormulaIndices.size === 0 &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      tryApplySingleDirectAggregateLiteralMutationFastPath({
        existingIndex,
        sheetId: ref.sheetId,
        sheetName,
        row: mutation.row,
        col: mutation.col,
        value: mutation.value,
        delta: newNumber - oldNumber,
        emitTracked: hasTrackedEventListeners,
        ...(isRangeEntity(singleExistingCellDependent) ? { singleRangeEntityDependent: singleExistingCellDependent } : {}),
      })
    ) {
      return true
    }
    let shouldNoteAggregateLiteralWrite = false
    if (hasAggregateDependents) {
      if (!directDependentsHandled) {
        directDependentsHandled = true
      }
      const singleAffected = collectSingleAffectedDirectRangeDependent({
        sheetName,
        sheetId: ref.sheetId,
        row: mutation.row,
        col: mutation.col,
      })
      if (singleAffected >= 0) {
        const formula = args.state.formulas.get(singleAffected)
        if (
          !formula ||
          formula.directAggregate?.aggregateKind !== 'sum' ||
          formula.dependencyIndices.length !== 0 ||
          args.getSingleEntityDependent(makeCellEntity(singleAffected)) !== -1
        ) {
          return false
        }
        postRecalcDirectFormulaIndices.addDelta(singleAffected, newNumber - oldNumber)
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(existingIndex)
        shouldNoteAggregateLiteralWrite = true
      } else if (singleAffected === -2) {
        const affected = collectAffectedDirectRangeDependents({
          sheetName,
          row: mutation.row,
          col: mutation.col,
        })
        if (affected.length === 0 || affected.length > DIRECT_RANGE_POST_RECALC_LIMIT) {
          return false
        }
        for (let index = 0; index < affected.length; index += 1) {
          const formulaCellIndex = affected[index]!
          const formula = args.state.formulas.get(formulaCellIndex)
          if (
            !formula ||
            formula.directAggregate?.aggregateKind !== 'sum' ||
            formula.dependencyIndices.length !== 0 ||
            args.getSingleEntityDependent(makeCellEntity(formulaCellIndex)) !== -1
          ) {
            return false
          }
        }
        postRecalcDirectFormulaIndices.appendConstantDelta(affected, newNumber - oldNumber)
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(existingIndex)
        shouldNoteAggregateLiteralWrite = true
      }
    }
    if (
      !directDependentsHandled ||
      (postRecalcDirectFormulaIndices.size > 1 && !hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices))
    ) {
      return false
    }

    const explicitChangedCount = hasTrackedEventListeners ? 1 : 0
    const postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts = {
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    }
    let recalculated: U32 = EMPTY_CHANGED_CELLS
    args.state.workbook.withBatchedColumnVersionUpdates(() => {
      writeLiteralToCellStore(args.state.workbook.cellStore, existingIndex, mutation.value, args.state.strings)
      args.state.workbook.notifyCellValueWritten(existingIndex)
      if (shouldNoteAggregateLiteralWrite) {
        args.noteAggregateLiteralWrite({
          sheetName,
          row: mutation.row,
          col: mutation.col,
          oldValue,
          newValue,
        })
      }
      if (postRecalcDirectFormulaIndices.size > 0) {
        if (hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices)) {
          countDirectFormulaDeltaSkip(args.state.formulas, postRecalcDirectFormulaIndices, args.state.counters)
        } else if (canEvaluatePostRecalcDirectFormulasWithoutKernel(args.state.formulas, postRecalcDirectFormulaIndices)) {
          addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
        }
        const directChanged =
          tryApplySinglePostRecalcDirectFormula(
            postRecalcDirectFormulaIndices,
            false,
            postRecalcDirectFormulaMetrics,
            hasTrackedEventListeners,
          ) ??
          tryApplyDirectScalarDeltas(postRecalcDirectFormulaIndices, hasTrackedEventListeners) ??
          tryApplyDirectFormulaDeltas(postRecalcDirectFormulaIndices, hasTrackedEventListeners)
        if (directChanged === undefined) {
          throw new Error('Failed to apply single direct literal mutation fast path')
        }
        recalculated = directChanged
      } else if (hasExactLookupDependents || hasSortedLookupDependents) {
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
      }
    })

    deferSingleCellKernelSync(existingIndex)
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      dirtyFormulaCount: 0,
      wasmFormulaCount: postRecalcDirectFormulaMetrics.wasmFormulaCount,
      jsFormulaCount: postRecalcDirectFormulaMetrics.jsFormulaCount,
      rangeNodeVisits: 0,
      recalcMs: 0,
      batchId: previousMetrics.batchId + 1,
      changedInputCount: 1,
      compileMs: 0,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasTrackedEventListeners) {
      const changed = composeSingleDisjointExplicitEventChanges(existingIndex, recalculated)
      if (changed.length > 4 && canTrustPhysicalTrackedChangeSplit(changed, ref.sheetId, explicitChangedCount)) {
        tagTrustedPhysicalTrackedChanges(changed, ref.sheetId, explicitChangedCount)
      }
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      })
    }
    return true
  }

  const applyExistingNumericCellMutationAtNow = (
    request: EngineExistingNumericCellMutationRef,
  ): EngineExistingNumericCellMutationResult | null => {
    if (
      args.state.workbook.hasPivots() ||
      args.state.events.hasListeners() ||
      args.state.events.hasCellListeners() ||
      args.hasVolatileFormulas?.()
    ) {
      return null
    }
    const sheet = args.state.workbook.getSheetById(request.sheetId)
    const cellStore = args.state.workbook.cellStore
    const existingIndex = request.cellIndex
    const trustedExistingNumericLiteral = request.trustedExistingNumericLiteral === true
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      (!trustedExistingNumericLiteral &&
        (cellStore.sheetIds[existingIndex] !== request.sheetId ||
          cellStore.rows[existingIndex] !== request.row ||
          cellStore.cols[existingIndex] !== request.col ||
          !canFastPathLiteralOverwrite(existingIndex)))
    ) {
      return null
    }
    const oldNumber = trustedExistingNumericLiteral
      ? request.oldNumericValue === undefined || Object.is(request.oldNumericValue, -0)
        ? 0
        : request.oldNumericValue
      : directScalarCellNumericValue(existingIndex)
    if (oldNumber === undefined || Object.is(request.value, -0)) {
      return null
    }
    const sheetName = sheet.name
    const singleExistingCellDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
    let hasExactLookupDependents: boolean | undefined
    let hasSortedLookupDependents: boolean | undefined
    if (trustedExistingNumericLiteral && request.emitTracked === false && isRangeEntity(singleExistingCellDependent)) {
      hasExactLookupDependents = hasTrackedExactLookupDependents(request.sheetId, request.col)
      hasSortedLookupDependents = hasTrackedSortedLookupDependents(request.sheetId, request.col)
      const trustedAggregateResult = tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation({
        existingIndex,
        rangeEntityDependent: singleExistingCellDependent,
        sheet,
        sheetId: request.sheetId,
        col: request.col,
        value: request.value,
        delta: request.value - oldNumber,
        hasExactLookupDependents,
        hasSortedLookupDependents,
      })
      if (trustedAggregateResult) {
        return trustedAggregateResult
      }
    }
    hasExactLookupDependents ??= hasTrackedExactLookupDependents(request.sheetId, request.col)
    hasSortedLookupDependents ??= hasTrackedSortedLookupDependents(request.sheetId, request.col)
    const hasTrackedEventListeners = request.emitTracked !== false && args.state.events.hasTrackedListeners()
    const hasAggregateDependents =
      isRangeEntity(singleExistingCellDependent) || hasTrackedDirectRangeDependents(request.sheetId, request.col)
    if (
      trustedExistingNumericLiteral &&
      !hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      singleExistingCellDependent >= 0
    ) {
      const scalarClosureResult = tryApplyTrustedDirectScalarClosureExistingNumericMutation({
        existingIndex,
        sheet,
        sheetId: request.sheetId,
        col: request.col,
        value: request.value,
        oldNumber,
        hasTrackedEventListeners,
      })
      if (scalarClosureResult) {
        return scalarClosureResult
      }
    }
    if (!hasAggregateDependents && (hasExactLookupDependents || hasSortedLookupDependents) && singleExistingCellDependent === -1) {
      const exactLookupWritePlan = hasExactLookupDependents
        ? planExactLookupNumericColumnWrite(request.sheetId, request.col, request.row, oldNumber, request.value)
        : { handled: true }
      const sortedLookupWritePlan = hasSortedLookupDependents
        ? planApproximateLookupNumericColumnWrite(request.sheetId, sheetName, request.col, request.row, oldNumber, request.value)
        : { handled: true }
      if (
        exactLookupWritePlan.handled &&
        sortedLookupWritePlan.handled &&
        (exactLookupWritePlan.tailPatchTarget === undefined || exactLookupWritePlan.tailPatchTarget.tailPatch === undefined) &&
        (sortedLookupWritePlan.tailPatchTarget === undefined || sortedLookupWritePlan.tailPatchTarget.tailPatch === undefined)
      ) {
        writeNumericLiteralToExistingCell(existingIndex, request.value)
        const currentColumnVersion = sheet.columnVersions[request.col] ?? 0
        if (exactLookupWritePlan.tailPatchTarget !== undefined) {
          exactLookupWritePlan.tailPatchTarget.tailPatch = {
            row: request.row,
            oldNumeric: oldNumber,
            newNumeric: request.value,
            columnVersion: currentColumnVersion,
          }
        }
        if (sortedLookupWritePlan.tailPatchTarget !== undefined) {
          sortedLookupWritePlan.tailPatchTarget.tailPatch = {
            row: request.row,
            oldNumeric: oldNumber,
            newNumeric: request.value,
            columnVersion: currentColumnVersion,
          }
        }
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
        deferSingleCellKernelSync(existingIndex)
        const lastMetrics = makeSingleLiteralSkipMetrics()
        args.state.setLastMetrics(lastMetrics)
        const changedCellIndices = Uint32Array.of(existingIndex)
        if (hasTrackedEventListeners) {
          args.state.events.emitTracked({
            kind: 'batch',
            invalidation: 'cells',
            changedCellIndices,
            invalidatedRanges: [],
            invalidatedRows: [],
            invalidatedColumns: [],
            metrics: lastMetrics,
            explicitChangedCount: 1,
          })
        }
        return makeExistingNumericMutationResult(changedCellIndices, 1)
      }
    }
    const aggregateFastPathResult =
      hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      (singleExistingCellDependent === -1 || isRangeEntity(singleExistingCellDependent))
        ? tryApplySingleDirectAggregateLiteralMutationFastPath({
            existingIndex,
            sheetId: request.sheetId,
            sheetName,
            row: request.row,
            col: request.col,
            value: request.value,
            delta: request.value - oldNumber,
            emitTracked: hasTrackedEventListeners,
            ...(isRangeEntity(singleExistingCellDependent) ? { singleRangeEntityDependent: singleExistingCellDependent } : {}),
          })
        : null
    if (aggregateFastPathResult) {
      return aggregateFastPathResult
    }
    const directLookupFastPathResult =
      !hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents
        ? tryApplySingleDirectLookupOperandMutationFastPath({
            existingIndex,
            formulaCellIndex: singleExistingCellDependent,
            value: request.value,
            exactLookupValue: request.value,
            approximateLookupValue: request.value,
            emitTracked: hasTrackedEventListeners,
            lookupSheetHint: sheet,
            ...(trustedExistingNumericLiteral ? { trustedInputSheet: sheet, trustedInputCol: request.col } : {}),
          })
        : null
    if (directLookupFastPathResult) {
      return directLookupFastPathResult
    }
    return null
  }

  const countPostRecalcDirectFormulaMetric = (cellIndex: number, counts: DirectFormulaMetricCounts): void => {
    const formula = args.state.formulas.get(cellIndex)
    if (!formula || (formula.directScalar === undefined && formula.directAggregate === undefined)) {
      return
    }
    if (formula.compiled.mode === FormulaMode.WasmFastPath) {
      counts.wasmFormulaCount += 1
      return
    }
    counts.jsFormulaCount += 1
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
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()
    const precomputedKernelSyncCellIndices: number[] = []
    let refreshAllPivots = false
    let appliedOps = 0
    const postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts = {
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    }
    const lookupHandledInputCellIndices: number[] = []
    const canSkipOrderChecks = source !== 'remote'
    let hadCycleMembersBefore: boolean | undefined
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
    const exactLookupImpactCaches: ExactLookupImpactCaches = new Map()
    const clearLookupImpactCaches = (): void => {
      exactLookupImpactCaches.clear()
    }

    const reservedNewCells = potentialNewCells ?? args.estimatePotentialNewCells(batch.ops)
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + reservedNewCells + 1)
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
        args.materializeDeferredStructuralFormulaSources()

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
            structural.transaction.removedCellIndices.forEach((cellIndex) => {
              precomputedKernelSyncCellIndices.push(cellIndex)
            })
            structural.precomputedChangedInputCellIndices.forEach((cellIndex) => {
              precomputedKernelSyncCellIndices.push(cellIndex)
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            })
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
            if (!isRestore && cellTouchesPivotSource(op.sheetName, parsedAddress.row, parsedAddress.col)) {
              refreshAllPivots = true
            }
            const needsLookupValueRead = hasExactLookupDependents || hasSortedLookupDependents || hasAggregateDependents
            const prior = readCellValueForLookup(existingIndex)
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
                clearLookupImpactCaches()
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
              const formulaChangedCountBeforeLookupNotes = formulaChangedCount
              const newValue = literalToValue(op.value, args.state.strings)
              const newStringId = typeof op.value === 'string' ? args.state.workbook.cellStore.stringIds[cellIndex] : undefined
              if (!isRestore) {
                const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                  cellIndex,
                  postRecalcDirectFormulaIndices,
                  prior.value,
                  newValue,
                )
                if (!directDependentsHandled) {
                  markDirectScalarDeltaClosure(cellIndex, prior.value, newValue, postRecalcDirectFormulaIndices)
                }
              }
              if (hasExactLookupDependents || hasAggregateDependents) {
                const exactLookupRequest = withOptionalLookupStringIds({
                  sheetName: op.sheetName,
                  row: parsedAddress.row,
                  col: parsedAddress.col,
                  oldValue: prior.value,
                  newValue,
                  oldStringId: prior.stringId,
                  newStringId,
                  inputCellIndex: cellIndex,
                })
                if (hasExactLookupDependents) {
                  formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                    exactLookupRequest,
                    formulaChangedCount,
                    exactLookupImpactCaches,
                  )
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
                formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
              }
              if (
                !hasAggregateDependents &&
                (hasExactLookupDependents || hasSortedLookupDependents) &&
                formulaChangedCount === formulaChangedCountBeforeLookupNotes
              ) {
                lookupHandledInputCellIndices.push(cellIndex)
              }
            } else if (!isRestore) {
              const newValue = literalToValue(op.value, args.state.strings)
              const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                cellIndex,
                postRecalcDirectFormulaIndices,
                prior.value,
                newValue,
              )
              if (!directDependentsHandled) {
                markDirectScalarDeltaClosure(cellIndex, prior.value, newValue, postRecalcDirectFormulaIndices)
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
            const sheetId = args.state.workbook.getSheet(op.sheetName)?.id
            if (!isRestore && cellTouchesPivotSource(op.sheetName, parsedAddress.row, parsedAddress.col)) {
              refreshAllPivots = true
            }
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
            const oldFormulaNumber = !isRestore && priorHadFormula ? readExactNumericValueForLookup(cellIndex) : undefined
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.AuthoredBlank
            if (!isRestore) {
              changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
            }
            const compileStarted = isRestore ? 0 : performance.now()
            try {
              const changedTopology = args.bindFormula(cellIndex, op.sheetName, op.formula)
              args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
              clearLookupImpactCaches()
              if (!isRestore) {
                compileMs += performance.now() - compileStarted
              }
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
              const handledFormulaReplacementAsDirectDelta =
                priorHadFormula &&
                sheetId !== undefined &&
                !hasTrackedExactLookupDependents(sheetId, parsedAddress.col) &&
                !hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) &&
                !hasTrackedDirectRangeDependents(sheetId, parsedAddress.col) &&
                tryApplyFormulaReplacementAsDirectScalarDeltaRoot({
                  cellIndex,
                  oldNumber: oldFormulaNumber,
                  changedTopology,
                  postRecalcDirectFormulaIndices,
                  postRecalcDirectFormulaMetrics,
                })
              if (!handledFormulaReplacementAsDirectDelta) {
                formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
              }
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
            if (!isRestore && cellTouchesPivotSource(op.sheetName, parsedAddress.row, parsedAddress.col)) {
              refreshAllPivots = true
            }
            const prior = readCellValueForLookup(cellIndex)
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
              clearLookupImpactCaches()
            }
            if (removedFormula) {
              formulaChangedCount = refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
            }
            topologyChanged = removedFormula || topologyChanged
            args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
            args.state.workbook.notifyCellValueWritten(cellIndex)
            if (!isRestore) {
              const nextValue = emptyValue()
              const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                cellIndex,
                postRecalcDirectFormulaIndices,
                prior.value,
                nextValue,
              )
              if (!directDependentsHandled) {
                markDirectScalarDeltaClosure(cellIndex, prior.value, nextValue, postRecalcDirectFormulaIndices)
              }
            }
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
                  inputCellIndex: cellIndex,
                })
                if (hasExactLookupDependents) {
                  formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                    exactLookupRequest,
                    formulaChangedCount,
                    exactLookupImpactCaches,
                  )
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
                formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
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
        !hadCycleMembersBeforeNow() &&
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
    const hasActiveFormulas = args.state.formulas.size > 0
    const hasActivePivots = args.state.workbook.hasPivots()
    const hasRecalcWork =
      changedInputCount > 0 ||
      formulaChangedCount > 0 ||
      precomputedKernelSyncCellIndices.length > 0 ||
      postRecalcDirectFormulaIndices.size > 0
    const hasVolatileFormulaWork = hasActiveFormulas && (args.hasVolatileFormulas ? args.hasVolatileFormulas() : true)
    const shouldRefreshPivots = refreshAllPivots && hasActivePivots
    let recalculated: U32 = new Uint32Array()
    let didRunRecalc = false
    let didFastDeferKernelSyncOnly = false
    let canComposeDisjointEventChanges = false
    if (
      hasActiveFormulas &&
      changedInputCount > 0 &&
      formulaChangedCount === 0 &&
      precomputedKernelSyncCellIndices.length === 0 &&
      postRecalcDirectFormulaIndices.size === 0 &&
      !refreshAllPivots &&
      !hasActivePivots &&
      !hasVolatileFormulaWork
    ) {
      const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
      if (
        canSkipDirtyTraversalForChangedInputs(changedInputArray, changedInputCount, postRecalcDirectFormulaIndices, {
          lookupHandledInputCellIndices,
        })
      ) {
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
        args.deferKernelSync(changedInputArray)
        didFastDeferKernelSyncOnly = true
      }
    }
    if (
      !didFastDeferKernelSyncOnly &&
      ((hasActiveFormulas && (hasRecalcWork || hasVolatileFormulaWork)) || (hasActivePivots && hasRecalcWork) || shouldRefreshPivots)
    ) {
      formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount)
      const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
      if (changedInputsNeedRegionQueryIndices(changedInputArray, changedInputCount, postRecalcDirectFormulaIndices)) {
        args.prepareRegionQueryIndices()
      }
      const canUseKernelSyncOnlyRecalc =
        formulaChangedCount === 0 &&
        changedInputCount > 0 &&
        precomputedKernelSyncCellIndices.length === 0 &&
        !refreshAllPivots &&
        canSkipDirtyTraversalForChangedInputs(changedInputArray, changedInputCount, postRecalcDirectFormulaIndices, {
          lookupHandledInputCellIndices,
        })
      const canDeferKernelSyncOnlyRecalc = canUseKernelSyncOnlyRecalc && postRecalcDirectFormulaIndices.size === 0
      const canSkipKernelSyncOnlyRecalc = canUseKernelSyncOnlyRecalc && postRecalcDirectFormulaIndices.size > 0
      const canSkipRecalcForDirectDeltas = canSkipKernelSyncOnlyRecalc && hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices)
      const canSkipRecalcForDirectEvaluation =
        canSkipKernelSyncOnlyRecalc && canEvaluatePostRecalcDirectFormulasWithoutKernel(args.state.formulas, postRecalcDirectFormulaIndices)
      const canUseDisjointDirectEventChanges =
        (canSkipRecalcForDirectDeltas || canSkipRecalcForDirectEvaluation) &&
        explicitChangedCount === changedInputCount &&
        !hasActivePivots &&
        !refreshAllPivots &&
        directFormulaChangesAreDisjointFromInputs(changedInputArray, changedInputCount, postRecalcDirectFormulaIndices)
      if (canDeferKernelSyncOnlyRecalc) {
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
        args.deferKernelSync(changedInputArray)
      } else if (!canSkipKernelSyncOnlyRecalc) {
        args.prepareRegionQueryIndices()
        const changedRoots = canUseKernelSyncOnlyRecalc
          ? new Uint32Array()
          : args.composeMutationRoots(changedInputCount, formulaChangedCount)
        const kernelSyncRoots =
          precomputedKernelSyncCellIndices.length === 0
            ? changedInputArray
            : Uint32Array.from([...changedInputArray, ...precomputedKernelSyncCellIndices])
        recalculated = args.recalculate(changedRoots, kernelSyncRoots)
        didRunRecalc = true
      } else if (canSkipRecalcForDirectDeltas) {
        countDirectFormulaDeltaSkip(args.state.formulas, postRecalcDirectFormulaIndices, args.state.counters)
        args.deferKernelSync(changedInputArray)
      } else if (canSkipRecalcForDirectEvaluation) {
        addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
        args.deferKernelSync(changedInputArray)
      } else {
        args.prepareRegionQueryIndices()
        args.recalculate(new Uint32Array(), changedInputArray)
      }
      if (postRecalcDirectFormulaIndices.size > 0) {
        const singleDirectChanged = tryApplySinglePostRecalcDirectFormula(
          postRecalcDirectFormulaIndices,
          didRunRecalc,
          postRecalcDirectFormulaMetrics,
        )
        if (singleDirectChanged !== undefined) {
          recalculated = mergeChangedCellIndices(recalculated, singleDirectChanged)
        } else {
          const constantScalarChanged = !didRunRecalc ? tryApplyDirectScalarDeltas(postRecalcDirectFormulaIndices) : undefined
          if (constantScalarChanged !== undefined) {
            recalculated = mergeChangedCellIndices(recalculated, constantScalarChanged)
          } else {
            const directDeltaChanged = !didRunRecalc ? tryApplyDirectFormulaDeltas(postRecalcDirectFormulaIndices) : undefined
            if (directDeltaChanged !== undefined) {
              recalculated = mergeChangedCellIndices(recalculated, directDeltaChanged)
            } else {
              const postRecalcChanged = new Uint32Array(postRecalcDirectFormulaIndices.size)
              let postRecalcChangedCount = 0
              let postRecalcExtraChanged: number[] | undefined
              let directAggregateDeltaApplicationCount = 0
              let directScalarDeltaApplicationCount = 0
              args.state.workbook.withBatchedColumnVersionUpdates(() => {
                postRecalcDirectFormulaIndices.forEachIndexed((cellIndex, directIndex) => {
                  if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
                    return
                  }
                  const currentResult = postRecalcDirectFormulaIndices.getCurrentResultAt(directIndex)
                  if (!didRunRecalc && currentResult !== undefined && applyDirectFormulaCurrentResult(cellIndex, currentResult)) {
                    postRecalcChanged[postRecalcChangedCount++] = cellIndex
                    return
                  }
                  const delta = postRecalcDirectFormulaIndices.getDeltaAt(directIndex)
                  if (!didRunRecalc && delta !== undefined && applyDirectFormulaNumericDelta(cellIndex, delta)) {
                    const formula = args.state.formulas.get(cellIndex)
                    if (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) {
                      directAggregateDeltaApplicationCount += 1
                    }
                    if (formula?.directScalar !== undefined) {
                      directScalarDeltaApplicationCount += 1
                    }
                    postRecalcChanged[postRecalcChangedCount++] = cellIndex
                    return
                  }
                  const formula = args.state.formulas.get(cellIndex)
                  if (
                    !didRunRecalc &&
                    formula?.directScalar !== undefined &&
                    !formula.compiled.producesSpill &&
                    applyDirectScalarCurrentValue(cellIndex, formula.directScalar)
                  ) {
                    countPostRecalcDirectFormulaMetric(cellIndex, postRecalcDirectFormulaMetrics)
                    postRecalcChanged[postRecalcChangedCount++] = cellIndex
                    return
                  }
                  countPostRecalcDirectFormulaMetric(cellIndex, postRecalcDirectFormulaMetrics)
                  const changedCellIndices = args.evaluateDirectFormula(cellIndex)
                  postRecalcChanged[postRecalcChangedCount++] = cellIndex
                  if (changedCellIndices) {
                    postRecalcExtraChanged ??= []
                    for (let index = 0; index < changedCellIndices.length; index += 1) {
                      postRecalcExtraChanged.push(changedCellIndices[index]!)
                    }
                  }
                })
              })
              if (directAggregateDeltaApplicationCount > 0) {
                addEngineCounter(args.state.counters, 'directAggregateDeltaApplications', directAggregateDeltaApplicationCount)
              }
              if (directScalarDeltaApplicationCount > 0) {
                addEngineCounter(args.state.counters, 'directScalarDeltaApplications', directScalarDeltaApplicationCount)
              }
              const directChanged = postRecalcChanged.subarray(0, postRecalcChangedCount)
              recalculated =
                postRecalcExtraChanged && postRecalcExtraChanged.length > 0
                  ? mergeChangedCellIndices(recalculated, mergeChangedCellIndices(directChanged, postRecalcExtraChanged))
                  : mergeChangedCellIndices(recalculated, directChanged)
            }
          }
        }
      }
      if (hasActivePivots || shouldRefreshPivots) {
        const pivotRefreshRoots =
          shouldRefreshPivots || changedInputArray.length === 0 ? recalculated : mergeChangedCellIndices(recalculated, changedInputArray)
        recalculated = args.reconcilePivotOutputs(pivotRefreshRoots, shouldRefreshPivots)
      } else if (canUseDisjointDirectEventChanges) {
        canComposeDisjointEventChanges = true
      }
    }
    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const invalidation = isRestore || sheetDeleted || structuralInvalidation ? 'full' : 'cells'
    const changed: U32 =
      isRestore || invalidation === 'full' || !requiresChangedSet
        ? new Uint32Array()
        : canComposeDisjointEventChanges
          ? args.composeDisjointEventChanges(recalculated, explicitChangedCount)
          : args.composeEventChanges(recalculated, explicitChangedCount)
    if (
      hasTrackedEventListeners &&
      canComposeDisjointEventChanges &&
      changed.length > 4 &&
      explicitChangedCount > 0 &&
      explicitChangedCount < changed.length
    ) {
      const sheetId = args.state.workbook.cellStore.sheetIds[changed[0]!]
      if (sheetId !== undefined && canTrustPhysicalTrackedChangeSplit(changed, sheetId, explicitChangedCount)) {
        tagTrustedPhysicalTrackedChanges(changed, sheetId, explicitChangedCount)
      }
    }
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      ...(didRunRecalc
        ? {
            dirtyFormulaCount: previousMetrics.dirtyFormulaCount,
            wasmFormulaCount: previousMetrics.wasmFormulaCount + postRecalcDirectFormulaMetrics.wasmFormulaCount,
            jsFormulaCount: previousMetrics.jsFormulaCount + postRecalcDirectFormulaMetrics.jsFormulaCount,
          }
        : {
            dirtyFormulaCount: 0,
            wasmFormulaCount: postRecalcDirectFormulaMetrics.wasmFormulaCount,
            jsFormulaCount: postRecalcDirectFormulaMetrics.jsFormulaCount,
            rangeNodeVisits: 0,
            recalcMs: 0,
          }),
      batchId: previousMetrics.batchId + 1,
      changedInputCount: changedInputCount + formulaChangedCount,
      compileMs,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasGeneralEventListeners || hasWatchedCellListeners) {
      const shouldMaterializeChangedCells =
        hasGeneralEventListeners &&
        invalidation !== 'full' &&
        (changed.length <= GENERAL_CHANGED_CELL_PAYLOAD_LIMIT ||
          (invalidatedRanges.length === 0 && invalidatedRows.length === 0 && invalidatedColumns.length === 0))
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation,
        changedCellIndices: changed,
        changedCells: shouldMaterializeChangedCells ? args.captureChangedCells(changed) : [],
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
      const patchRequest = {
        invalidation,
        invalidatedRanges,
        invalidatedRows,
        invalidatedColumns,
      } satisfies Parameters<typeof args.captureChangedPatches>[1]
      const shouldCapturePatches =
        patchRequest.invalidation !== 'cells' ||
        patchRequest.invalidatedRanges.length > 0 ||
        patchRequest.invalidatedRows.length > 0 ||
        patchRequest.invalidatedColumns.length > 0
      const patches = shouldCapturePatches ? args.captureChangedPatches(changed, patchRequest) : undefined
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation,
        changedCellIndices: changed,
        ...(patches ? { patches } : {}),
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
    if (tryApplySingleExistingDirectLiteralMutation(refs, batch, source)) {
      return
    }
    if (tryApplyCoalescedDirectScalarLiteralBatch(refs, batch, source, potentialNewCells)) {
      return
    }
    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    let changedInputCount = 0
    let formulaChangedCount = 0
    let explicitChangedCount = 0
    let topologyChanged = false
    let compileMs = 0
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()
    const postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts = {
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    }
    const lookupHandledInputCellIndices: number[] = []
    const pendingExactLookupInvalidations = new Map<number, { sheetName: string; col: number }>()
    const pendingSortedLookupInvalidations = new Map<number, { sheetName: string; col: number }>()
    const queueHandledLookupInvalidation = (sheetId: number, sheetName: string, col: number, exact: boolean, sorted: boolean): void => {
      const key = aggregateColumnDependencyKey(sheetId, col)
      if (exact) {
        pendingExactLookupInvalidations.set(key, { sheetName, col })
      }
      if (sorted) {
        pendingSortedLookupInvalidations.set(key, { sheetName, col })
      }
    }
    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const trackExplicitChanges = !isRestore && requiresChangedSet
    let hadCycleMembersBefore: boolean | undefined
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
    const reservedNewCells = potentialNewCells ?? refs.length
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + reservedNewCells + 1)
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
    const exactLookupImpactCaches: ExactLookupImpactCaches = new Map()
    const clearTrackedColumnDependencyFlagCache = (): void => {
      trackedColumnDependencyFlagsBySheet.clear()
      exactLookupImpactCaches.clear()
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
    const resolveExistingMutationCellIndex = (ref: EngineCellMutationRef): number | undefined => {
      const candidate = ref.cellIndex
      const { sheetId, mutation } = ref
      if (candidate !== undefined && args.state.workbook.cellStore.sheetIds[candidate] === sheetId) {
        const sheet = args.state.workbook.getSheetById(sheetId)
        if (sheet?.structureVersion === 1) {
          if (
            args.state.workbook.cellStore.rows[candidate] === mutation.row &&
            args.state.workbook.cellStore.cols[candidate] === mutation.col
          ) {
            return candidate
          }
        } else {
          const position = sheet?.logical.getCellVisiblePosition(candidate)
          if (position?.row === mutation.row && position.col === mutation.col) {
            return candidate
          }
        }
      }
      return args.state.workbook.cellKeyToIndex.get(makeCellKey(sheetId, mutation.row, mutation.col))
    }

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        refs.forEach((ref, refIndex) => {
          const { sheetId, mutation } = ref
          const order = args.state.trackReplicaVersions && batch ? batchOpOrder(batch, refIndex) : undefined
          const existingIndex = resolveExistingMutationCellIndex(ref)

          switch (mutation.kind) {
            case 'setCellValue': {
              const sheetName = resolveSheetName(sheetId)
              const { hasExactLookupDependents, hasSortedLookupDependents, hasAggregateDependents } = resolveTrackedColumnDependencyFlags(
                sheetId,
                mutation.col,
              )
              if (mutation.value === null && !isRestore && (existingIndex === undefined || isNullLiteralWriteNoOp(existingIndex))) {
                break
              }
              const canFastOverwriteExisting = existingIndex !== undefined && canFastPathLiteralOverwrite(existingIndex)
              const needsDirectLookupNumericValue = canFastOverwriteExisting
              const oldExactLookupNumber =
                canFastOverwriteExisting && hasExactLookupDependents ? readExactNumericValueForLookup(existingIndex) : undefined
              const newExactLookupNumber =
                hasExactLookupDependents || needsDirectLookupNumericValue ? exactLookupLiteralNumericValue(mutation.value) : undefined
              const oldApproximateLookupNumber =
                canFastOverwriteExisting && hasSortedLookupDependents ? readApproximateNumericValueForLookup(existingIndex) : undefined
              const newApproximateLookupNumber =
                hasSortedLookupDependents || needsDirectLookupNumericValue ? directScalarLiteralNumericValue(mutation.value) : undefined
              const exactLookupDependentsHandled =
                !isRestore &&
                hasExactLookupDependents &&
                !hasAggregateDependents &&
                oldExactLookupNumber !== undefined &&
                newExactLookupNumber !== undefined &&
                canSkipExactLookupNumericColumnWrite(sheetId, mutation.col, mutation.row, oldExactLookupNumber, newExactLookupNumber)
              const sortedLookupDependentsHandled =
                !isRestore &&
                hasSortedLookupDependents &&
                oldApproximateLookupNumber !== undefined &&
                newApproximateLookupNumber !== undefined &&
                canSkipApproximateLookupNumericColumnWrite(
                  sheetId,
                  sheetName,
                  mutation.col,
                  mutation.row,
                  oldApproximateLookupNumber,
                  newApproximateLookupNumber,
                )
              const needsLookupValueRead =
                hasAggregateDependents ||
                (hasExactLookupDependents && !exactLookupDependentsHandled) ||
                (hasSortedLookupDependents && !sortedLookupDependentsHandled)
              const needsLookupOwnerInvalidation =
                (hasExactLookupDependents && exactLookupDependentsHandled) || (hasSortedLookupDependents && sortedLookupDependentsHandled)
              let directDependentsHandled = false
              if (!isRestore && canFastOverwriteExisting) {
                const oldNumber = directScalarCellNumericValue(existingIndex)
                const newNumber = directScalarLiteralNumericValue(mutation.value)
                if (oldNumber !== undefined && newNumber !== undefined) {
                  directDependentsHandled = markPostRecalcDirectScalarNumericDependents(
                    existingIndex,
                    oldNumber,
                    newNumber,
                    postRecalcDirectFormulaIndices,
                    newExactLookupNumber,
                    newApproximateLookupNumber,
                  )
                }
              }
              const canUseDirectLookupCurrent =
                !isRestore &&
                canFastOverwriteExisting &&
                (newExactLookupNumber !== undefined || newApproximateLookupNumber !== undefined) &&
                !needsLookupValueRead &&
                !directDependentsHandled
              if (canUseDirectLookupCurrent) {
                directDependentsHandled = markPostRecalcDirectLookupCurrentDependentsFromNumeric(
                  existingIndex,
                  newExactLookupNumber,
                  newApproximateLookupNumber,
                  postRecalcDirectFormulaIndices,
                )
              }
              let prior = needsLookupValueRead || !directDependentsHandled ? readCellValueForLookup(existingIndex) : undefined
              if (canFastOverwriteExisting) {
                writeLiteralToCellStore(args.state.workbook.cellStore, existingIndex, mutation.value, args.state.strings)
                args.state.workbook.notifyCellValueWritten(existingIndex)
                if (needsLookupOwnerInvalidation) {
                  queueHandledLookupInvalidation(
                    sheetId,
                    sheetName,
                    mutation.col,
                    hasExactLookupDependents && exactLookupDependentsHandled,
                    hasSortedLookupDependents && sortedLookupDependentsHandled,
                  )
                  if (!needsLookupValueRead) {
                    lookupHandledInputCellIndices.push(existingIndex)
                  }
                }
                const newValue =
                  needsLookupValueRead || !directDependentsHandled ? literalToValue(mutation.value, args.state.strings) : undefined
                if (!isRestore && !directDependentsHandled && newValue) {
                  prior ??= readCellValueForLookup(existingIndex)
                  const genericDirectDependentsHandled = markPostRecalcDirectFormulaDependents(
                    existingIndex,
                    postRecalcDirectFormulaIndices,
                    prior.value,
                    newValue,
                  )
                  if (!genericDirectDependentsHandled) {
                    markDirectScalarDeltaClosure(existingIndex, prior.value, newValue, postRecalcDirectFormulaIndices)
                  }
                }
                if (needsLookupValueRead) {
                  const newStringId =
                    typeof mutation.value === 'string' ? args.state.workbook.cellStore.stringIds[existingIndex] : undefined
                  const priorLookup = prior ?? readCellValueForLookup(existingIndex)
                  const newLookupValue = newValue ?? literalToValue(mutation.value, args.state.strings)
                  if (hasExactLookupDependents || hasAggregateDependents) {
                    const exactLookupRequest = withOptionalLookupStringIds({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                      oldValue: priorLookup.value,
                      newValue: newLookupValue,
                      oldStringId: priorLookup.stringId,
                      newStringId,
                      inputCellIndex: existingIndex,
                    })
                    if (hasExactLookupDependents && !exactLookupDependentsHandled) {
                      formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                        exactLookupRequest,
                        formulaChangedCount,
                        exactLookupImpactCaches,
                      )
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
                      oldValue: priorLookup.value,
                      newValue: newLookupValue,
                      oldStringId: priorLookup.stringId,
                      newStringId,
                    })
                    if (!sortedLookupDependentsHandled) {
                      formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
                    }
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
              const newValue =
                needsLookupValueRead || !directDependentsHandled ? literalToValue(mutation.value, args.state.strings) : undefined
              if (!isRestore && !directDependentsHandled && newValue) {
                prior ??= readCellValueForLookup(existingIndex)
                const genericDirectDependentsHandled = markPostRecalcDirectFormulaDependents(
                  cellIndex,
                  postRecalcDirectFormulaIndices,
                  prior.value,
                  newValue,
                )
                if (!genericDirectDependentsHandled) {
                  markDirectScalarDeltaClosure(cellIndex, prior.value, newValue, postRecalcDirectFormulaIndices)
                }
              }
              if (needsLookupValueRead) {
                const newStringId = typeof mutation.value === 'string' ? args.state.workbook.cellStore.stringIds[cellIndex] : undefined
                const priorLookup = prior ?? readCellValueForLookup(existingIndex)
                const newLookupValue = newValue ?? literalToValue(mutation.value, args.state.strings)
                if (hasExactLookupDependents || hasAggregateDependents) {
                  const exactLookupRequest = withOptionalLookupStringIds({
                    sheetName,
                    row: mutation.row,
                    col: mutation.col,
                    oldValue: priorLookup.value,
                    newValue: newLookupValue,
                    oldStringId: priorLookup.stringId,
                    newStringId,
                    inputCellIndex: cellIndex,
                  })
                  if (hasExactLookupDependents && !exactLookupDependentsHandled) {
                    formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                      exactLookupRequest,
                      formulaChangedCount,
                      exactLookupImpactCaches,
                    )
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
                    oldValue: priorLookup.value,
                    newValue: newLookupValue,
                    oldStringId: priorLookup.stringId,
                    newStringId,
                  })
                  if (!sortedLookupDependentsHandled) {
                    formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
                  }
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
              const { hasExactLookupDependents, hasSortedLookupDependents, hasAggregateDependents } = resolveTrackedColumnDependencyFlags(
                sheetId,
                mutation.col,
              )
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
              const priorHadFormula = args.state.formulas.get(cellIndex) !== undefined
              const oldFormulaNumber = !isRestore && priorHadFormula ? readExactNumericValueForLookup(cellIndex) : undefined
              const compileStarted = isRestore ? 0 : performance.now()
              try {
                const changedTopology = args.bindFormula(cellIndex, sheetName, mutation.formula)
                args.invalidateAggregateColumn({ sheetName, col: mutation.col })
                if (!isRestore) {
                  compileMs += performance.now() - compileStarted
                }
                clearTrackedColumnDependencyFlagCache()
                changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
                const handledFormulaReplacementAsDirectDelta =
                  priorHadFormula &&
                  !hasExactLookupDependents &&
                  !hasSortedLookupDependents &&
                  !hasAggregateDependents &&
                  tryApplyFormulaReplacementAsDirectScalarDeltaRoot({
                    cellIndex,
                    oldNumber: oldFormulaNumber,
                    changedTopology,
                    postRecalcDirectFormulaIndices,
                    postRecalcDirectFormulaMetrics,
                  })
                if (!handledFormulaReplacementAsDirectDelta) {
                  formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
                }
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
              const prior = readCellValueForLookup(existingIndex)
              if (existingIndex !== undefined && isClearCellNoOp(existingIndex)) {
                break
              }
              if (existingIndex !== undefined && canFastPathLiteralOverwrite(existingIndex)) {
                args.state.workbook.cellStore.setValue(existingIndex, emptyValue())
                args.state.workbook.notifyCellValueWritten(existingIndex)
                if (!isRestore) {
                  const nextValue = emptyValue()
                  const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                    existingIndex,
                    postRecalcDirectFormulaIndices,
                    prior.value,
                    nextValue,
                  )
                  if (!directDependentsHandled) {
                    markDirectScalarDeltaClosure(existingIndex, prior.value, nextValue, postRecalcDirectFormulaIndices)
                  }
                }
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
                      inputCellIndex: existingIndex,
                    })
                    if (hasExactLookupDependents) {
                      formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                        exactLookupRequest,
                        formulaChangedCount,
                        exactLookupImpactCaches,
                      )
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
                    formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
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
              if (!isRestore) {
                const nextValue = emptyValue()
                const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                  existingIndex,
                  postRecalcDirectFormulaIndices,
                  prior.value,
                  nextValue,
                )
                if (!directDependentsHandled) {
                  markDirectScalarDeltaClosure(existingIndex, prior.value, nextValue, postRecalcDirectFormulaIndices)
                }
              }
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
                    inputCellIndex: existingIndex,
                  })
                  if (hasExactLookupDependents) {
                    formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                      exactLookupRequest,
                      formulaChangedCount,
                      exactLookupImpactCaches,
                    )
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
                  formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
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
    }

    pendingExactLookupInvalidations.forEach((entry) => args.invalidateExactLookupColumn(entry))
    pendingSortedLookupInvalidations.forEach((entry) => args.invalidateSortedLookupColumn(entry))

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
        !hadCycleMembersBeforeNow() &&
        formulaChangedCount > 0 &&
        args.repairTopoRanks(args.getChangedFormulaBuffer().subarray(0, formulaChangedCount))
      if (!repaired) {
        args.rebuildTopoRanks()
        args.detectCycles()
        changedInputCount = markCycleMemberInputsChanged(changedInputCount)
      }
    }
    const hasActiveFormulas = args.state.formulas.size > 0
    const hasActivePivots = args.state.workbook.hasPivots()
    const hasRecalcWork = changedInputCount > 0 || formulaChangedCount > 0 || postRecalcDirectFormulaIndices.size > 0
    const hasVolatileFormulaWork = hasActiveFormulas && (args.hasVolatileFormulas ? args.hasVolatileFormulas() : true)
    let recalculated: U32 = new Uint32Array()
    let didRunRecalc = false
    let didFastDeferKernelSyncOnly = false
    let canComposeDisjointEventChanges = false
    if (
      hasActiveFormulas &&
      changedInputCount > 0 &&
      formulaChangedCount === 0 &&
      postRecalcDirectFormulaIndices.size === 0 &&
      !hasActivePivots &&
      !hasVolatileFormulaWork
    ) {
      const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
      if (
        canSkipDirtyTraversalForChangedInputs(changedInputArray, changedInputCount, postRecalcDirectFormulaIndices, {
          lookupHandledInputCellIndices,
        })
      ) {
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
        args.deferKernelSync(changedInputArray)
        didFastDeferKernelSyncOnly = true
      }
    }
    if (
      !didFastDeferKernelSyncOnly &&
      ((hasActiveFormulas && (hasRecalcWork || hasVolatileFormulaWork)) || (hasActivePivots && hasRecalcWork))
    ) {
      formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount)
      const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
      if (changedInputsNeedRegionQueryIndices(changedInputArray, changedInputCount, postRecalcDirectFormulaIndices)) {
        args.prepareRegionQueryIndices()
      }
      const canUseKernelSyncOnlyRecalc =
        formulaChangedCount === 0 &&
        changedInputCount > 0 &&
        canSkipDirtyTraversalForChangedInputs(changedInputArray, changedInputCount, postRecalcDirectFormulaIndices, {
          lookupHandledInputCellIndices,
        })
      const canDeferKernelSyncOnlyRecalc = canUseKernelSyncOnlyRecalc && postRecalcDirectFormulaIndices.size === 0
      const canSkipKernelSyncOnlyRecalc = canUseKernelSyncOnlyRecalc && postRecalcDirectFormulaIndices.size > 0
      const canSkipRecalcForDirectDeltas = canSkipKernelSyncOnlyRecalc && hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices)
      const canSkipRecalcForDirectEvaluation =
        canSkipKernelSyncOnlyRecalc && canEvaluatePostRecalcDirectFormulasWithoutKernel(args.state.formulas, postRecalcDirectFormulaIndices)
      const canUseDisjointDirectEventChanges =
        (canSkipRecalcForDirectDeltas || canSkipRecalcForDirectEvaluation) &&
        explicitChangedCount === changedInputCount &&
        !hasActivePivots &&
        directFormulaChangesAreDisjointFromInputs(changedInputArray, changedInputCount, postRecalcDirectFormulaIndices)
      if (canDeferKernelSyncOnlyRecalc) {
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
        args.deferKernelSync(changedInputArray)
      } else if (!canSkipKernelSyncOnlyRecalc) {
        args.prepareRegionQueryIndices()
        const changedRoots = canUseKernelSyncOnlyRecalc
          ? new Uint32Array()
          : args.composeMutationRoots(changedInputCount, formulaChangedCount)
        recalculated = args.recalculate(changedRoots, changedInputArray)
        didRunRecalc = true
      } else if (canSkipRecalcForDirectDeltas) {
        countDirectFormulaDeltaSkip(args.state.formulas, postRecalcDirectFormulaIndices, args.state.counters)
        args.deferKernelSync(changedInputArray)
      } else if (canSkipRecalcForDirectEvaluation) {
        addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
        args.deferKernelSync(changedInputArray)
      } else {
        args.prepareRegionQueryIndices()
        args.recalculate(new Uint32Array(), changedInputArray)
      }
      if (postRecalcDirectFormulaIndices.size > 0) {
        const singleDirectChanged = tryApplySinglePostRecalcDirectFormula(
          postRecalcDirectFormulaIndices,
          didRunRecalc,
          postRecalcDirectFormulaMetrics,
        )
        if (singleDirectChanged !== undefined) {
          recalculated = mergeChangedCellIndices(recalculated, singleDirectChanged)
        } else {
          const constantScalarChanged = !didRunRecalc ? tryApplyDirectScalarDeltas(postRecalcDirectFormulaIndices) : undefined
          if (constantScalarChanged !== undefined) {
            recalculated = mergeChangedCellIndices(recalculated, constantScalarChanged)
          } else {
            const directDeltaChanged = !didRunRecalc ? tryApplyDirectFormulaDeltas(postRecalcDirectFormulaIndices) : undefined
            if (directDeltaChanged !== undefined) {
              recalculated = mergeChangedCellIndices(recalculated, directDeltaChanged)
            } else {
              const postRecalcChanged = new Uint32Array(postRecalcDirectFormulaIndices.size)
              let postRecalcChangedCount = 0
              let postRecalcExtraChanged: number[] | undefined
              let directAggregateDeltaApplicationCount = 0
              let directScalarDeltaApplicationCount = 0
              args.state.workbook.withBatchedColumnVersionUpdates(() => {
                postRecalcDirectFormulaIndices.forEachIndexed((cellIndex, directIndex) => {
                  if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
                    return
                  }
                  const currentResult = postRecalcDirectFormulaIndices.getCurrentResultAt(directIndex)
                  if (!didRunRecalc && currentResult !== undefined && applyDirectFormulaCurrentResult(cellIndex, currentResult)) {
                    postRecalcChanged[postRecalcChangedCount++] = cellIndex
                    return
                  }
                  const delta = postRecalcDirectFormulaIndices.getDeltaAt(directIndex)
                  if (!didRunRecalc && delta !== undefined && applyDirectFormulaNumericDelta(cellIndex, delta)) {
                    const formula = args.state.formulas.get(cellIndex)
                    if (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) {
                      directAggregateDeltaApplicationCount += 1
                    }
                    if (formula?.directScalar !== undefined) {
                      directScalarDeltaApplicationCount += 1
                    }
                    postRecalcChanged[postRecalcChangedCount++] = cellIndex
                    return
                  }
                  const formula = args.state.formulas.get(cellIndex)
                  if (
                    !didRunRecalc &&
                    formula?.directScalar !== undefined &&
                    !formula.compiled.producesSpill &&
                    applyDirectScalarCurrentValue(cellIndex, formula.directScalar)
                  ) {
                    countPostRecalcDirectFormulaMetric(cellIndex, postRecalcDirectFormulaMetrics)
                    postRecalcChanged[postRecalcChangedCount++] = cellIndex
                    return
                  }
                  countPostRecalcDirectFormulaMetric(cellIndex, postRecalcDirectFormulaMetrics)
                  const changedCellIndices = args.evaluateDirectFormula(cellIndex)
                  postRecalcChanged[postRecalcChangedCount++] = cellIndex
                  if (changedCellIndices) {
                    postRecalcExtraChanged ??= []
                    for (let index = 0; index < changedCellIndices.length; index += 1) {
                      postRecalcExtraChanged.push(changedCellIndices[index]!)
                    }
                  }
                })
              })
              if (directAggregateDeltaApplicationCount > 0) {
                addEngineCounter(args.state.counters, 'directAggregateDeltaApplications', directAggregateDeltaApplicationCount)
              }
              if (directScalarDeltaApplicationCount > 0) {
                addEngineCounter(args.state.counters, 'directScalarDeltaApplications', directScalarDeltaApplicationCount)
              }
              const directChanged = postRecalcChanged.subarray(0, postRecalcChangedCount)
              recalculated =
                postRecalcExtraChanged && postRecalcExtraChanged.length > 0
                  ? mergeChangedCellIndices(recalculated, mergeChangedCellIndices(directChanged, postRecalcExtraChanged))
                  : mergeChangedCellIndices(recalculated, directChanged)
            }
          }
        }
      }
      if (hasActivePivots) {
        const pivotRefreshRoots = changedInputArray.length === 0 ? recalculated : mergeChangedCellIndices(recalculated, changedInputArray)
        recalculated = args.reconcilePivotOutputs(pivotRefreshRoots, false)
      } else if (canUseDisjointDirectEventChanges) {
        canComposeDisjointEventChanges = true
      }
    }
    const changed: U32 =
      isRestore || !requiresChangedSet
        ? new Uint32Array()
        : canComposeDisjointEventChanges
          ? args.composeDisjointEventChanges(recalculated, explicitChangedCount)
          : args.composeEventChanges(recalculated, explicitChangedCount)
    if (
      hasTrackedEventListeners &&
      canComposeDisjointEventChanges &&
      changed.length > 4 &&
      explicitChangedCount > 0 &&
      explicitChangedCount < changed.length
    ) {
      const sheetId = args.state.workbook.cellStore.sheetIds[changed[0]!]
      if (sheetId !== undefined && canTrustPhysicalTrackedChangeSplit(changed, sheetId, explicitChangedCount)) {
        tagTrustedPhysicalTrackedChanges(changed, sheetId, explicitChangedCount)
      }
    }
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      ...(didRunRecalc
        ? {
            dirtyFormulaCount: previousMetrics.dirtyFormulaCount,
            wasmFormulaCount: previousMetrics.wasmFormulaCount + postRecalcDirectFormulaMetrics.wasmFormulaCount,
            jsFormulaCount: previousMetrics.jsFormulaCount + postRecalcDirectFormulaMetrics.jsFormulaCount,
          }
        : {
            dirtyFormulaCount: 0,
            wasmFormulaCount: postRecalcDirectFormulaMetrics.wasmFormulaCount,
            jsFormulaCount: postRecalcDirectFormulaMetrics.jsFormulaCount,
            rangeNodeVisits: 0,
            recalcMs: 0,
          }),
      batchId: previousMetrics.batchId + 1,
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
      const patchRequest = {
        invalidation,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
      } satisfies Parameters<typeof args.captureChangedPatches>[1]
      const shouldCapturePatches =
        patchRequest.invalidation !== 'cells' ||
        patchRequest.invalidatedRanges.length > 0 ||
        patchRequest.invalidatedRows.length > 0 ||
        patchRequest.invalidatedColumns.length > 0
      const patches = shouldCapturePatches ? args.captureChangedPatches(changed, patchRequest) : undefined
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation,
        changedCellIndices: changed,
        ...(patches ? { patches } : {}),
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

  const __testHooks: Record<string, unknown> = ENGINE_OPERATION_TEST_HOOKS_ENABLED
    ? {
        assertProtectionAllowsOp,
        canPatchUniformLookupTailWrite,
        canSkipApproximateLookupDirtyMark,
        canSkipApproximateLookupNewNumericColumnWrite,
        canSkipApproximateLookupNumericColumnWrite,
        canSkipExactLookupNumericColumnWrite,
        collectAffectedDirectRangeDependents,
        collectSingleAffectedDirectRangeDependent,
        directCriteriaMatchesChangedAggregateRow,
        entityKeyForOp,
        isLocallySortedNumericWrite,
        isLocallySortedTextWrite,
        markAffectedApproximateLookupDependents,
        markAffectedDirectRangeDependents,
        patchUniformLookupTailWrites,
        planApproximateLookupNumericColumnWrite,
        planExactLookupNumericColumnWrite,
        planSingleApproximateLookupNumericColumnWrite,
        planSingleExactLookupNumericColumnWrite,
        rangeIsProtected,
        readApproximateNumericValueAtForLookup,
        readApproximateNumericValueForLookup,
        readCellValueAtForLookup,
        readCellValueForLookup,
        readDirectCriteriaOperandValue,
        readExactNumericValueForLookup,
        sheetDeleteBarrierForOp,
        sheetHasProtection,
        shouldApplyOp,
        tryApplyDenseRowPairDirectScalarLiteralBatch,
        tryApplyLookupOnlyNumericColumnLiteralBatch,
        tryApplySingleExistingDirectLiteralMutation,
        tryApplySingleDirectAggregateLiteralMutationFastPath,
        tryApplySingleDirectFormulaLiteralMutationWithoutEvents,
        tryApplySingleDirectLookupOperandMutationFastPath,
        tryApplySingleDirectScalarLiteralMutationWithoutEvents,
        tryApplySingleKernelSyncOnlyLiteralMutationFastPath,
        tryDirectCriteriaSumDelta,
      }
    : {}

  return {
    __testHooks,
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
    applyCellMutationsAtNow,
    applyExistingNumericCellMutationAtNow,
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
