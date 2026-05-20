import { parseCellAddress, type CompiledFormula } from '@bilig/formula'
import type { CellValue, WorkbookSnapshot } from '@bilig/protocol'
import type {
  EngineCellMutationRef,
  EngineFormulaSourceRef,
  EngineFormulaSourceRefs,
  EngineFormulaSourceRefTable,
} from '../cell-mutations-at.js'
import { CellFlags } from '../cell-store.js'
import { writeLiteralToCellStore } from '../engine-value-utils.js'
import type { DeferredInitialFormulaFamilyRun } from '../engine/services/formula-initialization-family-runs.js'
import type { InitialFormulaEntryRefSource } from '../engine/services/formula-initialization-refs.js'
import type { FormulaInstanceSnapshot } from '../formula/formula-instance-table.js'
import type { FormulaTemplateResolution, FormulaTemplateSnapshot } from '../formula/template-bank.js'
import { collectDefinedFormulaNames, formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc } from './unsupported-formula-cache.js'
import type { StringPool } from '../string-pool.js'
import type { SheetRecord, WorkbookStore } from '../workbook-store.js'
import {
  createWrittenColumnTracker,
  markWrittenColumn,
  materializeWrittenColumns,
  type WrittenColumnTracker,
} from '../written-column-tracker.js'
import { restoreAlignedRuntimeFormulaFamilyRuns, type RuntimeImageFormulaFamilyRunSnapshot } from './runtime-image-formula-family-runs.js'
import { formulaCachedLiteralToRestoredValue, restoreLiteralCell } from './runtime-image-literal-restore.js'
import { restoreVisualMetadata, restoreWorkbookStructure } from './runtime-image-metadata-restore.js'

type WorkbookSnapshotCell = WorkbookSnapshot['sheets'][number]['cells'][number]

export interface RuntimeImage {
  readonly version: 1
  readonly templateBank: readonly FormulaTemplateSnapshot[]
  readonly formulaInstances: readonly FormulaInstanceSnapshot[]
  readonly formulaValues: readonly RuntimeImageFormulaValueSnapshot[]
  readonly formulaFamilyRuns?: readonly RuntimeImageFormulaFamilyRunSnapshot[]
  readonly sheetCells?: readonly RuntimeImageSheetCellsSnapshot[]
}

export interface RuntimeImageFormulaValueSnapshot {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly value: CellValue
}

export interface RuntimeImageSheetCellsSnapshot {
  readonly sheetName: string
  readonly coords: readonly RuntimeImageCellCoordinateSnapshot[]
  readonly coordinateOrder?: 'dense-row-major'
  readonly dimensions?: RuntimeImageSheetDimensionsSnapshot
  readonly cellCount?: number
}

export interface RuntimeImageCellCoordinateSnapshot {
  readonly row: number
  readonly col: number
}

export interface RuntimeImageSheetDimensionsSnapshot {
  readonly width: number
  readonly height: number
}

export interface RuntimeImageRestoreArgs {
  readonly snapshot: WorkbookSnapshot
  readonly runtimeImage: RuntimeImage
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly resetWorkbook: (workbookName?: string) => void
  readonly checkEvaluationBudget?: (stepCost?: number) => void
  readonly hydrateTemplateBank: (templates: readonly FormulaTemplateSnapshot[]) => void
  readonly resolveTemplateById?: (templateId: number, source: string, row: number, col: number) => FormulaTemplateResolution | undefined
  readonly initializeCellFormulasAt: (refs: readonly EngineCellMutationRef[], potentialNewCells?: number) => void
  readonly initializeFormulaSourcesAt?: (refs: EngineFormulaSourceRefs, potentialNewCells?: number) => void
  readonly resolveTemplateForCell?: (source: string, row: number, col: number) => FormulaTemplateResolution
  readonly initializePreparedCellFormulasAt?: (refs: readonly PreparedRuntimeFormulaRef[], potentialNewCells?: number) => void
  readonly initializeHydratedPreparedCellFormulasAt?: (
    refs: InitialFormulaEntryRefSource<HydratedPreparedRuntimeFormulaRef>,
    potentialNewCells?: number,
  ) => void
  readonly initializeCachedFormulaSourcesAt?: (refs: readonly CachedRuntimeFormulaRef[], potentialNewCells?: number) => void
}

export interface WorkbookRestoreResult {
  readonly formulaCount: number
}

export interface WorkbookSnapshotRestoreArgs {
  readonly snapshot: WorkbookSnapshot
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly resetWorkbook: (workbookName?: string) => void
  readonly checkEvaluationBudget?: (stepCost?: number) => void
  readonly initializeCellFormulasAt: (refs: readonly EngineCellMutationRef[], potentialNewCells?: number) => void
  readonly initializeFormulaSourcesAt?: (refs: EngineFormulaSourceRefs, potentialNewCells?: number) => void
  readonly resolveTemplateForCell?: (source: string, row: number, col: number) => FormulaTemplateResolution
  readonly initializeHydratedPreparedCellFormulasAt?: (
    refs: InitialFormulaEntryRefSource<HydratedPreparedRuntimeFormulaRef>,
    potentialNewCells?: number,
  ) => void
  readonly initializeCachedFormulaSourcesAt?: (refs: readonly CachedRuntimeFormulaRef[], potentialNewCells?: number) => void
}

export interface PreparedRuntimeFormulaRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId?: number
  readonly cellIndex?: number
}

export interface HydratedPreparedRuntimeFormulaRef extends PreparedRuntimeFormulaRef {
  readonly value: CellValue
  readonly preserveCachedValueOnFullRecalc?: boolean
}

export interface CachedRuntimeFormulaRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly value: CellValue
  readonly cellIndex?: number
}

const RUNTIME_IMAGE_COORD_STRIDE = 1_048_576

interface MutableHydratedPreparedRuntimeFormulaRef {
  sheetId: number
  row: number
  col: number
  source: string
  compiled: CompiledFormula
  templateId?: number
  cellIndex?: number
  value: CellValue
  preserveCachedValueOnFullRecalc?: boolean
}

function toFormulaInstanceKey(row: number, col: number): number {
  return row * RUNTIME_IMAGE_COORD_STRIDE + col
}

function getOrCreateSheetFormulaMap<T>(maps: Map<string, Map<number, T>>, sheetName: string): Map<number, T> {
  let sheetMap = maps.get(sheetName)
  if (!sheetMap) {
    sheetMap = new Map()
    maps.set(sheetName, sheetMap)
  }
  return sheetMap
}

function formulaValueMatchesInstance(
  record: FormulaInstanceSnapshot,
  value: RuntimeImageFormulaValueSnapshot | undefined,
): value is RuntimeImageFormulaValueSnapshot {
  return value !== undefined && value.sheetName === record.sheetName && value.row === record.row && value.col === record.col
}

function formulaValuesAreAligned(
  instances: readonly FormulaInstanceSnapshot[],
  values: readonly RuntimeImageFormulaValueSnapshot[],
): boolean {
  if (instances.length !== values.length) {
    return false
  }
  for (let index = 0; index < instances.length; index += 1) {
    if (!formulaValueMatchesInstance(instances[index]!, values[index])) {
      return false
    }
  }
  return true
}

function compareFormulaInstanceToRowCol(record: FormulaInstanceSnapshot, row: number, col: number): number {
  return record.row - row || record.col - col
}

interface RuntimeFormulaSheetSpan {
  readonly start: number
  readonly end: number
}

function buildRuntimeFormulaSheetSpans(records: readonly FormulaInstanceSnapshot[]): Map<string, RuntimeFormulaSheetSpan> {
  const spans = new Map<string, RuntimeFormulaSheetSpan>()
  let index = 0
  while (index < records.length) {
    const first = records[index]!
    const sheetName = first.sheetName
    const start = index
    index += 1
    while (index < records.length && records[index]!.sheetName === sheetName) {
      index += 1
    }
    spans.set(sheetName, { start, end: index })
  }
  return spans
}

function hasSnapshotCoordinate(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0
}

function readRestoredCellCoordinates(sheetName: string, cell: WorkbookSnapshotCell): RuntimeImageCellCoordinateSnapshot {
  if (hasSnapshotCoordinate(cell.row) && hasSnapshotCoordinate(cell.col)) {
    return {
      row: cell.row,
      col: cell.col,
    }
  }
  const parsed = parseCellAddress(cell.address, sheetName)
  return {
    row: parsed.row,
    col: parsed.col,
  }
}

interface FreshRuntimeLogicalSheetInternals {
  readonly deferVisibleCellPageRebuild?: () => void
  readonly setFreshVisibleCellIdentityWithAxisIdsDeferred?: (cellIndex: number, rowId: string, colId: string) => void
  readonly setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred?: (
    firstCellIndex: number,
    rowIds: readonly string[],
    colIds: readonly string[],
  ) => void
  readonly setFreshVisibleCellWithAxisIdsDeferred?: (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void
}

type FreshRuntimeCellAttacher = (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void

interface DenseRuntimeSheetRestorePlan {
  readonly width: number
  readonly height: number
}

class RestoredFormulaSourceRefTable implements EngineFormulaSourceRefTable {
  readonly sheetIds: Uint32Array
  readonly cellIndices: Uint32Array
  readonly rows: Uint32Array
  readonly cols: Uint32Array
  readonly sources: string[]
  readonly reusable: EngineFormulaSourceRef = {
    sheetId: 0,
    cellIndex: 0,
    row: 0,
    col: 0,
    source: '',
  }
  length = 0

  constructor(capacity: number) {
    this.sheetIds = new Uint32Array(capacity)
    this.cellIndices = new Uint32Array(capacity)
    this.rows = new Uint32Array(capacity)
    this.cols = new Uint32Array(capacity)
    this.sources = []
  }

  push(sheetId: number, cellIndex: number, row: number, col: number, source: string): void {
    const index = this.length
    this.sheetIds[index] = sheetId
    this.cellIndices[index] = cellIndex
    this.rows[index] = row
    this.cols[index] = col
    this.sources[index] = source
    this.length = index + 1
  }

  at(index: number): EngineFormulaSourceRef {
    this.reusable.sheetId = this.sheetIds[index]!
    this.reusable.cellIndex = this.cellIndices[index]!
    this.reusable.row = this.rows[index]!
    this.reusable.col = this.cols[index]!
    this.reusable.source = this.sources[index]!
    return this.reusable
  }
}

class RestoredHydratedPreparedFormulaRefTable implements Iterable<HydratedPreparedRuntimeFormulaRef> {
  readonly sheetIds: Uint32Array
  readonly cellIndices: Uint32Array
  readonly rows: Uint32Array
  readonly cols: Uint32Array
  readonly templateIds: Int32Array
  readonly sources: string[]
  readonly compiled: CompiledFormula[]
  readonly values: CellValue[]
  readonly preserveCachedValueOnFullRecalc: Uint8Array
  freshFormulaInstances: readonly FormulaInstanceSnapshot[] | undefined
  freshFormulaFamilyRuns: readonly DeferredInitialFormulaFamilyRun[] | undefined
  freshFormulaFamilyRunFallbackCount = 0
  private reusable: MutableHydratedPreparedRuntimeFormulaRef | undefined
  length = 0

  constructor(capacity: number, freshFormulaInstances?: readonly FormulaInstanceSnapshot[]) {
    this.sheetIds = new Uint32Array(capacity)
    this.cellIndices = new Uint32Array(capacity)
    this.rows = new Uint32Array(capacity)
    this.cols = new Uint32Array(capacity)
    this.templateIds = new Int32Array(capacity)
    this.templateIds.fill(-1)
    this.sources = []
    this.compiled = []
    this.values = []
    this.preserveCachedValueOnFullRecalc = new Uint8Array(capacity)
    this.freshFormulaInstances = freshFormulaInstances
  }

  push(
    sheetId: number,
    cellIndex: number,
    row: number,
    col: number,
    source: string,
    compiled: CompiledFormula,
    templateId: number | undefined,
    value: CellValue,
    runtimeImageCellIndex: number,
    preserveCachedValueOnFullRecalc = false,
  ): void {
    const index = this.length
    if (this.freshFormulaInstances !== undefined && runtimeImageCellIndex !== cellIndex) {
      this.freshFormulaInstances = undefined
    }
    this.sheetIds[index] = sheetId
    this.cellIndices[index] = cellIndex
    this.rows[index] = row
    this.cols[index] = col
    this.templateIds[index] = templateId ?? -1
    this.sources[index] = source
    this.compiled[index] = compiled
    this.values[index] = value
    this.preserveCachedValueOnFullRecalc[index] = preserveCachedValueOnFullRecalc ? 1 : 0
    this.length = index + 1
  }

  at(index: number): HydratedPreparedRuntimeFormulaRef {
    const reusable =
      this.reusable ??
      (this.reusable = {
        sheetId: 0,
        row: 0,
        col: 0,
        source: '',
        compiled: this.compiled[index]!,
        value: this.values[index]!,
      })
    reusable.sheetId = this.sheetIds[index]!
    reusable.cellIndex = this.cellIndices[index]!
    reusable.row = this.rows[index]!
    reusable.col = this.cols[index]!
    reusable.source = this.sources[index]!
    reusable.compiled = this.compiled[index]!
    const templateId = this.templateIds[index]!
    if (templateId === -1) {
      delete reusable.templateId
    } else {
      reusable.templateId = templateId
    }
    reusable.value = this.values[index]!
    if (this.preserveCachedValueOnFullRecalc[index] === 1) {
      reusable.preserveCachedValueOnFullRecalc = true
    } else {
      delete reusable.preserveCachedValueOnFullRecalc
    }
    return reusable
  }

  *[Symbol.iterator](): IterableIterator<HydratedPreparedRuntimeFormulaRef> {
    for (let index = 0; index < this.length; index += 1) {
      const templateId = this.templateIds[index]!
      yield {
        sheetId: this.sheetIds[index]!,
        cellIndex: this.cellIndices[index]!,
        row: this.rows[index]!,
        col: this.cols[index]!,
        source: this.sources[index]!,
        compiled: this.compiled[index]!,
        ...(templateId === -1 ? {} : { templateId }),
        value: this.values[index]!,
        ...(this.preserveCachedValueOnFullRecalc[index] === 1 ? { preserveCachedValueOnFullRecalc: true } : {}),
      }
    }
  }
}

function isFreshRuntimeLogicalSheetInternals(value: unknown): value is FreshRuntimeLogicalSheetInternals {
  return typeof value === 'object' && value !== null
}

function createFreshRuntimeCellAttacher(workbook: WorkbookStore, sheet: SheetRecord): FreshRuntimeCellAttacher {
  const logicalCandidate: unknown = sheet.logical
  const logical = isFreshRuntimeLogicalSheetInternals(logicalCandidate) ? logicalCandidate : undefined
  const attachFreshVisibleCellIdentity = logical?.setFreshVisibleCellIdentityWithAxisIdsDeferred?.bind(logical)
  if (attachFreshVisibleCellIdentity) {
    logical?.deferVisibleCellPageRebuild?.()
    const setGridCell = sheet.grid.createRowMajorSetter()
    return (row, col, cellIndex, rowId, colId) => {
      attachFreshVisibleCellIdentity(cellIndex, rowId, colId)
      setGridCell(row, col, cellIndex)
    }
  }
  const attachFreshVisibleCell = logical?.setFreshVisibleCellWithAxisIdsDeferred?.bind(logical)
  if (!attachFreshVisibleCell) {
    return (row, col, cellIndex, rowId, colId) => {
      workbook.attachAllocatedCellWithLogicalAxisIds(sheet.id, row, col, cellIndex, rowId, colId)
    }
  }

  const setGridCell = sheet.grid.createRowMajorSetter()

  return (row, col, cellIndex, rowId, colId) => {
    attachFreshVisibleCell(row, col, cellIndex, rowId, colId)
    setGridCell(row, col, cellIndex)
  }
}

function attachDenseFreshRuntimeCells(
  sheet: SheetRecord,
  firstCellIndex: number,
  rowStart: number,
  colStart: number,
  rowIds: readonly string[],
  colIds: readonly string[],
): boolean {
  const logicalCandidate: unknown = sheet.logical
  const logical = isFreshRuntimeLogicalSheetInternals(logicalCandidate) ? logicalCandidate : undefined
  const attachDenseFreshVisibleCellIdentities = logical?.setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred?.bind(logical)
  if (!attachDenseFreshVisibleCellIdentities) {
    return false
  }
  logical?.deferVisibleCellPageRebuild?.()
  attachDenseFreshVisibleCellIdentities(firstCellIndex, rowIds, colIds)
  sheet.grid.setDenseRowMajor(rowStart, colStart, rowIds.length, colIds.length, firstCellIndex)
  return true
}

function getDenseRuntimeSheetRestorePlan(
  sheet: WorkbookSnapshot['sheets'][number],
  sheetCells: RuntimeImageSheetCellsSnapshot | undefined,
): DenseRuntimeSheetRestorePlan | undefined {
  const dimensions = sheetCells?.dimensions
  if (!dimensions) {
    return undefined
  }
  const { width, height } = dimensions
  const cellCount = sheetCells.cellCount ?? sheetCells.coords.length
  const hasDenseCoordinateOrder = sheetCells.coordinateOrder === 'dense-row-major'
  if (
    !Number.isInteger(width) ||
    !Number.isInteger(height) ||
    width <= 0 ||
    height <= 0 ||
    cellCount !== sheet.cells.length ||
    width * height !== sheet.cells.length
  ) {
    return undefined
  }
  if (hasDenseCoordinateOrder) {
    return { width, height }
  }
  if (sheetCells.coords.length !== sheet.cells.length) {
    return undefined
  }
  for (let index = 0; index < sheetCells.coords.length; index += 1) {
    const coords = sheetCells.coords[index]!
    if (coords.row !== Math.floor(index / width) || coords.col !== index % width) {
      return undefined
    }
  }
  return { width, height }
}

export function restoreWorkbookFromSnapshot(args: WorkbookSnapshotRestoreArgs): WorkbookRestoreResult {
  const orderedSheets = restoreWorkbookStructure(args)
  args.checkEvaluationBudget?.()
  const potentialNewCells = orderedSheets.reduce((count, sheet) => count + sheet.cells.length, 0)
  const formulaRefs: EngineCellMutationRef[] = []
  const formulaSourceRefs = args.initializeFormulaSourcesAt ? new RestoredFormulaSourceRefTable(potentialNewCells) : undefined
  const hydratedPreparedFormulaRefs: HydratedPreparedRuntimeFormulaRef[] = []
  const cachedFormulaRefs: CachedRuntimeFormulaRef[] = []
  const canHydratePreparedCachedFormulaValues = args.initializeHydratedPreparedCellFormulasAt && args.resolveTemplateForCell
  const canHydrateImportedCachedFormulaValues = args.initializeCachedFormulaSourcesAt !== undefined
  const shouldHydrateIterativeFormulaValues = args.snapshot.workbook.metadata?.calculationSettings?.iterate === true
  const shouldHydrateImportedCachedFormulaValues =
    args.snapshot.workbook.metadata?.calculationSettings?.fullCalcOnLoad === false ||
    args.snapshot.workbook.metadata?.calculationSettings?.mode === 'manual'
  const definedFormulaNames =
    shouldHydrateIterativeFormulaValues || !canHydratePreparedCachedFormulaValues ? undefined : collectDefinedFormulaNames(args.snapshot)
  const restoredStringIds = new Map<string, number>()

  args.checkEvaluationBudget?.()
  args.workbook.cellStore.ensureCapacity(args.workbook.cellStore.size + potentialNewCells)
  const previousOnSetValue = args.workbook.cellStore.onSetValue
  args.workbook.cellStore.onSetValue = null
  args.workbook.withBatchedColumnVersionUpdates(() => {
    try {
      for (let sheetIndex = 0; sheetIndex < orderedSheets.length; sheetIndex += 1) {
        args.checkEvaluationBudget?.()
        const sheet = orderedSheets[sheetIndex]!
        const sheetRecord = args.workbook.getSheet(sheet.name)
        if (!sheetRecord) {
          throw new Error(`Missing restore sheet: ${sheet.name}`)
        }
        const sheetId = sheetRecord.id
        const rowIds: string[] = []
        const colIds: string[] = []
        const ensureRowId = args.workbook.createLogicalAxisIdEnsurer(sheetId, 'row')
        const ensureColId = args.workbook.createLogicalAxisIdEnsurer(sheetId, 'column')
        const attachFreshCell = createFreshRuntimeCellAttacher(args.workbook, sheetRecord)
        let literalColumns: WrittenColumnTracker | undefined
        for (let cellIndex = 0; cellIndex < sheet.cells.length; cellIndex += 1) {
          args.checkEvaluationBudget?.()
          const cell = sheet.cells[cellIndex]!
          const coords = readRestoredCellCoordinates(sheet.name, cell)
          const restoredCellIndex = args.workbook.cellStore.allocateReserved(sheetId, coords.row, coords.col)
          const rowId = (rowIds[coords.row] ??= ensureRowId(coords.row))
          const colId = (colIds[coords.col] ??= ensureColId(coords.col))
          attachFreshCell(coords.row, coords.col, restoredCellIndex, rowId, colId)
          if (cell.formula !== undefined) {
            let hydratedCachedFormula = false
            const shouldPreserveCachedUnsupportedValue =
              canHydratePreparedCachedFormulaValues &&
              cell.value !== undefined &&
              !shouldHydrateIterativeFormulaValues &&
              formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc(cell.formula, definedFormulaNames!)
            if (canHydrateImportedCachedFormulaValues && shouldHydrateImportedCachedFormulaValues && cell.value !== undefined) {
              cachedFormulaRefs.push({
                sheetId,
                row: coords.row,
                col: coords.col,
                cellIndex: restoredCellIndex,
                source: cell.formula,
                value: formulaCachedLiteralToRestoredValue(cell.value, args.strings, restoredStringIds),
              })
              hydratedCachedFormula = true
            } else if (
              canHydratePreparedCachedFormulaValues &&
              cell.value !== undefined &&
              (shouldHydrateIterativeFormulaValues || shouldPreserveCachedUnsupportedValue)
            ) {
              try {
                const template = args.resolveTemplateForCell(cell.formula, coords.row, coords.col)
                if (!template.compiled.volatile && !template.compiled.producesSpill) {
                  hydratedPreparedFormulaRefs.push({
                    sheetId,
                    row: coords.row,
                    col: coords.col,
                    cellIndex: restoredCellIndex,
                    source: cell.formula,
                    compiled: template.compiled,
                    templateId: template.templateId,
                    value: formulaCachedLiteralToRestoredValue(cell.value, args.strings, restoredStringIds),
                    ...(shouldPreserveCachedUnsupportedValue ? { preserveCachedValueOnFullRecalc: true } : {}),
                  })
                  hydratedCachedFormula = true
                }
              } catch {
                hydratedCachedFormula = false
              }
            }
            if (!hydratedCachedFormula) {
              if (formulaSourceRefs) {
                formulaSourceRefs.push(sheetId, restoredCellIndex, coords.row, coords.col, cell.formula)
              } else {
                formulaRefs.push({
                  sheetId,
                  cellIndex: restoredCellIndex,
                  mutation: {
                    kind: 'setCellFormula',
                    row: coords.row,
                    col: coords.col,
                    formula: cell.formula,
                  },
                })
              }
            }
          } else {
            restoreLiteralCell(args.workbook, args.strings, restoredCellIndex, cell.value ?? null, restoredStringIds)
            literalColumns ??= createWrittenColumnTracker()
            markWrittenColumn(literalColumns, coords.col)
          }
          if (cell.format !== undefined) {
            args.workbook.setCellFormat(restoredCellIndex, cell.format)
          }
        }
        if (literalColumns && literalColumns.count > 0) {
          args.workbook.notifyColumnsWritten(sheetId, materializeWrittenColumns(literalColumns))
        }
      }
    } finally {
      args.workbook.cellStore.onSetValue = previousOnSetValue
    }
  })

  if (hydratedPreparedFormulaRefs.length > 0 && args.initializeHydratedPreparedCellFormulasAt) {
    args.checkEvaluationBudget?.()
    args.initializeHydratedPreparedCellFormulasAt(hydratedPreparedFormulaRefs, hydratedPreparedFormulaRefs.length)
  }
  if (cachedFormulaRefs.length > 0 && args.initializeCachedFormulaSourcesAt) {
    args.checkEvaluationBudget?.()
    args.initializeCachedFormulaSourcesAt(cachedFormulaRefs, cachedFormulaRefs.length)
  }
  if (formulaSourceRefs && formulaSourceRefs.length > 0) {
    args.checkEvaluationBudget?.()
    args.initializeFormulaSourcesAt!(formulaSourceRefs, potentialNewCells)
  } else if (formulaRefs.length > 0) {
    args.checkEvaluationBudget?.()
    args.initializeCellFormulasAt(formulaRefs, potentialNewCells)
  }

  args.checkEvaluationBudget?.()
  restoreVisualMetadata({
    workbook: args.workbook,
    workbookMetadata: args.snapshot.workbook.metadata,
  })
  return { formulaCount: hydratedPreparedFormulaRefs.length + (formulaSourceRefs?.length ?? formulaRefs.length) }
}

export function restoreWorkbookFromRuntimeImage(args: RuntimeImageRestoreArgs): WorkbookRestoreResult {
  const orderedSheets = restoreWorkbookStructure(args)
  const sheetIdsByName = new Map<string, number>()

  args.checkEvaluationBudget?.()
  args.hydrateTemplateBank(args.runtimeImage.templateBank)

  args.checkEvaluationBudget?.()
  const formulaValueIndexAligned = formulaValuesAreAligned(args.runtimeImage.formulaInstances, args.runtimeImage.formulaValues)
  const formulaValuesByAddress = formulaValueIndexAligned ? undefined : new Map<string, Map<number, CellValue>>()
  if (formulaValuesByAddress) {
    args.runtimeImage.formulaValues.forEach((record) => {
      args.checkEvaluationBudget?.()
      getOrCreateSheetFormulaMap(formulaValuesByAddress, record.sheetName).set(toFormulaInstanceKey(record.row, record.col), record.value)
    })
  }
  args.checkEvaluationBudget?.()
  const formulaSpansBySheet = buildRuntimeFormulaSheetSpans(args.runtimeImage.formulaInstances)
  const sheetCellsByName = new Map<string, RuntimeImageSheetCellsSnapshot>(
    (args.runtimeImage.sheetCells ?? []).map((record) => [record.sheetName, record]),
  )
  const totalCellCount = orderedSheets.reduce((sum, sheet) => sum + sheet.cells.length, 0)
  if (totalCellCount > 0) {
    args.checkEvaluationBudget?.()
    args.workbook.cellStore.ensureCapacity(args.workbook.cellStore.size + totalCellCount)
  }

  const formulaRefs: EngineCellMutationRef[] = []
  const formulaSourceRefs = args.initializeFormulaSourcesAt ? new RestoredFormulaSourceRefTable(totalCellCount) : undefined
  const preparedFormulaRefs: PreparedRuntimeFormulaRef[] = []
  const hydratedPreparedFormulaRefs = new RestoredHydratedPreparedFormulaRefTable(totalCellCount, args.runtimeImage.formulaInstances)
  const cachedFormulaRefs: CachedRuntimeFormulaRef[] = []
  const canHydratePreparedSnapshotFormulaValues = args.initializeHydratedPreparedCellFormulasAt && args.resolveTemplateForCell
  const canHydrateImportedCachedSnapshotFormulaValues = args.initializeCachedFormulaSourcesAt !== undefined
  const shouldHydrateIterativeFormulaValues = args.snapshot.workbook.metadata?.calculationSettings?.iterate === true
  const shouldHydrateImportedCachedFormulaValues =
    args.snapshot.workbook.metadata?.calculationSettings?.fullCalcOnLoad === false ||
    args.snapshot.workbook.metadata?.calculationSettings?.mode === 'manual'
  const definedFormulaNames = shouldHydrateIterativeFormulaValues ? undefined : collectDefinedFormulaNames(args.snapshot)
  const restoredStringIds = new Map<string, number>()
  const previousOnSetValue = args.workbook.cellStore.onSetValue
  args.workbook.cellStore.onSetValue = null
  args.workbook.withBatchedColumnVersionUpdates(() => {
    try {
      orderedSheets.forEach((sheet) => {
        args.checkEvaluationBudget?.()
        const sheetRecord = args.workbook.getSheet(sheet.name)
        if (!sheetRecord) {
          throw new Error(`Missing runtime restore sheet: ${sheet.name}`)
        }
        const sheetId = sheetRecord.id
        sheetIdsByName.set(sheet.name, sheetId)
        const sheetCellSnapshot = sheetCellsByName.get(sheet.name)
        const sheetCoords = sheetCellSnapshot?.coords
        const rowIds: string[] = []
        const colIds: string[] = []
        const ensureRowId = args.workbook.createLogicalAxisIdEnsurer(sheetId, 'row')
        const ensureColId = args.workbook.createLogicalAxisIdEnsurer(sheetId, 'column')
        let attachFreshCell: FreshRuntimeCellAttacher | undefined
        const getFreshCellAttacher = (): FreshRuntimeCellAttacher => {
          attachFreshCell ??= createFreshRuntimeCellAttacher(args.workbook, sheetRecord)
          return attachFreshCell
        }
        const formulaSpan = formulaSpansBySheet.get(sheet.name)
        let formulaInstanceIndex = formulaSpan?.start ?? 0
        const formulaInstanceEnd = formulaSpan?.end ?? formulaInstanceIndex
        let literalColumns: WrittenColumnTracker | undefined

        const restoreRuntimeCell = (cell: WorkbookSnapshotCell, row: number, col: number, cellIndex: number): void => {
          while (
            formulaInstanceIndex < formulaInstanceEnd &&
            compareFormulaInstanceToRowCol(args.runtimeImage.formulaInstances[formulaInstanceIndex]!, row, col) < 0
          ) {
            args.checkEvaluationBudget?.()
            formulaInstanceIndex += 1
          }
          const candidateFormula =
            formulaInstanceIndex < formulaInstanceEnd ? args.runtimeImage.formulaInstances[formulaInstanceIndex] : undefined
          const restoredFormula =
            candidateFormula && compareFormulaInstanceToRowCol(candidateFormula, row, col) === 0 ? candidateFormula : undefined
          if (cell.formula === undefined && restoredFormula === undefined) {
            if (cell.value === undefined) {
              writeLiteralToCellStore(args.workbook.cellStore, cellIndex, null, args.strings)
            } else {
              restoreLiteralCell(args.workbook, args.strings, cellIndex, cell.value, restoredStringIds)
            }
            literalColumns ??= createWrittenColumnTracker()
            markWrittenColumn(literalColumns, col)
            if (cell.value === null) {
              args.workbook.cellStore.flags[cellIndex] = (args.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.AuthoredBlank
            }
          }
          if (cell.format !== undefined) {
            args.workbook.setCellFormat(cellIndex, cell.format)
          }
          if (restoredFormula) {
            const cachedValue = formulaValueIndexAligned
              ? args.runtimeImage.formulaValues[formulaInstanceIndex]?.value
              : formulaValuesByAddress?.get(sheet.name)?.get(toFormulaInstanceKey(row, col))
            const shouldPreserveCachedUnsupportedValue =
              cachedValue !== undefined &&
              !shouldHydrateIterativeFormulaValues &&
              definedFormulaNames !== undefined &&
              formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc(restoredFormula.source, definedFormulaNames)
            const template =
              restoredFormula.templateId !== undefined && args.resolveTemplateById
                ? args.resolveTemplateById(restoredFormula.templateId, restoredFormula.source, row, col)
                : undefined
            if (args.initializeHydratedPreparedCellFormulasAt && cachedValue !== undefined) {
              if (template && !template.compiled.volatile && !template.compiled.producesSpill) {
                hydratedPreparedFormulaRefs.push(
                  sheetId,
                  cellIndex,
                  row,
                  col,
                  restoredFormula.source,
                  template.compiled,
                  template.templateId,
                  cachedValue,
                  restoredFormula.cellIndex,
                  shouldPreserveCachedUnsupportedValue,
                )
                return
              }
            }
            if (template && args.initializePreparedCellFormulasAt) {
              preparedFormulaRefs.push({
                sheetId,
                row,
                col,
                cellIndex,
                source: restoredFormula.source,
                compiled: template.compiled,
                templateId: template.templateId,
              })
              return
            }
            if (formulaSourceRefs) {
              formulaSourceRefs.push(sheetId, cellIndex, row, col, restoredFormula.source)
            } else {
              formulaRefs.push({
                sheetId,
                cellIndex,
                mutation: {
                  kind: 'setCellFormula',
                  row,
                  col,
                  formula: restoredFormula.source,
                },
              })
            }
          } else if (cell.formula !== undefined) {
            let hydratedCachedFormula = false
            const shouldPreserveCachedUnsupportedValue =
              canHydratePreparedSnapshotFormulaValues &&
              cell.value !== undefined &&
              !shouldHydrateIterativeFormulaValues &&
              definedFormulaNames !== undefined &&
              formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc(cell.formula, definedFormulaNames)
            if (canHydrateImportedCachedSnapshotFormulaValues && shouldHydrateImportedCachedFormulaValues && cell.value !== undefined) {
              cachedFormulaRefs.push({
                sheetId,
                row,
                col,
                cellIndex,
                source: cell.formula,
                value: formulaCachedLiteralToRestoredValue(cell.value, args.strings, restoredStringIds),
              })
              hydratedCachedFormula = true
            } else if (
              canHydratePreparedSnapshotFormulaValues &&
              cell.value !== undefined &&
              (shouldHydrateIterativeFormulaValues || shouldPreserveCachedUnsupportedValue)
            ) {
              try {
                const template = args.resolveTemplateForCell(cell.formula, row, col)
                if (!template.compiled.volatile && !template.compiled.producesSpill) {
                  hydratedPreparedFormulaRefs.push(
                    sheetId,
                    cellIndex,
                    row,
                    col,
                    cell.formula,
                    template.compiled,
                    template.templateId,
                    formulaCachedLiteralToRestoredValue(cell.value, args.strings, restoredStringIds),
                    cellIndex,
                    shouldPreserveCachedUnsupportedValue,
                  )
                  hydratedCachedFormula = true
                }
              } catch {
                hydratedCachedFormula = false
              }
            }
            if (!hydratedCachedFormula) {
              if (formulaSourceRefs) {
                formulaSourceRefs.push(sheetId, cellIndex, row, col, cell.formula)
              } else {
                formulaRefs.push({
                  sheetId,
                  cellIndex,
                  mutation: {
                    kind: 'setCellFormula',
                    row,
                    col,
                    formula: cell.formula,
                  },
                })
              }
            }
          }
        }

        const denseRestorePlan = getDenseRuntimeSheetRestorePlan(sheet, sheetCellSnapshot)
        if (denseRestorePlan) {
          const firstCellIndex = args.workbook.cellStore.allocateDenseRowMajorAtReserved(
            sheetId,
            0,
            denseRestorePlan.height,
            0,
            denseRestorePlan.width,
          )
          for (let col = 0; col < denseRestorePlan.width; col += 1) {
            colIds[col] = ensureColId(col)
          }
          for (let row = 0; row < denseRestorePlan.height; row += 1) {
            rowIds[row] = ensureRowId(row)
          }
          const attachedDenseCells = attachDenseFreshRuntimeCells(sheetRecord, firstCellIndex, 0, 0, rowIds, colIds)
          let index = 0
          if (attachedDenseCells) {
            for (let row = 0; row < denseRestorePlan.height; row += 1) {
              args.checkEvaluationBudget?.()
              for (let col = 0; col < denseRestorePlan.width; col += 1) {
                restoreRuntimeCell(sheet.cells[index]!, row, col, firstCellIndex + index)
                index += 1
              }
            }
          } else {
            const attachCell = getFreshCellAttacher()
            for (let row = 0; row < denseRestorePlan.height; row += 1) {
              args.checkEvaluationBudget?.()
              const rowId = rowIds[row]!
              for (let col = 0; col < denseRestorePlan.width; col += 1) {
                const cellIndex = firstCellIndex + index
                attachCell(row, col, cellIndex, rowId, colIds[col]!)
                restoreRuntimeCell(sheet.cells[index]!, row, col, cellIndex)
                index += 1
              }
            }
          }
        } else {
          const attachCell = getFreshCellAttacher()
          for (let index = 0; index < sheet.cells.length; index += 1) {
            args.checkEvaluationBudget?.()
            const cell = sheet.cells[index]!
            const coords = sheetCoords?.[index] ?? readRestoredCellCoordinates(sheet.name, cell)
            const cellIndex = args.workbook.cellStore.allocateReserved(sheetId, coords.row, coords.col)
            const rowId = (rowIds[coords.row] ??= ensureRowId(coords.row))
            const colId = (colIds[coords.col] ??= ensureColId(coords.col))
            attachCell(coords.row, coords.col, cellIndex, rowId, colId)
            restoreRuntimeCell(cell, coords.row, coords.col, cellIndex)
          }
        }
        if (literalColumns && literalColumns.count > 0) {
          args.workbook.notifyColumnsWritten(sheetId, materializeWrittenColumns(literalColumns))
        }
      })
    } finally {
      args.workbook.cellStore.onSetValue = previousOnSetValue
    }
  })

  if (hydratedPreparedFormulaRefs.length > 0 && args.initializeHydratedPreparedCellFormulasAt) {
    args.checkEvaluationBudget?.()
    const restoredFormulaFamilyRuns =
      hydratedPreparedFormulaRefs.freshFormulaInstances === undefined
        ? undefined
        : restoreAlignedRuntimeFormulaFamilyRuns({
            runs: args.runtimeImage.formulaFamilyRuns,
            sheetIdsByName,
          })
    if (restoredFormulaFamilyRuns !== undefined) {
      if (restoredFormulaFamilyRuns.fallbackCount === 0 && restoredFormulaFamilyRuns.runs.length > 0) {
        hydratedPreparedFormulaRefs.freshFormulaFamilyRuns = restoredFormulaFamilyRuns.runs
      }
      hydratedPreparedFormulaRefs.freshFormulaFamilyRunFallbackCount = restoredFormulaFamilyRuns.fallbackCount
    } else if ((args.runtimeImage.formulaFamilyRuns?.length ?? 0) > 0) {
      hydratedPreparedFormulaRefs.freshFormulaFamilyRunFallbackCount = args.runtimeImage.formulaFamilyRuns!.length
    }
    args.initializeHydratedPreparedCellFormulasAt(hydratedPreparedFormulaRefs, hydratedPreparedFormulaRefs.length)
  }
  if (cachedFormulaRefs.length > 0 && args.initializeCachedFormulaSourcesAt) {
    args.checkEvaluationBudget?.()
    args.initializeCachedFormulaSourcesAt(cachedFormulaRefs, cachedFormulaRefs.length)
  }
  if (preparedFormulaRefs.length > 0 && args.initializePreparedCellFormulasAt) {
    args.checkEvaluationBudget?.()
    args.initializePreparedCellFormulasAt(preparedFormulaRefs, preparedFormulaRefs.length)
  }
  if (formulaSourceRefs && formulaSourceRefs.length > 0) {
    args.checkEvaluationBudget?.()
    args.initializeFormulaSourcesAt!(formulaSourceRefs, totalCellCount)
  }
  if (formulaRefs.length > 0) {
    args.checkEvaluationBudget?.()
    args.initializeCellFormulasAt(formulaRefs, formulaRefs.length)
  }

  args.checkEvaluationBudget?.()
  restoreVisualMetadata({
    workbook: args.workbook,
    workbookMetadata: args.snapshot.workbook.metadata,
  })
  return {
    formulaCount: hydratedPreparedFormulaRefs.length + preparedFormulaRefs.length + (formulaSourceRefs?.length ?? 0) + formulaRefs.length,
  }
}
