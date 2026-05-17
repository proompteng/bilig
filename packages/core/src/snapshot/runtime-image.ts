import { parseCellAddress, type CompiledFormula } from '@bilig/formula'
import type { CellValue, LiteralInput, WorkbookSnapshot } from '@bilig/protocol'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type {
  EngineCellMutationRef,
  EngineFormulaSourceRef,
  EngineFormulaSourceRefs,
  EngineFormulaSourceRefTable,
} from '../cell-mutations-at.js'
import { CellFlags } from '../cell-store.js'
import { literalToValue, writeLiteralToCellStore } from '../engine-value-utils.js'
import type { FormulaInstanceSnapshot } from '../formula/formula-instance-table.js'
import type { FormulaTemplateResolution, FormulaTemplateSnapshot } from '../formula/template-bank.js'
import { collectDefinedFormulaNames, formulaShouldUseCachedUnsupportedFunctionValue } from './unsupported-formula-cache.js'
import type { StringPool } from '../string-pool.js'
import type { SheetRecord, WorkbookStore } from '../workbook-store.js'
import { restoreVisualMetadata, restoreWorkbookStructure } from './runtime-image-metadata-restore.js'

type WorkbookSnapshotCell = WorkbookSnapshot['sheets'][number]['cells'][number]

export interface RuntimeImage {
  readonly version: 1
  readonly templateBank: readonly FormulaTemplateSnapshot[]
  readonly formulaInstances: readonly FormulaInstanceSnapshot[]
  readonly formulaValues: readonly RuntimeImageFormulaValueSnapshot[]
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
  readonly initializePreparedCellFormulasAt?: (refs: readonly PreparedRuntimeFormulaRef[], potentialNewCells?: number) => void
  readonly initializeHydratedPreparedCellFormulasAt?: (
    refs: readonly HydratedPreparedRuntimeFormulaRef[],
    potentialNewCells?: number,
  ) => void
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
    refs: readonly HydratedPreparedRuntimeFormulaRef[],
    potentialNewCells?: number,
  ) => void
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
}

const RUNTIME_IMAGE_COORD_STRIDE = 1_048_576

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

function compareFormulaInstanceToCoordinate(record: FormulaInstanceSnapshot, coords: RuntimeImageCellCoordinateSnapshot): number {
  return record.row - coords.row || record.col - coords.col
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

interface WrittenColumnTracker {
  smallColumns: number
  columns?: Uint8Array
  count: number
}

interface FreshRuntimeLogicalSheetInternals {
  readonly setFreshVisibleCellWithAxisIdsDeferred?: (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void
}

type FreshRuntimeCellAttacher = (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void

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

function isFreshRuntimeLogicalSheetInternals(value: unknown): value is FreshRuntimeLogicalSheetInternals {
  return typeof value === 'object' && value !== null
}

function createWrittenColumnTracker(): WrittenColumnTracker {
  return {
    smallColumns: 0,
    count: 0,
  }
}

function markWrittenColumn(tracker: WrittenColumnTracker, col: number): void {
  if (col < 30) {
    const bit = 1 << col
    if ((tracker.smallColumns & bit) !== 0) {
      return
    }
    tracker.smallColumns |= bit
    tracker.count += 1
    return
  }
  let columns = tracker.columns
  if (!columns) {
    columns = new Uint8Array(Math.max(32, col + 1))
    tracker.columns = columns
  } else if (col >= columns.length) {
    let nextLength = columns.length
    while (nextLength <= col) {
      nextLength *= 2
    }
    const nextColumns = new Uint8Array(nextLength)
    nextColumns.set(columns)
    columns = nextColumns
    tracker.columns = columns
  }
  if (columns[col] !== 0) {
    return
  }
  columns[col] = 1
  tracker.count += 1
}

function materializeWrittenColumns(tracker: WrittenColumnTracker): Uint32Array {
  const columns = new Uint32Array(tracker.count)
  let writeIndex = 0
  for (let col = 0; col < 30; col += 1) {
    if ((tracker.smallColumns & (1 << col)) !== 0) {
      columns[writeIndex] = col
      writeIndex += 1
    }
  }
  const largeColumns = tracker.columns
  if (largeColumns) {
    for (let col = 30; col < largeColumns.length; col += 1) {
      if (largeColumns[col] === 0) {
        continue
      }
      columns[writeIndex] = col
      writeIndex += 1
    }
  }
  return columns
}

function createFreshRuntimeCellAttacher(workbook: WorkbookStore, sheet: SheetRecord): FreshRuntimeCellAttacher {
  const logicalCandidate: unknown = sheet.logical
  const logical = isFreshRuntimeLogicalSheetInternals(logicalCandidate) ? logicalCandidate : undefined
  const attachFreshVisibleCell = logical?.setFreshVisibleCellWithAxisIdsDeferred?.bind(logical)
  if (!attachFreshVisibleCell) {
    return (row, col, cellIndex, rowId, colId) => {
      workbook.attachAllocatedCellWithLogicalAxisIds(sheet.id, row, col, cellIndex, rowId, colId)
    }
  }

  return (row, col, cellIndex, rowId, colId) => {
    attachFreshVisibleCell(row, col, cellIndex, rowId, colId)
    sheet.grid.set(row, col, cellIndex)
  }
}

function restoreLiteralCell(
  workbook: WorkbookStore,
  strings: StringPool,
  cellIndex: number,
  value: LiteralInput,
  stringIdCache?: Map<string, number>,
): void {
  const cellStore = workbook.cellStore
  const flags = cellStore.flags[cellIndex] ?? 0
  if (value === null) {
    cellStore.tags[cellIndex] = ValueTag.Empty
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = 0
    cellStore.numbers[cellIndex] = 0
    cellStore.flags[cellIndex] = flags | CellFlags.AuthoredBlank
  } else if (typeof value === 'number') {
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = 0
    cellStore.numbers[cellIndex] = value
    if ((flags & CellFlags.AuthoredBlank) !== 0) {
      cellStore.flags[cellIndex] = flags & ~CellFlags.AuthoredBlank
    }
  } else if (typeof value === 'boolean') {
    cellStore.tags[cellIndex] = ValueTag.Boolean
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = 0
    cellStore.numbers[cellIndex] = value ? 1 : 0
    if ((flags & CellFlags.AuthoredBlank) !== 0) {
      cellStore.flags[cellIndex] = flags & ~CellFlags.AuthoredBlank
    }
  } else {
    let stringId = stringIdCache?.get(value)
    if (stringId === undefined) {
      stringId = strings.intern(value)
      stringIdCache?.set(value, stringId)
    }
    cellStore.tags[cellIndex] = ValueTag.String
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = stringId
    cellStore.numbers[cellIndex] = 0
    if ((flags & CellFlags.AuthoredBlank) !== 0) {
      cellStore.flags[cellIndex] = flags & ~CellFlags.AuthoredBlank
    }
  }
  cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  cellStore.onSetValue?.(cellIndex)
}

export function restoreWorkbookFromSnapshot(args: WorkbookSnapshotRestoreArgs): WorkbookRestoreResult {
  const orderedSheets = restoreWorkbookStructure(args)
  args.checkEvaluationBudget?.()
  const potentialNewCells = orderedSheets.reduce((count, sheet) => count + sheet.cells.length, 0)
  const formulaRefs: EngineCellMutationRef[] = []
  const formulaSourceRefs = args.initializeFormulaSourcesAt ? new RestoredFormulaSourceRefTable(potentialNewCells) : undefined
  const hydratedPreparedFormulaRefs: HydratedPreparedRuntimeFormulaRef[] = []
  const canHydrateCachedFormulaValues = args.initializeHydratedPreparedCellFormulasAt && args.resolveTemplateForCell
  const shouldHydrateIterativeFormulaValues = args.snapshot.workbook.metadata?.calculationSettings?.iterate === true
  const definedFormulaNames =
    shouldHydrateIterativeFormulaValues || !canHydrateCachedFormulaValues ? undefined : collectDefinedFormulaNames(args.snapshot)
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
            if (
              canHydrateCachedFormulaValues &&
              cell.value !== undefined &&
              (shouldHydrateIterativeFormulaValues || formulaShouldUseCachedUnsupportedFunctionValue(cell.formula, definedFormulaNames!))
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
                    value: literalToValue(cell.value, args.strings),
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
  const sheetCellsByName = new Map<string, readonly RuntimeImageCellCoordinateSnapshot[]>(
    (args.runtimeImage.sheetCells ?? []).map((record) => [record.sheetName, record.coords]),
  )
  const totalCellCount = orderedSheets.reduce((sum, sheet) => sum + sheet.cells.length, 0)
  if (totalCellCount > 0) {
    args.checkEvaluationBudget?.()
    args.workbook.cellStore.ensureCapacity(args.workbook.cellStore.size + totalCellCount)
  }

  const formulaRefs: EngineCellMutationRef[] = []
  const preparedFormulaRefs: PreparedRuntimeFormulaRef[] = []
  const hydratedPreparedFormulaRefs: HydratedPreparedRuntimeFormulaRef[] = []
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
        const sheetCoords = sheetCellsByName.get(sheet.name)
        const rowIds: string[] = []
        const colIds: string[] = []
        const ensureRowId = args.workbook.createLogicalAxisIdEnsurer(sheetId, 'row')
        const ensureColId = args.workbook.createLogicalAxisIdEnsurer(sheetId, 'column')
        const attachFreshCell = createFreshRuntimeCellAttacher(args.workbook, sheetRecord)
        const formulaSpan = formulaSpansBySheet.get(sheet.name)
        let formulaInstanceIndex = formulaSpan?.start ?? 0
        const formulaInstanceEnd = formulaSpan?.end ?? formulaInstanceIndex
        let literalColumns: WrittenColumnTracker | undefined
        for (let index = 0; index < sheet.cells.length; index += 1) {
          args.checkEvaluationBudget?.()
          const cell = sheet.cells[index]!
          const coords = sheetCoords?.[index] ?? readRestoredCellCoordinates(sheet.name, cell)
          while (
            formulaInstanceIndex < formulaInstanceEnd &&
            compareFormulaInstanceToCoordinate(args.runtimeImage.formulaInstances[formulaInstanceIndex]!, coords) < 0
          ) {
            args.checkEvaluationBudget?.()
            formulaInstanceIndex += 1
          }
          const candidateFormula =
            formulaInstanceIndex < formulaInstanceEnd ? args.runtimeImage.formulaInstances[formulaInstanceIndex] : undefined
          const restoredFormula =
            candidateFormula && compareFormulaInstanceToCoordinate(candidateFormula, coords) === 0 ? candidateFormula : undefined
          const cellIndex = args.workbook.cellStore.allocateReserved(sheetId, coords.row, coords.col)
          const rowId = (rowIds[coords.row] ??= ensureRowId(coords.row))
          const colId = (colIds[coords.col] ??= ensureColId(coords.col))
          attachFreshCell(coords.row, coords.col, cellIndex, rowId, colId)
          if (cell.formula === undefined && restoredFormula === undefined) {
            writeLiteralToCellStore(args.workbook.cellStore, cellIndex, cell.value ?? null, args.strings)
            literalColumns ??= createWrittenColumnTracker()
            markWrittenColumn(literalColumns, coords.col)
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
              : formulaValuesByAddress?.get(sheet.name)?.get(toFormulaInstanceKey(coords.row, coords.col))
            const template =
              restoredFormula.templateId !== undefined && args.resolveTemplateById
                ? args.resolveTemplateById(restoredFormula.templateId, restoredFormula.source, coords.row, coords.col)
                : undefined
            if (args.initializeHydratedPreparedCellFormulasAt && cachedValue !== undefined) {
              if (template && !template.compiled.volatile && !template.compiled.producesSpill) {
                hydratedPreparedFormulaRefs.push({
                  sheetId,
                  row: coords.row,
                  col: coords.col,
                  cellIndex,
                  source: restoredFormula.source,
                  compiled: template.compiled,
                  templateId: template.templateId,
                  value: cachedValue,
                })
                continue
              }
            }
            if (template && args.initializePreparedCellFormulasAt) {
              preparedFormulaRefs.push({
                sheetId,
                row: coords.row,
                col: coords.col,
                cellIndex,
                source: restoredFormula.source,
                compiled: template.compiled,
                templateId: template.templateId,
              })
              continue
            }
            formulaRefs.push({
              sheetId,
              cellIndex,
              mutation: {
                kind: 'setCellFormula',
                row: coords.row,
                col: coords.col,
                formula: restoredFormula.source,
              },
            })
          } else if (cell.formula !== undefined) {
            formulaRefs.push({
              sheetId,
              cellIndex,
              mutation: {
                kind: 'setCellFormula',
                row: coords.row,
                col: coords.col,
                formula: cell.formula,
              },
            })
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
    args.initializeHydratedPreparedCellFormulasAt(hydratedPreparedFormulaRefs, hydratedPreparedFormulaRefs.length)
  }
  if (preparedFormulaRefs.length > 0 && args.initializePreparedCellFormulasAt) {
    args.checkEvaluationBudget?.()
    args.initializePreparedCellFormulasAt(preparedFormulaRefs, preparedFormulaRefs.length)
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
    formulaCount: hydratedPreparedFormulaRefs.length + preparedFormulaRefs.length + formulaRefs.length,
  }
}
