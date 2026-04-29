import { parseCellAddress, type CompiledFormula } from '@bilig/formula'
import type {
  CellValue,
  LiteralInput,
  WorkbookAxisEntrySnapshot,
  WorkbookCalculationSettingsSnapshot,
  WorkbookChartSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookImageSnapshot,
  WorkbookPivotSnapshot,
  WorkbookSnapshot,
  WorkbookSortSnapshot,
  WorkbookVolatileContextSnapshot,
} from '@bilig/protocol'
import type { EngineCellMutationRef } from '../cell-mutations-at.js'
import { CellFlags } from '../cell-store.js'
import { writeLiteralToCellStore } from '../engine-value-utils.js'
import type { FormulaInstanceSnapshot } from '../formula/formula-instance-table.js'
import type { FormulaTemplateResolution, FormulaTemplateSnapshot } from '../formula/template-bank.js'
import type { StringPool } from '../string-pool.js'
import { makeCellKey, makeLogicalCellKey, type SheetRecord, type WorkbookStore } from '../workbook-store.js'

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
  readonly initializeCellFormulasAt: (refs: readonly EngineCellMutationRef[], potentialNewCells?: number) => void
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

function parseRestoredCellCoordinates(sheetName: string, cellAddress: string): RuntimeImageCellCoordinateSnapshot {
  const parsed = parseCellAddress(cellAddress, sheetName)
  return {
    row: parsed.row,
    col: parsed.col,
  }
}

interface WrittenColumnTracker {
  columns: Uint8Array
  count: number
}

interface FreshRuntimeCellPageInternals {
  readonly cells?: Map<string, number>
}

interface FreshRuntimeCellIdentityInternals {
  readonly identities?: Map<number, { readonly sheetId: number; readonly rowId: string; readonly colId: string }>
}

interface FreshRuntimeResidentCellInternals {
  readonly byCell?: Map<number, { readonly rowId: string; readonly colId: string }>
  readonly byRow?: Map<string, Set<number>>
  readonly byColumn?: Map<string, Set<number>>
}

interface FreshRuntimeLogicalSheetInternals {
  readonly cellPages?: FreshRuntimeCellPageInternals
  readonly cellIdentities?: FreshRuntimeCellIdentityInternals
  readonly residentCells?: FreshRuntimeResidentCellInternals
}

type FreshRuntimeCellAttacher = (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void

function isFreshRuntimeLogicalSheetInternals(value: unknown): value is FreshRuntimeLogicalSheetInternals {
  return typeof value === 'object' && value !== null
}

function ensureFreshRuntimeResidentSet(sets: Map<string, Set<number>>, storedSets: Map<string, Set<number>>, id: string): Set<number> {
  const cached = sets.get(id)
  if (cached) {
    return cached
  }
  let stored = storedSets.get(id)
  if (!stored) {
    stored = new Set<number>()
    storedSets.set(id, stored)
  }
  sets.set(id, stored)
  return stored
}

function createWrittenColumnTracker(): WrittenColumnTracker {
  return {
    columns: new Uint8Array(8),
    count: 0,
  }
}

function markWrittenColumn(tracker: WrittenColumnTracker, col: number): void {
  if (col >= tracker.columns.length) {
    let nextLength = tracker.columns.length
    while (nextLength <= col) {
      nextLength *= 2
    }
    const nextColumns = new Uint8Array(nextLength)
    nextColumns.set(tracker.columns)
    tracker.columns = nextColumns
  }
  if (tracker.columns[col] !== 0) {
    return
  }
  tracker.columns[col] = 1
  tracker.count += 1
}

function materializeWrittenColumns(tracker: WrittenColumnTracker): Uint32Array {
  const columns = new Uint32Array(tracker.count)
  let writeIndex = 0
  for (let col = 0; col < tracker.columns.length; col += 1) {
    if (tracker.columns[col] !== 0) {
      columns[writeIndex] = col
      writeIndex += 1
    }
  }
  return columns
}

function createFreshRuntimeCellAttacher(workbook: WorkbookStore, sheet: SheetRecord): FreshRuntimeCellAttacher {
  const logicalCandidate: unknown = sheet.logical
  const logical = isFreshRuntimeLogicalSheetInternals(logicalCandidate) ? logicalCandidate : undefined
  const cells = logical?.cellPages?.cells
  const identities = logical?.cellIdentities?.identities
  const residentByCell = logical?.residentCells?.byCell
  const residentByRow = logical?.residentCells?.byRow
  const residentByColumn = logical?.residentCells?.byColumn
  if (!cells || !identities || !residentByCell || !residentByRow || !residentByColumn) {
    return (row, col, cellIndex, rowId, colId) => {
      workbook.attachAllocatedCellWithLogicalAxisIds(sheet.id, row, col, cellIndex, rowId, colId)
    }
  }

  const rowSets = new Map<string, Set<number>>()
  const columnSets = new Map<string, Set<number>>()

  return (row, col, cellIndex, rowId, colId) => {
    cells.set(makeLogicalCellKey(sheet.id, rowId, colId), cellIndex)
    identities.set(cellIndex, { sheetId: sheet.id, rowId, colId })
    residentByCell.set(cellIndex, { rowId, colId })
    ensureFreshRuntimeResidentSet(rowSets, residentByRow, rowId).add(cellIndex)
    ensureFreshRuntimeResidentSet(columnSets, residentByColumn, colId).add(cellIndex)
    workbook.cellKeyToIndex.set(makeCellKey(sheet.id, row, col), cellIndex)
    sheet.grid.set(row, col, cellIndex)
  }
}

function restoreWorkbookMetadata(args: {
  readonly workbook: WorkbookStore
  readonly workbookMetadata:
    | {
        properties?: Array<{ key: string; value: LiteralInput }>
        definedNames?: Array<{ name: string; value: WorkbookDefinedNameValueSnapshot }>
        calculationSettings?: WorkbookCalculationSettingsSnapshot
        volatileContext?: WorkbookVolatileContextSnapshot
        tables?: readonly Parameters<WorkbookStore['setTable']>[0][]
        spills?: Array<{ sheetName: string; address: string; rows: number; cols: number }>
        pivots?: WorkbookPivotSnapshot[]
        charts?: WorkbookChartSnapshot[]
        images?: WorkbookImageSnapshot[]
        shapes?: Array<Parameters<WorkbookStore['setShape']>[0]>
        styles?: Array<Parameters<WorkbookStore['upsertCellStyle']>[0]>
        formats?: Array<Parameters<WorkbookStore['upsertCellNumberFormat']>[0]>
      }
    | undefined
}): void {
  args.workbookMetadata?.properties?.forEach(({ key, value }) => {
    args.workbook.setWorkbookProperty(key, value)
  })
  if (args.workbookMetadata?.calculationSettings) {
    args.workbook.setCalculationSettings(args.workbookMetadata.calculationSettings)
  }
  if (args.workbookMetadata?.volatileContext) {
    args.workbook.setVolatileContext(args.workbookMetadata.volatileContext)
  }
  args.workbookMetadata?.styles?.forEach((style) => {
    args.workbook.upsertCellStyle(style)
  })
  args.workbookMetadata?.formats?.forEach((format) => {
    args.workbook.upsertCellNumberFormat(format)
  })
}

function restoreSheetMetadata(args: { readonly workbook: WorkbookStore; readonly sheet: WorkbookSnapshot['sheets'][number] }): void {
  const { workbook, sheet } = args
  sheet.metadata?.rows?.forEach((entry) => {
    workbook.insertRows(sheet.name, entry.index, 1, [cloneAxisEntry(entry)])
  })
  sheet.metadata?.columns?.forEach((entry) => {
    workbook.insertColumns(sheet.name, entry.index, 1, [cloneAxisEntry(entry)])
  })
  sheet.metadata?.rowMetadata?.forEach((record) => {
    workbook.setRowMetadata(sheet.name, record.start, record.count, record.size ?? null, record.hidden ?? null)
  })
  sheet.metadata?.columnMetadata?.forEach((record) => {
    workbook.setColumnMetadata(sheet.name, record.start, record.count, record.size ?? null, record.hidden ?? null)
  })
  if (sheet.metadata?.freezePane) {
    workbook.setFreezePane(sheet.name, sheet.metadata.freezePane.rows, sheet.metadata.freezePane.cols)
  }
  if (sheet.metadata?.sheetProtection) {
    workbook.setSheetProtection(structuredClone(sheet.metadata.sheetProtection))
  }
  sheet.metadata?.styleRanges?.forEach((range) => {
    workbook.setStyleRange({ ...range.range }, range.styleId)
  })
  sheet.metadata?.formatRanges?.forEach((range) => {
    workbook.setFormatRange({ ...range.range }, range.formatId)
  })
  sheet.metadata?.filters?.forEach((range) => {
    workbook.setFilter(sheet.name, { ...range })
  })
  sheet.metadata?.sorts?.forEach((sort) => {
    workbook.setSort(sheet.name, { ...sort.range }, sort.keys.map((key) => ({ ...key })) as WorkbookSortSnapshot['keys'])
  })
  sheet.metadata?.validations?.forEach((validation) => {
    workbook.setDataValidation(structuredClone(validation))
  })
  sheet.metadata?.conditionalFormats?.forEach((format) => {
    workbook.setConditionalFormat(structuredClone(format))
  })
  sheet.metadata?.protectedRanges?.forEach((range) => {
    workbook.setRangeProtection(structuredClone(range))
  })
  sheet.metadata?.commentThreads?.forEach((thread) => {
    workbook.setCommentThread(structuredClone(thread))
  })
  sheet.metadata?.notes?.forEach((note) => {
    workbook.setNote(structuredClone(note))
  })
}

function cloneAxisEntry(entry: WorkbookAxisEntrySnapshot): WorkbookAxisEntrySnapshot {
  return {
    id: entry.id,
    index: entry.index,
    ...(entry.size !== undefined ? { size: entry.size } : {}),
    ...(entry.hidden !== undefined ? { hidden: entry.hidden } : {}),
  }
}

function restoreMetadataStructures(args: {
  readonly workbook: WorkbookStore
  readonly workbookMetadata:
    | {
        definedNames?: Array<{ name: string; value: WorkbookDefinedNameValueSnapshot }>
        tables?: readonly Parameters<WorkbookStore['setTable']>[0][]
        spills?: Array<{ sheetName: string; address: string; rows: number; cols: number }>
        pivots?: WorkbookPivotSnapshot[]
        charts?: WorkbookChartSnapshot[]
        images?: WorkbookImageSnapshot[]
        shapes?: Array<Parameters<WorkbookStore['setShape']>[0]>
      }
    | undefined
}): void {
  args.workbookMetadata?.tables?.forEach((table) => {
    args.workbook.setTable(structuredClone(table))
  })
  args.workbookMetadata?.definedNames?.forEach(({ name, value }) => {
    args.workbook.setDefinedName(name, structuredClone(value))
  })
  args.workbookMetadata?.spills?.forEach((spill) => {
    args.workbook.setSpill(spill.sheetName, spill.address, spill.rows, spill.cols)
  })
  args.workbookMetadata?.pivots?.forEach((pivot) => {
    args.workbook.setPivot(structuredClone(pivot))
  })
}

function restoreVisualMetadata(args: {
  readonly workbook: WorkbookStore
  readonly workbookMetadata:
    | {
        charts?: WorkbookChartSnapshot[]
        images?: WorkbookImageSnapshot[]
        shapes?: Array<Parameters<WorkbookStore['setShape']>[0]>
      }
    | undefined
}): void {
  args.workbookMetadata?.charts?.forEach((chart) => {
    args.workbook.setChart(structuredClone(chart))
  })
  args.workbookMetadata?.images?.forEach((image) => {
    args.workbook.setImage(structuredClone(image))
  })
  args.workbookMetadata?.shapes?.forEach((shape) => {
    args.workbook.setShape(structuredClone(shape))
  })
}

function restoreWorkbookStructure(args: {
  readonly snapshot: WorkbookSnapshot
  readonly workbook: WorkbookStore
  readonly resetWorkbook: (workbookName?: string) => void
}): readonly WorkbookSnapshot['sheets'][number][] {
  args.resetWorkbook(args.snapshot.workbook.name)
  restoreWorkbookMetadata({
    workbook: args.workbook,
    workbookMetadata: args.snapshot.workbook.metadata,
  })

  const orderedSheets = [...args.snapshot.sheets].toSorted((left, right) => left.order - right.order)
  orderedSheets.forEach((sheet) => {
    args.workbook.createSheet(sheet.name, sheet.order, sheet.id)
  })
  orderedSheets.forEach((sheet) => {
    restoreSheetMetadata({ workbook: args.workbook, sheet })
  })
  restoreMetadataStructures({
    workbook: args.workbook,
    workbookMetadata: args.snapshot.workbook.metadata,
  })
  return orderedSheets
}

function restoreLiteralCell(args: {
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly cellIndex: number
  readonly value: LiteralInput
}): void {
  writeLiteralToCellStore(args.workbook.cellStore, args.cellIndex, args.value, args.strings)
  if (args.value === null) {
    args.workbook.cellStore.flags[args.cellIndex] = (args.workbook.cellStore.flags[args.cellIndex] ?? 0) | CellFlags.AuthoredBlank
  }
}

export function restoreWorkbookFromSnapshot(args: WorkbookSnapshotRestoreArgs): WorkbookRestoreResult {
  const orderedSheets = restoreWorkbookStructure(args)
  const potentialNewCells = orderedSheets.reduce((count, sheet) => count + sheet.cells.length, 0)
  const formulaRefs: EngineCellMutationRef[] = []

  args.workbook.cellStore.ensureCapacity(args.workbook.cellStore.size + potentialNewCells)
  args.workbook.withBatchedColumnVersionUpdates(() => {
    for (let sheetIndex = 0; sheetIndex < orderedSheets.length; sheetIndex += 1) {
      const sheet = orderedSheets[sheetIndex]!
      const sheetId = args.workbook.getSheet(sheet.name)?.id
      if (sheetId === undefined) {
        throw new Error(`Missing sheet during snapshot restore: ${sheet.name}`)
      }
      for (let cellIndex = 0; cellIndex < sheet.cells.length; cellIndex += 1) {
        const cell = sheet.cells[cellIndex]!
        const coords = parseCellAddress(cell.address, sheet.name)
        const ensured = args.workbook.ensureCellAt(sheetId, coords.row, coords.col)
        if (cell.formula !== undefined) {
          formulaRefs.push({
            sheetId,
            mutation: {
              kind: 'setCellFormula',
              row: coords.row,
              col: coords.col,
              formula: cell.formula,
            },
          })
        } else {
          restoreLiteralCell({
            workbook: args.workbook,
            strings: args.strings,
            cellIndex: ensured.cellIndex,
            value: cell.value ?? null,
          })
        }
        if (cell.format !== undefined) {
          args.workbook.setCellFormat(ensured.cellIndex, cell.format)
        }
      }
    }
  })

  if (formulaRefs.length > 0) {
    args.initializeCellFormulasAt(formulaRefs, potentialNewCells)
  }

  restoreVisualMetadata({
    workbook: args.workbook,
    workbookMetadata: args.snapshot.workbook.metadata,
  })
  return { formulaCount: formulaRefs.length }
}

export function restoreWorkbookFromRuntimeImage(args: RuntimeImageRestoreArgs): WorkbookRestoreResult {
  const orderedSheets = restoreWorkbookStructure(args)

  args.hydrateTemplateBank(args.runtimeImage.templateBank)

  const formulaValueIndexAligned = formulaValuesAreAligned(args.runtimeImage.formulaInstances, args.runtimeImage.formulaValues)
  const formulaValuesByAddress = formulaValueIndexAligned ? undefined : new Map<string, Map<number, CellValue>>()
  if (formulaValuesByAddress) {
    args.runtimeImage.formulaValues.forEach((record) => {
      getOrCreateSheetFormulaMap(formulaValuesByAddress, record.sheetName).set(toFormulaInstanceKey(record.row, record.col), record.value)
    })
  }
  const formulaSpansBySheet = buildRuntimeFormulaSheetSpans(args.runtimeImage.formulaInstances)
  const sheetCellsByName = new Map<string, readonly RuntimeImageCellCoordinateSnapshot[]>(
    (args.runtimeImage.sheetCells ?? []).map((record) => [record.sheetName, record.coords]),
  )
  const totalCellCount = orderedSheets.reduce((sum, sheet) => sum + sheet.cells.length, 0)
  if (totalCellCount > 0) {
    args.workbook.cellStore.ensureCapacity(args.workbook.cellStore.size + totalCellCount)
  }

  const formulaRefs: EngineCellMutationRef[] = []
  const preparedFormulaRefs: PreparedRuntimeFormulaRef[] = []
  const hydratedPreparedFormulaRefs: HydratedPreparedRuntimeFormulaRef[] = []
  const previousOnSetValue = args.workbook.cellStore.onSetValue
  args.workbook.cellStore.onSetValue = null
  try {
    orderedSheets.forEach((sheet) => {
      const sheetRecord = args.workbook.getSheet(sheet.name)
      if (!sheetRecord) {
        throw new Error(`Missing sheet during runtime image restore: ${sheet.name}`)
      }
      const sheetId = sheetRecord.id
      const sheetCoords = sheetCellsByName.get(sheet.name)
      const rowIds: string[] = []
      const colIds: string[] = []
      const attachFreshCell = createFreshRuntimeCellAttacher(args.workbook, sheetRecord)
      const formulaSpan = formulaSpansBySheet.get(sheet.name)
      let formulaInstanceIndex = formulaSpan?.start ?? 0
      const formulaInstanceEnd = formulaSpan?.end ?? formulaInstanceIndex
      let literalColumns: WrittenColumnTracker | undefined
      for (let index = 0; index < sheet.cells.length; index += 1) {
        const cell = sheet.cells[index]!
        const coords = sheetCoords?.[index] ?? parseRestoredCellCoordinates(sheet.name, cell.address)
        while (
          formulaInstanceIndex < formulaInstanceEnd &&
          compareFormulaInstanceToCoordinate(args.runtimeImage.formulaInstances[formulaInstanceIndex]!, coords) < 0
        ) {
          formulaInstanceIndex += 1
        }
        const candidateFormula =
          formulaInstanceIndex < formulaInstanceEnd ? args.runtimeImage.formulaInstances[formulaInstanceIndex] : undefined
        const restoredFormula =
          candidateFormula && compareFormulaInstanceToCoordinate(candidateFormula, coords) === 0 ? candidateFormula : undefined
        const cellIndex = args.workbook.cellStore.allocateReserved(sheetId, coords.row, coords.col)
        const rowId = (rowIds[coords.row] ??= args.workbook.ensureLogicalAxisId(sheetId, 'row', coords.row))
        const colId = (colIds[coords.col] ??= args.workbook.ensureLogicalAxisId(sheetId, 'column', coords.col))
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

  if (hydratedPreparedFormulaRefs.length > 0 && args.initializeHydratedPreparedCellFormulasAt) {
    args.initializeHydratedPreparedCellFormulasAt(hydratedPreparedFormulaRefs, hydratedPreparedFormulaRefs.length)
  }
  if (preparedFormulaRefs.length > 0 && args.initializePreparedCellFormulasAt) {
    args.initializePreparedCellFormulasAt(preparedFormulaRefs, preparedFormulaRefs.length)
  }
  if (formulaRefs.length > 0) {
    args.initializeCellFormulasAt(formulaRefs, formulaRefs.length)
  }

  restoreVisualMetadata({
    workbook: args.workbook,
    workbookMetadata: args.snapshot.workbook.metadata,
  })
  return {
    formulaCount: hydratedPreparedFormulaRefs.length + preparedFormulaRefs.length + formulaRefs.length,
  }
}
