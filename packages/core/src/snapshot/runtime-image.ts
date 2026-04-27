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
import type { WorkbookStore } from '../workbook-store.js'

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
}

export interface RuntimeImageCellCoordinateSnapshot {
  readonly row: number
  readonly col: number
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
}

export interface HydratedPreparedRuntimeFormulaRef extends PreparedRuntimeFormulaRef {
  readonly value: CellValue
}

interface RestoredFormulaInstance {
  readonly record: FormulaInstanceSnapshot
  readonly value?: CellValue
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

function resolveRestoredCellCoordinates(args: {
  readonly sheetName: string
  readonly cellAddress: string
  readonly indexedCoordinate: RuntimeImageCellCoordinateSnapshot | undefined
}): RuntimeImageCellCoordinateSnapshot {
  if (args.indexedCoordinate) {
    return args.indexedCoordinate
  }
  const parsed = parseCellAddress(args.cellAddress, args.sheetName)
  return {
    row: parsed.row,
    col: parsed.col,
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

  let formulaValuesByAddress: Map<string, Map<number, CellValue>> | undefined
  const formulaInstancesByAddress = new Map<string, Map<number, RestoredFormulaInstance>>()
  args.runtimeImage.formulaInstances.forEach((record, index) => {
    const valueRecord = args.runtimeImage.formulaValues[index]
    if (!formulaValueMatchesInstance(record, valueRecord)) {
      formulaValuesByAddress ??= new Map<string, Map<number, CellValue>>()
    }
    getOrCreateSheetFormulaMap(formulaInstancesByAddress, record.sheetName).set(toFormulaInstanceKey(record.row, record.col), {
      record,
      ...(formulaValueMatchesInstance(record, valueRecord) ? { value: valueRecord.value } : {}),
    })
  })
  if (formulaValuesByAddress) {
    args.runtimeImage.formulaValues.forEach((record) => {
      getOrCreateSheetFormulaMap(formulaValuesByAddress!, record.sheetName).set(toFormulaInstanceKey(record.row, record.col), record.value)
    })
  }
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
  orderedSheets.forEach((sheet) => {
    const sheetId = args.workbook.getSheet(sheet.name)?.id
    if (sheetId === undefined) {
      throw new Error(`Missing sheet during runtime image restore: ${sheet.name}`)
    }
    const sheetCoords = sheetCellsByName.get(sheet.name)
    for (let index = 0; index < sheet.cells.length; index += 1) {
      const cell = sheet.cells[index]!
      const coords = resolveRestoredCellCoordinates({
        sheetName: sheet.name,
        cellAddress: cell.address,
        indexedCoordinate: sheetCoords?.[index],
      })
      const coordKey = toFormulaInstanceKey(coords.row, coords.col)
      const restoredFormula = formulaInstancesByAddress.get(sheet.name)?.get(coordKey)
      const formulaInstance = restoredFormula?.record
      const cellIndex = args.workbook.cellStore.allocateReserved(sheetId, coords.row, coords.col)
      args.workbook.attachAllocatedCell(sheetId, coords.row, coords.col, cellIndex)
      if (cell.formula === undefined && formulaInstance === undefined) {
        writeLiteralToCellStore(args.workbook.cellStore, cellIndex, cell.value ?? null, args.strings)
        if (cell.value === null) {
          args.workbook.cellStore.flags[cellIndex] = (args.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.AuthoredBlank
        }
      }
      if (cell.format !== undefined) {
        args.workbook.setCellFormat(cellIndex, cell.format)
      }
      if (formulaInstance) {
        const cachedValue = restoredFormula.value ?? formulaValuesByAddress?.get(sheet.name)?.get(coordKey)
        if (
          formulaInstance.templateId !== undefined &&
          args.resolveTemplateById &&
          args.initializeHydratedPreparedCellFormulasAt &&
          cachedValue !== undefined
        ) {
          const template = args.resolveTemplateById(formulaInstance.templateId, formulaInstance.source, coords.row, coords.col)
          if (template && !template.compiled.volatile && !template.compiled.producesSpill) {
            hydratedPreparedFormulaRefs.push({
              sheetId,
              row: coords.row,
              col: coords.col,
              source: formulaInstance.source,
              compiled: template.compiled,
              templateId: template.templateId,
              value: cachedValue,
            })
            continue
          }
        }
        if (formulaInstance.templateId !== undefined && args.resolveTemplateById && args.initializePreparedCellFormulasAt) {
          const template = args.resolveTemplateById(formulaInstance.templateId, formulaInstance.source, coords.row, coords.col)
          if (template) {
            preparedFormulaRefs.push({
              sheetId,
              row: coords.row,
              col: coords.col,
              source: formulaInstance.source,
              compiled: template.compiled,
              templateId: template.templateId,
            })
            continue
          }
        }
        formulaRefs.push({
          sheetId,
          mutation: {
            kind: 'setCellFormula',
            row: coords.row,
            col: coords.col,
            formula: formulaInstance.source,
          },
        })
      } else if (cell.formula !== undefined) {
        formulaRefs.push({
          sheetId,
          mutation: {
            kind: 'setCellFormula',
            row: coords.row,
            col: coords.col,
            formula: cell.formula,
          },
        })
      }
    }
  })

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
