import { parseCellAddress, type CompiledFormula } from '@bilig/formula'
import type {
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
}

export interface PreparedRuntimeFormulaRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId?: number
}

function toFormulaInstanceKey(sheetName: string, row: number, col: number): string {
  return `${sheetName}\t${row}\t${col}`
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

export function restoreWorkbookFromRuntimeImage(args: RuntimeImageRestoreArgs): void {
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

  args.hydrateTemplateBank(args.runtimeImage.templateBank)

  const formulaInstancesByAddress = new Map<string, FormulaInstanceSnapshot>()
  args.runtimeImage.formulaInstances.forEach((record) => {
    formulaInstancesByAddress.set(toFormulaInstanceKey(record.sheetName, record.row, record.col), record)
  })

  const formulaRefs: EngineCellMutationRef[] = []
  const preparedFormulaRefs: PreparedRuntimeFormulaRef[] = []
  orderedSheets.forEach((sheet) => {
    const sheetId = args.workbook.getSheet(sheet.name)?.id
    if (sheetId === undefined) {
      throw new Error(`Missing sheet during runtime image restore: ${sheet.name}`)
    }
    sheet.cells.forEach((cell) => {
      const parsed = parseCellAddress(cell.address, sheet.name)
      const formulaInstance = formulaInstancesByAddress.get(toFormulaInstanceKey(sheet.name, parsed.row, parsed.col))
      const ensured = args.workbook.ensureCellAt(sheetId, parsed.row, parsed.col)
      if (cell.formula === undefined && formulaInstance === undefined) {
        writeLiteralToCellStore(args.workbook.cellStore, ensured.cellIndex, cell.value ?? null, args.strings)
        if (cell.value === null) {
          args.workbook.cellStore.flags[ensured.cellIndex] =
            (args.workbook.cellStore.flags[ensured.cellIndex] ?? 0) | CellFlags.AuthoredBlank
        }
      }
      if (cell.format !== undefined) {
        args.workbook.setCellFormat(ensured.cellIndex, cell.format)
      }
      if (formulaInstance) {
        if (formulaInstance.templateId !== undefined && args.resolveTemplateById && args.initializePreparedCellFormulasAt) {
          const template = args.resolveTemplateById(formulaInstance.templateId, formulaInstance.source, parsed.row, parsed.col)
          if (template) {
            preparedFormulaRefs.push({
              sheetId,
              row: parsed.row,
              col: parsed.col,
              source: formulaInstance.source,
              compiled: template.compiled,
              templateId: template.templateId,
            })
            return
          }
        }
        formulaRefs.push({
          sheetId,
          mutation: {
            kind: 'setCellFormula',
            row: parsed.row,
            col: parsed.col,
            formula: formulaInstance.source,
          },
        })
      } else if (cell.formula !== undefined) {
        formulaRefs.push({
          sheetId,
          mutation: {
            kind: 'setCellFormula',
            row: parsed.row,
            col: parsed.col,
            formula: cell.formula,
          },
        })
      }
    })
  })

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
}
