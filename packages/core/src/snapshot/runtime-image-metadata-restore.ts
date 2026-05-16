import type {
  LiteralInput,
  WorkbookAxisEntrySnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookCalculationSettingsSnapshot,
  WorkbookChartSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookImageSnapshot,
  WorkbookMacroPayloadSnapshot,
  WorkbookPivotSnapshot,
  WorkbookProtectionSnapshot,
  WorkbookSnapshot,
  WorkbookSortSnapshot,
  WorkbookVolatileContextSnapshot,
} from '@bilig/protocol'
import { axisGeometryKeys, syncAxisMetadataBucket } from '../workbook-axis-records.js'
import type { WorkbookAxisEntryRecord, WorkbookAxisMetadataRecord, WorkbookStore } from '../workbook-store.js'

function restoreWorkbookMetadata(args: {
  readonly workbook: WorkbookStore
  readonly workbookMetadata:
    | {
        properties?: Array<{ key: string; value: LiteralInput }>
        workbookProtection?: WorkbookProtectionSnapshot
        macroPayloads?: WorkbookMacroPayloadSnapshot[]
        definedNames?: Array<{ name: string; scopeSheetName?: string; value: WorkbookDefinedNameValueSnapshot }>
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
  if (args.workbookMetadata?.workbookProtection) {
    args.workbook.setWorkbookProtection(args.workbookMetadata.workbookProtection)
  }
  args.workbookMetadata?.macroPayloads?.forEach((payload) => {
    args.workbook.setMacroPayload(payload)
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
  restoreAxisMetadata({
    workbook,
    sheetName: sheet.name,
    axis: 'row',
    records: sheet.metadata?.rowMetadata ?? [],
  })
  restoreAxisMetadata({
    workbook,
    sheetName: sheet.name,
    axis: 'column',
    records: sheet.metadata?.columnMetadata ?? [],
  })
  if (sheet.metadata?.sheetFormatPr) {
    workbook.setSheetFormatPr(sheet.name, sheet.metadata.sheetFormatPr)
  }
  if (sheet.metadata?.freezePane) {
    workbook.setFreezePane(sheet.name, sheet.metadata.freezePane.rows, sheet.metadata.freezePane.cols, {
      ...(sheet.metadata.freezePane.topLeftCell !== undefined ? { topLeftCell: sheet.metadata.freezePane.topLeftCell } : {}),
      ...(sheet.metadata.freezePane.activePane !== undefined ? { activePane: sheet.metadata.freezePane.activePane } : {}),
    })
  }
  if (sheet.metadata?.tabColor) {
    workbook.setSheetTabColor(sheet.name, structuredClone(sheet.metadata.tabColor))
  }
  if (sheet.metadata?.merges && sheet.metadata.merges.length > 0) {
    workbook.setMergeRanges(
      sheet.name,
      sheet.metadata.merges.map((range) => ({ ...range, sheetName: range.sheetName ?? sheet.name })),
    )
  }
  if (sheet.metadata?.sheetProtection) {
    workbook.setSheetProtection(structuredClone(sheet.metadata.sheetProtection))
  }
  if (sheet.metadata?.styleRanges && sheet.metadata.styleRanges.length > 0) {
    workbook.setStyleRanges(sheet.name, sheet.metadata.styleRanges)
  }
  if (sheet.metadata?.formatRanges && sheet.metadata.formatRanges.length > 0) {
    workbook.setFormatRanges(sheet.name, sheet.metadata.formatRanges)
  }
  sheet.metadata?.filters?.forEach((range) => {
    workbook.setFilter(sheet.name, structuredClone(range))
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

function restoreAxisMetadata(args: {
  readonly workbook: WorkbookStore
  readonly sheetName: string
  readonly axis: 'row' | 'column'
  readonly records: readonly WorkbookAxisMetadataSnapshot[]
}): void {
  if (args.records.length === 0) {
    return
  }
  const sheet = args.workbook.getSheet(args.sheetName)
  if (!sheet) {
    return
  }
  const entries = args.axis === 'row' ? sheet.rowAxis : sheet.columnAxis
  for (const record of args.records) {
    if (record.count <= 0) {
      continue
    }
    if (args.axis === 'row') {
      args.workbook.materializeRowAxisEntries(args.sheetName, record.start, record.count)
    } else {
      args.workbook.materializeColumnAxisEntries(args.sheetName, record.start, record.count)
    }
    for (let offset = 0; offset < record.count; offset += 1) {
      const entry = entries[record.start + offset]
      if (entry) {
        applyAxisMetadataRecord(entry, record)
      }
    }
  }
  const bucket: Map<string, WorkbookAxisMetadataRecord> =
    args.axis === 'row' ? args.workbook.metadata.rowMetadata : args.workbook.metadata.columnMetadata
  syncAxisMetadataBucket(bucket, args.sheetName, entries)
}

function applyAxisMetadataRecord(entry: WorkbookAxisEntryRecord, record: WorkbookAxisMetadataSnapshot): void {
  entry.size = record.size ?? null
  entry.hidden = record.hidden ?? null
  for (const key of axisGeometryKeys) {
    const value = record[key]
    if (value === undefined) {
      delete entry[key]
    } else {
      Object.assign(entry, { [key]: value })
    }
  }
}

function cloneAxisEntry(entry: WorkbookAxisEntrySnapshot): WorkbookAxisEntrySnapshot {
  return {
    id: entry.id,
    index: entry.index,
    ...(entry.size !== undefined ? { size: entry.size } : {}),
    ...(entry.hidden !== undefined ? { hidden: entry.hidden } : {}),
    ...cloneAxisGeometry(entry),
  }
}

function cloneAxisGeometry(entry: {
  readonly styleIndex?: number | null
  readonly xlsxWidth?: number | null
  readonly xlsxHeight?: number | null
  readonly customFormat?: boolean | null
  readonly customWidth?: boolean | null
  readonly bestFit?: boolean | null
  readonly outlineLevel?: number | null
  readonly collapsed?: boolean | null
  readonly customHeight?: boolean | null
  readonly thickTop?: boolean | null
  readonly thickBottom?: boolean | null
}): {
  styleIndex?: number | null
  xlsxWidth?: number | null
  xlsxHeight?: number | null
  customFormat?: boolean | null
  customWidth?: boolean | null
  bestFit?: boolean | null
  outlineLevel?: number | null
  collapsed?: boolean | null
  customHeight?: boolean | null
  thickTop?: boolean | null
  thickBottom?: boolean | null
} {
  return {
    ...(entry.styleIndex !== undefined ? { styleIndex: entry.styleIndex } : {}),
    ...(entry.xlsxWidth !== undefined ? { xlsxWidth: entry.xlsxWidth } : {}),
    ...(entry.xlsxHeight !== undefined ? { xlsxHeight: entry.xlsxHeight } : {}),
    ...(entry.customFormat !== undefined ? { customFormat: entry.customFormat } : {}),
    ...(entry.customWidth !== undefined ? { customWidth: entry.customWidth } : {}),
    ...(entry.bestFit !== undefined ? { bestFit: entry.bestFit } : {}),
    ...(entry.outlineLevel !== undefined ? { outlineLevel: entry.outlineLevel } : {}),
    ...(entry.collapsed !== undefined ? { collapsed: entry.collapsed } : {}),
    ...(entry.customHeight !== undefined ? { customHeight: entry.customHeight } : {}),
    ...(entry.thickTop !== undefined ? { thickTop: entry.thickTop } : {}),
    ...(entry.thickBottom !== undefined ? { thickBottom: entry.thickBottom } : {}),
  }
}

function restoreMetadataStructures(args: {
  readonly workbook: WorkbookStore
  readonly workbookMetadata:
    | {
        definedNames?: Array<{ name: string; scopeSheetName?: string; value: WorkbookDefinedNameValueSnapshot }>
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
  args.workbookMetadata?.definedNames?.forEach(({ name, scopeSheetName, value }) => {
    args.workbook.setDefinedName(name, structuredClone(value), scopeSheetName)
  })
  args.workbookMetadata?.spills?.forEach((spill) => {
    args.workbook.setSpill(spill.sheetName, spill.address, spill.rows, spill.cols)
  })
  args.workbookMetadata?.pivots?.forEach((pivot) => {
    args.workbook.setPivot(structuredClone(pivot))
  })
}

export function restoreVisualMetadata(args: {
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

export function restoreWorkbookStructure(args: {
  readonly snapshot: WorkbookSnapshot
  readonly workbook: WorkbookStore
  readonly resetWorkbook: (workbookName?: string) => void
  readonly checkEvaluationBudget?: (stepCost?: number) => void
}): readonly WorkbookSnapshot['sheets'][number][] {
  args.checkEvaluationBudget?.()
  args.resetWorkbook(args.snapshot.workbook.name)
  restoreWorkbookMetadata({
    workbook: args.workbook,
    workbookMetadata: args.snapshot.workbook.metadata,
  })

  const orderedSheets = [...args.snapshot.sheets].toSorted((left, right) => left.order - right.order)
  orderedSheets.forEach((sheet) => {
    args.checkEvaluationBudget?.()
    args.workbook.createSheet(sheet.name, sheet.order, sheet.id)
  })
  orderedSheets.forEach((sheet) => {
    args.checkEvaluationBudget?.()
    restoreSheetMetadata({ workbook: args.workbook, sheet })
  })
  args.checkEvaluationBudget?.()
  restoreMetadataStructures({
    workbook: args.workbook,
    workbookMetadata: args.snapshot.workbook.metadata,
  })
  return orderedSheets
}
