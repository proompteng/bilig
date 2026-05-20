import { formatAddress, parseCellAddress } from '@bilig/formula'
import type {
  CellRangeRef,
  CellStyleRecord,
  LiteralInput,
  SheetMetadataSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookCalculationSettingsSnapshot,
  WorkbookChartSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookPivotSnapshot,
  WorkbookPivotValueSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookSheetProtectionSnapshot,
  WorkbookSortSnapshot,
  WorkbookSnapshot,
  WorkbookTableSnapshot,
} from '@bilig/protocol'

type ComparableRangeRef = {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
}

type WorkbookSheetSnapshot = WorkbookSnapshot['sheets'][number]
type WorkbookAxisEntrySemantics = Omit<WorkbookAxisEntrySnapshot, 'id'>
type WorkbookSheetMetadataSemantics = Omit<SheetMetadataSnapshot, 'rows' | 'columns'> & {
  rows?: WorkbookAxisEntrySemantics[]
  columns?: WorkbookAxisEntrySemantics[]
}

export interface WorkbookSemanticComparableSnapshot extends Omit<WorkbookSnapshot, 'sheets'> {
  sheets: Array<Omit<WorkbookSheetSnapshot, 'metadata'> & { metadata?: WorkbookSheetMetadataSemantics }>
}

export interface ProjectedWorkbookChartSemantics {
  readonly id: string
  readonly sheetName: string
  readonly address: string
  readonly source: CellRangeRef
  readonly chartType: WorkbookChartSnapshot['chartType']
  readonly seriesOrientation: WorkbookChartSnapshot['seriesOrientation']
  readonly firstRowAsHeaders?: boolean
  readonly firstColumnAsLabels?: boolean
  readonly title?: string
  readonly legendPosition?: WorkbookChartSnapshot['legendPosition']
  readonly rows: number
  readonly cols: number
}

export interface ProjectedWorkbookPortableStyle {
  readonly fill?: CellStyleRecord['fill']
  readonly font?: CellStyleRecord['font']
  readonly alignment?: CellStyleRecord['alignment']
  readonly borders?: CellStyleRecord['borders']
  readonly protection?: CellStyleRecord['protection']
}

export interface ProjectedWorkbookStyleRangeSemantics {
  readonly range: CellRangeRef
  readonly style: ProjectedWorkbookPortableStyle | undefined
}

export interface ProjectedWorkbookValueFormulaFormatSheet {
  readonly name: string
  readonly order: number
  readonly cells: Array<{
    readonly address: string
    readonly value?: WorkbookSnapshot['sheets'][number]['cells'][number]['value']
    readonly formula?: string
    readonly format?: string
  }>
}

interface ProjectedWorkbookDimensionAxisEntry {
  readonly index: number
  readonly size?: number | null
}

export interface ProjectedWorkbookDimensionSheet {
  readonly name: string
  readonly columns: ProjectedWorkbookDimensionAxisEntry[]
  readonly rows: ProjectedWorkbookDimensionAxisEntry[]
  readonly merges: Array<{
    readonly sheetName: string
    readonly startAddress: string
    readonly endAddress: string
  }>
}

export interface WorkbookSemanticSnapshot {
  readonly properties: Array<{
    readonly key: string
    readonly value: LiteralInput
  }>
  readonly calculationSettings: WorkbookCalculationSettingsSnapshot | undefined
  readonly definedNames: Array<{
    readonly name: string
    readonly value: WorkbookDefinedNameValueSnapshot
  }>
  readonly commentThreads: Array<{
    readonly sheetName: string
    readonly address: string
    readonly comments: Array<{
      readonly body: string
      readonly authorDisplayName?: string
    }>
  }>
  readonly styleRanges: ProjectedWorkbookStyleRangeSemantics[]
  readonly tables: WorkbookTableSnapshot[]
  readonly charts: ProjectedWorkbookChartSemantics[]
  readonly pivots: WorkbookPivotSnapshot[]
  readonly validations: WorkbookDataValidationSnapshot[]
  readonly conditionalFormats: ReturnType<typeof projectConditionalFormatSemantics>[]
  readonly freezePanes: Array<{
    readonly sheetName: string
    readonly freezePane: WorkbookFreezePaneSnapshot
  }>
  readonly filters: Array<{
    readonly sheetName: string
    readonly startAddress: string
    readonly endAddress: string
  }>
  readonly sorts: WorkbookSortSnapshot[]
  readonly sheetProtections: WorkbookSheetProtectionSnapshot[]
  readonly protectedRanges: WorkbookRangeProtectionSnapshot[]
  readonly valueFormulaFormatSheets: ProjectedWorkbookValueFormulaFormatSheet[]
  readonly dimensionSheets: ProjectedWorkbookDimensionSheet[]
}

export function normalizeWorkbookSnapshotForSemanticComparison(snapshot: WorkbookSnapshot): WorkbookSemanticComparableSnapshot {
  const clone: WorkbookSemanticComparableSnapshot = structuredClone(snapshot)
  if (clone.workbook.metadata) {
    if (clone.workbook.metadata.properties) {
      clone.workbook.metadata.properties = clone.workbook.metadata.properties.toSorted((left, right) => left.key.localeCompare(right.key))
    }
    if (clone.workbook.metadata.definedNames) {
      clone.workbook.metadata.definedNames = clone.workbook.metadata.definedNames.toSorted((left, right) =>
        left.name.localeCompare(right.name),
      )
    }
    if (clone.workbook.metadata.tables) {
      clone.workbook.metadata.tables = clone.workbook.metadata.tables.toSorted((left, right) => left.name.localeCompare(right.name))
    }
    if (clone.workbook.metadata.styles) {
      clone.workbook.metadata.styles = clone.workbook.metadata.styles.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (clone.workbook.metadata.formats) {
      clone.workbook.metadata.formats = clone.workbook.metadata.formats.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (clone.workbook.metadata.pivots) {
      clone.workbook.metadata.pivots = clone.workbook.metadata.pivots.toSorted((left, right) =>
        `${left.sheetName}!${left.address}:${left.name}`.localeCompare(`${right.sheetName}!${right.address}:${right.name}`),
      )
    }
    if (clone.workbook.metadata.charts) {
      clone.workbook.metadata.charts = clone.workbook.metadata.charts.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (clone.workbook.metadata.images) {
      clone.workbook.metadata.images = clone.workbook.metadata.images.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (clone.workbook.metadata.shapes) {
      clone.workbook.metadata.shapes = clone.workbook.metadata.shapes.toSorted((left, right) => left.id.localeCompare(right.id))
    }
  }
  clone.sheets = snapshot.sheets.map((sheet) => {
    const cells = sheet.cells.filter((cell) => cell.formula !== undefined || cell.format !== undefined || cell.value !== null)
    if (!sheet.metadata) {
      return {
        ...sheet,
        cells,
      }
    }
    const metadata: WorkbookSheetMetadataSemantics = { ...sheet.metadata }
    if (sheet.metadata.rows) {
      metadata.rows = sheet.metadata.rows.map(({ id: _id, ...rest }) => rest).toSorted((left, right) => left.index - right.index)
    }
    if (sheet.metadata.columns) {
      metadata.columns = sheet.metadata.columns.map(({ id: _id, ...rest }) => rest).toSorted((left, right) => left.index - right.index)
    }
    if (metadata.rowMetadata) {
      metadata.rowMetadata = metadata.rowMetadata.toSorted((left, right) => left.start - right.start || left.count - right.count)
    }
    if (metadata.columnMetadata) {
      metadata.columnMetadata = metadata.columnMetadata.toSorted((left, right) => left.start - right.start || left.count - right.count)
    }
    if (metadata.styleRanges) {
      metadata.styleRanges = normalizeRangeRecords(metadata.styleRanges, 'styleId')
    }
    if (metadata.formatRanges) {
      metadata.formatRanges = normalizeRangeRecords(metadata.formatRanges, 'formatId')
    }
    if (metadata.filters) {
      metadata.filters = metadata.filters.toSorted(compareRangeRefs)
    }
    if (metadata.sorts) {
      metadata.sorts = metadata.sorts.toSorted((left, right) => {
        const rangeComparison = compareRangeRefs(left.range, right.range)
        if (rangeComparison !== 0) {
          return rangeComparison
        }
        return JSON.stringify(left.keys).localeCompare(JSON.stringify(right.keys))
      })
    }
    if (metadata.validations) {
      metadata.validations = metadata.validations.toSorted((left, right) => {
        const rangeComparison = compareRangeRefs(left.range, right.range)
        if (rangeComparison !== 0) {
          return rangeComparison
        }
        return JSON.stringify(left.rule).localeCompare(JSON.stringify(right.rule))
      })
    }
    if (metadata.conditionalFormats) {
      metadata.conditionalFormats = metadata.conditionalFormats.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (metadata.protectedRanges) {
      metadata.protectedRanges = metadata.protectedRanges.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (metadata.commentThreads) {
      metadata.commentThreads = metadata.commentThreads.toSorted((left, right) =>
        `${left.sheetName}!${left.address}:${left.threadId}`.localeCompare(`${right.sheetName}!${right.address}:${right.threadId}`),
      )
    }
    if (metadata.notes) {
      metadata.notes = metadata.notes.toSorted((left, right) =>
        `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`),
      )
    }
    return {
      ...sheet,
      cells,
      metadata,
    }
  })
  return clone
}

export function projectWorkbookSemanticSnapshot(snapshot: WorkbookSnapshot): WorkbookSemanticSnapshot {
  const stylesById = new Map((snapshot.workbook.metadata?.styles ?? []).map((style) => [style.id, style]))
  const portableStyle = (styleId: string): ProjectedWorkbookPortableStyle | undefined => {
    const style = stylesById.get(styleId)
    if (!style) {
      return undefined
    }
    return {
      ...(style.fill ? { fill: style.fill } : {}),
      ...(style.font ? { font: style.font } : {}),
      ...(style.alignment ? { alignment: style.alignment } : {}),
      ...(style.borders ? { borders: style.borders } : {}),
      ...(style.protection ? { protection: style.protection } : {}),
    }
  }
  return {
    properties: (snapshot.workbook.metadata?.properties ?? [])
      .map((property) => ({ key: property.key, value: property.value }))
      .toSorted((left, right) => left.key.localeCompare(right.key)),
    calculationSettings: snapshot.workbook.metadata?.calculationSettings,
    definedNames: (snapshot.workbook.metadata?.definedNames ?? [])
      .map((definedName) => ({ name: definedName.name, value: definedName.value }))
      .toSorted((left, right) => left.name.localeCompare(right.name)),
    commentThreads: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.commentThreads ?? [])
      .map((thread) => ({
        sheetName: thread.sheetName,
        address: thread.address,
        comments: thread.comments.map((comment) => ({
          body: comment.body,
          ...(comment.authorDisplayName !== undefined ? { authorDisplayName: comment.authorDisplayName } : {}),
        })),
      }))
      .toSorted((left, right) => `${left.sheetName}:${left.address}`.localeCompare(`${right.sheetName}:${right.address}`)),
    styleRanges: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.styleRanges ?? [])
      .map((styleRange) => ({
        range: styleRange.range,
        style: portableStyle(styleRange.styleId),
      }))
      .toSorted((left, right) => compareRangeSortKeys(left.range, right.range)),
    tables: (snapshot.workbook.metadata?.tables ?? [])
      .map(projectTableSemantics)
      .toSorted((left, right) => left.name.localeCompare(right.name)),
    charts: (snapshot.workbook.metadata?.charts ?? [])
      .map(projectChartSemantics)
      .toSorted((left, right) => left.id.localeCompare(right.id)),
    pivots: (snapshot.workbook.metadata?.pivots ?? [])
      .map(projectPivotSemantics)
      .toSorted((left, right) =>
        `${left.sheetName}:${left.address}:${left.name}`.localeCompare(`${right.sheetName}:${right.address}:${right.name}`),
      ),
    validations: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.validations ?? [])
      .map(projectValidationSemantics)
      .toSorted((left, right) => compareRangeSortKeys(left.range, right.range)),
    conditionalFormats: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.conditionalFormats ?? [])
      .map(projectConditionalFormatSemantics)
      .toSorted((left, right) => compareRangeSortKeys(left.range, right.range)),
    freezePanes: snapshot.sheets
      .flatMap((sheet) =>
        sheet.metadata?.freezePane ? [{ sheetName: sheet.name, freezePane: projectFreezePaneSemantics(sheet.metadata.freezePane) }] : [],
      )
      .toSorted((left, right) => left.sheetName.localeCompare(right.sheetName)),
    filters: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.filters ?? [])
      .map(({ sheetName, startAddress, endAddress }) => ({ sheetName, startAddress, endAddress }))
      .toSorted((left, right) => compareRangeSortKeys(left, right)),
    sorts: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.sorts ?? [])
      .map(projectSortSemantics)
      .toSorted((left, right) => compareRangeSortKeys(left.range, right.range)),
    sheetProtections: snapshot.sheets
      .flatMap((sheet) => (sheet.metadata?.sheetProtection ? [structuredClone(sheet.metadata.sheetProtection)] : []))
      .toSorted((left, right) => left.sheetName.localeCompare(right.sheetName)),
    protectedRanges: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.protectedRanges ?? [])
      .map(projectRangeProtectionSemantics)
      .toSorted((left, right) => `${rangeSortKey(left.range)}:${left.id}`.localeCompare(`${rangeSortKey(right.range)}:${right.id}`)),
    valueFormulaFormatSheets: snapshot.sheets
      .toSorted((left, right) => left.order - right.order)
      .map((sheet) => ({
        name: sheet.name,
        order: sheet.order,
        cells: sheet.cells
          .map((cell) => ({
            address: cell.address,
            ...(cell.value !== undefined ? { value: cell.value } : {}),
            ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
            ...(cell.format !== undefined ? { format: cell.format } : {}),
          }))
          .toSorted((left, right) => left.address.localeCompare(right.address)),
      })),
    dimensionSheets: snapshot.sheets
      .toSorted((left, right) => left.order - right.order)
      .map((sheet) => ({
        name: sheet.name,
        columns: (sheet.metadata?.columns ?? []).map(projectDimensionAxisEntry).toSorted((left, right) => left.index - right.index),
        rows: (sheet.metadata?.rows ?? []).map(projectDimensionAxisEntry).toSorted((left, right) => left.index - right.index),
        merges: (sheet.metadata?.merges ?? [])
          .map(({ sheetName, startAddress, endAddress }) => ({ sheetName, startAddress, endAddress }))
          .toSorted((left, right) => compareRangeSortKeys(left, right)),
      })),
  }
}

function projectChartSemantics(chart: WorkbookChartSnapshot): ProjectedWorkbookChartSemantics {
  const projected: ProjectedWorkbookChartSemantics = {
    id: chart.id,
    sheetName: chart.sheetName,
    address: chart.address,
    source: chart.source,
    chartType: chart.chartType,
    seriesOrientation: chart.seriesOrientation ?? 'columns',
    rows: chart.rows,
    cols: chart.cols,
  }
  if (chart.firstRowAsHeaders !== undefined) {
    return {
      ...projected,
      firstRowAsHeaders: chart.firstRowAsHeaders,
      ...(chart.firstColumnAsLabels !== undefined ? { firstColumnAsLabels: chart.firstColumnAsLabels } : {}),
      ...(chart.title !== undefined ? { title: chart.title } : {}),
      ...(chart.legendPosition !== undefined ? { legendPosition: chart.legendPosition } : {}),
    }
  }
  return {
    ...projected,
    ...(chart.firstColumnAsLabels !== undefined ? { firstColumnAsLabels: chart.firstColumnAsLabels } : {}),
    ...(chart.title !== undefined ? { title: chart.title } : {}),
    ...(chart.legendPosition !== undefined ? { legendPosition: chart.legendPosition } : {}),
  }
}

function projectPivotValue(value: WorkbookPivotValueSnapshot): WorkbookPivotValueSnapshot {
  const projected: WorkbookPivotValueSnapshot = {
    sourceColumn: value.sourceColumn,
    summarizeBy: value.summarizeBy,
  }
  if (value.outputLabel !== undefined) {
    projected.outputLabel = value.outputLabel
  }
  return projected
}

function projectPivotSemantics(pivot: WorkbookPivotSnapshot): WorkbookPivotSnapshot {
  return {
    name: pivot.name,
    sheetName: pivot.sheetName,
    address: pivot.address,
    ...(pivot.source ? { source: pivot.source } : {}),
    groupBy: [...pivot.groupBy],
    ...(pivot.columnFields ? { columnFields: [...pivot.columnFields] } : {}),
    ...(pivot.pageFields ? { pageFields: structuredClone(pivot.pageFields) } : {}),
    ...(pivot.filters ? { filters: structuredClone(pivot.filters) } : {}),
    ...(pivot.hiddenItems ? { hiddenItems: structuredClone(pivot.hiddenItems) } : {}),
    ...(pivot.calculatedFields ? { calculatedFields: structuredClone(pivot.calculatedFields) } : {}),
    ...(pivot.calculatedItems ? { calculatedItems: structuredClone(pivot.calculatedItems) } : {}),
    values: pivot.values.map(projectPivotValue),
    rows: pivot.rows,
    cols: pivot.cols,
  }
}

function projectTableSemantics(table: WorkbookTableSnapshot): WorkbookTableSnapshot {
  return {
    name: table.name,
    sheetName: table.sheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
    columnNames: [...table.columnNames],
    headerRow: table.headerRow,
    totalsRow: table.totalsRow,
  }
}

function projectDimensionAxisEntry(entry: { readonly index: number; readonly size?: number | null }): ProjectedWorkbookDimensionAxisEntry {
  return {
    index: entry.index,
    ...(entry.size !== undefined ? { size: entry.size } : {}),
  }
}

function projectValidationSemantics(validation: WorkbookDataValidationSnapshot): WorkbookDataValidationSnapshot {
  return structuredClone(validation)
}

function projectRangeProtectionSemantics(protection: WorkbookRangeProtectionSnapshot): WorkbookRangeProtectionSnapshot {
  return structuredClone(protection)
}

function projectSortSemantics(sort: WorkbookSortSnapshot): WorkbookSortSnapshot {
  return structuredClone(sort)
}

function projectConditionalFormatSemantics(format: WorkbookConditionalFormatSnapshot) {
  return {
    range: format.range,
    rule: structuredClone(format.rule),
    style: structuredClone(format.style),
    ...(format.stopIfTrue !== undefined ? { stopIfTrue: format.stopIfTrue } : {}),
    ...(format.priority !== undefined ? { priority: format.priority } : {}),
  }
}

function projectFreezePaneSemantics(freezePane: WorkbookFreezePaneSnapshot): WorkbookFreezePaneSnapshot {
  return {
    rows: freezePane.rows,
    cols: freezePane.cols,
    ...(freezePane.topLeftCell !== undefined && freezePane.topLeftCell !== defaultFreezePaneTopLeftCell(freezePane)
      ? { topLeftCell: freezePane.topLeftCell }
      : {}),
    ...(freezePane.activePane !== undefined && freezePane.activePane !== defaultFreezePaneActivePane(freezePane)
      ? { activePane: freezePane.activePane }
      : {}),
  }
}

function defaultFreezePaneActivePane(freezePane: WorkbookFreezePaneSnapshot): string {
  if (freezePane.rows > 0 && freezePane.cols > 0) {
    return 'bottomRight'
  }
  return freezePane.rows > 0 ? 'bottomLeft' : 'topRight'
}

function defaultFreezePaneTopLeftCell(freezePane: WorkbookFreezePaneSnapshot): string {
  return formatAddress(freezePane.rows, freezePane.cols)
}

function compareRangeRefs(left: ComparableRangeRef, right: ComparableRangeRef): number {
  const leftStart = parseCellAddress(left.startAddress, left.sheetName)
  const leftEnd = parseCellAddress(left.endAddress, left.sheetName)
  const rightStart = parseCellAddress(right.startAddress, right.sheetName)
  const rightEnd = parseCellAddress(right.endAddress, right.sheetName)
  return (
    left.sheetName.localeCompare(right.sheetName) ||
    leftStart.row - rightStart.row ||
    leftStart.col - rightStart.col ||
    leftEnd.row - rightEnd.row ||
    leftEnd.col - rightEnd.col
  )
}

function compareRangeSortKeys(left: ComparableRangeRef, right: ComparableRangeRef): number {
  return rangeSortKey(left).localeCompare(rangeSortKey(right))
}

function rangeSortKey(range: ComparableRangeRef): string {
  return `${range.sheetName}:${range.startAddress}:${range.endAddress}`
}

function normalizeRangeRecords<TRecord extends { range: ComparableRangeRef } & Record<TKey, string>, TKey extends keyof TRecord & string>(
  records: readonly TRecord[],
  idKey: TKey,
): TRecord[] {
  const groups = new Map<
    string,
    {
      prototype: TRecord
      readonly id: string
      readonly sheetName: string
      readonly colsByRow: Map<number, Set<number>>
    }
  >()
  for (const record of records) {
    const id = record[idKey]
    const sheetName = record.range.sheetName
    const groupKey = `${id}\u0000${sheetName}`
    let group = groups.get(groupKey)
    if (!group) {
      group = {
        prototype: record,
        id,
        sheetName,
        colsByRow: new Map(),
      }
      groups.set(groupKey, group)
    }
    const start = parseCellAddress(record.range.startAddress, sheetName)
    const end = parseCellAddress(record.range.endAddress, sheetName)
    const rowStart = Math.min(start.row, end.row)
    const rowEnd = Math.max(start.row, end.row)
    const colStart = Math.min(start.col, end.col)
    const colEnd = Math.max(start.col, end.col)
    for (let row = rowStart; row <= rowEnd; row += 1) {
      let cols = group.colsByRow.get(row)
      if (!cols) {
        cols = new Set()
        group.colsByRow.set(row, cols)
      }
      for (let col = colStart; col <= colEnd; col += 1) {
        cols.add(col)
      }
    }
  }

  return [...groups.values()]
    .toSorted((left, right) => left.id.localeCompare(right.id) || left.sheetName.localeCompare(right.sheetName))
    .flatMap((group) => {
      const normalized: TRecord[] = []
      const rows = [...group.colsByRow.entries()].toSorted(([left], [right]) => left - right)
      for (const [row, cols] of rows) {
        const sortedCols = [...cols].toSorted((left, right) => left - right)
        let runStart: number | undefined
        let runEnd: number | undefined
        const pushRun = () => {
          if (runStart === undefined || runEnd === undefined) {
            return
          }
          normalized.push({
            ...group.prototype,
            range: {
              sheetName: group.sheetName,
              startAddress: formatAddress(row, runStart),
              endAddress: formatAddress(row, runEnd),
            },
          })
        }
        for (const col of sortedCols) {
          if (runStart === undefined || runEnd === undefined) {
            runStart = col
            runEnd = col
            continue
          }
          if (col === runEnd + 1) {
            runEnd = col
            continue
          }
          pushRun()
          runStart = col
          runEnd = col
        }
        pushRun()
      }
      return normalized
    })
}
