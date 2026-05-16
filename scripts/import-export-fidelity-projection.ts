import type {
  WorkbookChartSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookPivotSnapshot,
  WorkbookPivotValueSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookSortSnapshot,
  WorkbookSnapshot,
  WorkbookTableSnapshot,
} from '../packages/protocol/src/types.js'

interface ProjectedChartSemantics {
  id: string
  sheetName: string
  address: string
  source: WorkbookChartSnapshot['source']
  chartType: WorkbookChartSnapshot['chartType']
  seriesOrientation?: WorkbookChartSnapshot['seriesOrientation']
  firstRowAsHeaders?: WorkbookChartSnapshot['firstRowAsHeaders']
  firstColumnAsLabels?: WorkbookChartSnapshot['firstColumnAsLabels']
  title?: WorkbookChartSnapshot['title']
  legendPosition?: WorkbookChartSnapshot['legendPosition']
  rows: number
  cols: number
}

function projectChartSemantics(chart: WorkbookChartSnapshot): ProjectedChartSemantics {
  const projected: ProjectedChartSemantics = {
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
    projected.firstRowAsHeaders = chart.firstRowAsHeaders
  }
  if (chart.firstColumnAsLabels !== undefined) {
    projected.firstColumnAsLabels = chart.firstColumnAsLabels
  }
  if (chart.title !== undefined) {
    projected.title = chart.title
  }
  if (chart.legendPosition !== undefined) {
    projected.legendPosition = chart.legendPosition
  }
  return projected
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
    source: pivot.source,
    groupBy: [...pivot.groupBy],
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

function projectValidationSemantics(validation: WorkbookDataValidationSnapshot): WorkbookDataValidationSnapshot {
  return structuredClone(validation)
}

function projectRangeProtectionSemantics(protection: WorkbookRangeProtectionSnapshot): WorkbookRangeProtectionSnapshot {
  return structuredClone(protection)
}

function projectSortSemantics(sort: WorkbookSortSnapshot): WorkbookSortSnapshot {
  return structuredClone(sort)
}

function defaultFreezePaneActivePane(freezePane: WorkbookFreezePaneSnapshot): string {
  if (freezePane.rows > 0 && freezePane.cols > 0) {
    return 'bottomRight'
  }
  return freezePane.rows > 0 ? 'bottomLeft' : 'topRight'
}

function defaultFreezePaneTopLeftCell(freezePane: WorkbookFreezePaneSnapshot): string {
  let column = ''
  let index = freezePane.cols
  do {
    const remainder = index % 26
    column = String.fromCharCode(65 + remainder) + column
    index = Math.floor(index / 26) - 1
  } while (index >= 0)
  return `${column}${String(freezePane.rows + 1)}`
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

function projectConditionalFormatSemantics(format: WorkbookConditionalFormatSnapshot) {
  return {
    range: format.range,
    rule: structuredClone(format.rule),
    style: structuredClone(format.style),
    ...(format.stopIfTrue !== undefined ? { stopIfTrue: format.stopIfTrue } : {}),
    ...(format.priority !== undefined ? { priority: format.priority } : {}),
  }
}

export function projectSupportedSnapshotSemantics(snapshot: WorkbookSnapshot) {
  const stylesById = new Map((snapshot.workbook.metadata?.styles ?? []).map((style) => [style.id, style]))
  const portableStyle = (styleId: string) => {
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
      .toSorted((left, right) =>
        `${left.range.sheetName}:${left.range.startAddress}:${left.range.endAddress}`.localeCompare(
          `${right.range.sheetName}:${right.range.startAddress}:${right.range.endAddress}`,
        ),
      ),
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
      .toSorted((left, right) =>
        `${left.range.sheetName}:${left.range.startAddress}:${left.range.endAddress}`.localeCompare(
          `${right.range.sheetName}:${right.range.startAddress}:${right.range.endAddress}`,
        ),
      ),
    conditionalFormats: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.conditionalFormats ?? [])
      .map(projectConditionalFormatSemantics)
      .toSorted((left, right) =>
        `${left.range.sheetName}:${left.range.startAddress}:${left.range.endAddress}`.localeCompare(
          `${right.range.sheetName}:${right.range.startAddress}:${right.range.endAddress}`,
        ),
      ),
    freezePanes: snapshot.sheets
      .flatMap((sheet) =>
        sheet.metadata?.freezePane ? [{ sheetName: sheet.name, freezePane: projectFreezePaneSemantics(sheet.metadata.freezePane) }] : [],
      )
      .toSorted((left, right) => left.sheetName.localeCompare(right.sheetName)),
    filters: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.filters ?? [])
      .map(({ sheetName, startAddress, endAddress }) => ({ sheetName, startAddress, endAddress }))
      .toSorted((left, right) =>
        `${left.sheetName}:${left.startAddress}:${left.endAddress}`.localeCompare(
          `${right.sheetName}:${right.startAddress}:${right.endAddress}`,
        ),
      ),
    sorts: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.sorts ?? [])
      .map(projectSortSemantics)
      .toSorted((left, right) =>
        `${left.range.sheetName}:${left.range.startAddress}:${left.range.endAddress}`.localeCompare(
          `${right.range.sheetName}:${right.range.startAddress}:${right.range.endAddress}`,
        ),
      ),
    sheetProtections: snapshot.sheets
      .flatMap((sheet) => (sheet.metadata?.sheetProtection ? [structuredClone(sheet.metadata.sheetProtection)] : []))
      .toSorted((left, right) => left.sheetName.localeCompare(right.sheetName)),
    protectedRanges: snapshot.sheets
      .flatMap((sheet) => sheet.metadata?.protectedRanges ?? [])
      .map(projectRangeProtectionSemantics)
      .toSorted((left, right) =>
        `${left.range.sheetName}:${left.range.startAddress}:${left.range.endAddress}:${left.id}`.localeCompare(
          `${right.range.sheetName}:${right.range.startAddress}:${right.range.endAddress}:${right.id}`,
        ),
      ),
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
        columns: (sheet.metadata?.columns ?? [])
          .map(({ index, size }) => ({ index, size }))
          .toSorted((left, right) => left.index - right.index),
        rows: (sheet.metadata?.rows ?? []).map(({ index, size }) => ({ index, size })).toSorted((left, right) => left.index - right.index),
        merges: (sheet.metadata?.merges ?? [])
          .map(({ sheetName, startAddress, endAddress }) => ({ sheetName, startAddress, endAddress }))
          .toSorted((left, right) =>
            `${left.sheetName}:${left.startAddress}:${left.endAddress}`.localeCompare(
              `${right.sheetName}:${right.startAddress}:${right.endAddress}`,
            ),
          ),
      })),
  }
}
