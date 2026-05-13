import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type {
  CellRangeRef,
  WorkbookChartLegendPosition,
  WorkbookChartSeriesOrientation,
  WorkbookChartSnapshot,
  WorkbookChartType,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
}

interface ChartSeriesRefs {
  readonly name?: CellRangeRef
  readonly category?: CellRangeRef
  readonly value?: CellRangeRef
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const worksheetDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const chartRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const drawingContentType = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const chartContentType = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function stringChild(value: unknown, key: string): string | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return typeof child === 'string' ? child : null
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function getZipText(zip: ZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}

function setZipText(zip: ZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function nextPartIndex(zip: ZipEntries, prefix: string, suffix: string): number {
  let next = 1
  for (const path of Object.keys(zip)) {
    if (!path.startsWith(prefix) || !path.endsWith(suffix)) {
      continue
    }
    const raw = path.slice(prefix.length, -suffix.length)
    const value = Number(raw)
    if (Number.isInteger(value) && value >= next) {
      next = value + 1
    }
  }
  return next
}

function absoluteAddress(address: string): string {
  const decoded = XLSX.utils.decode_cell(address)
  return `$${XLSX.utils.encode_col(decoded.c)}$${decoded.r + 1}`
}

function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replaceAll("'", "''")}'`
}

function formulaForRange(sheetName: string, range: CellRangeRef): string {
  const start = absoluteAddress(range.startAddress)
  const end = absoluteAddress(range.endAddress)
  return start === end ? `${quoteSheetName(sheetName)}!${start}` : `${quoteSheetName(sheetName)}!${start}:${end}`
}

function rangeFromIndexes(sheetName: string, startRow: number, startCol: number, endRow: number, endCol: number): CellRangeRef {
  return {
    sheetName,
    startAddress: XLSX.utils.encode_cell({ r: startRow, c: startCol }),
    endAddress: XLSX.utils.encode_cell({ r: endRow, c: endCol }),
  }
}

function chartTypeElement(chartType: WorkbookChartType): string {
  switch (chartType) {
    case 'bar':
    case 'column':
      return 'barChart'
    case 'area':
      return 'areaChart'
    case 'pie':
      return 'pieChart'
    case 'scatter':
      return 'scatterChart'
    case 'line':
    default:
      return 'lineChart'
  }
}

function legendPositionValue(position: WorkbookChartLegendPosition | undefined): string | null {
  switch (position) {
    case 'top':
      return 't'
    case 'right':
      return 'r'
    case 'bottom':
      return 'b'
    case 'left':
      return 'l'
    case 'hidden':
      return null
    case undefined:
      return null
    default:
      return 'r'
  }
}

function parseLegendPosition(value: unknown): WorkbookChartLegendPosition | undefined {
  switch (value) {
    case 't':
      return 'top'
    case 'r':
      return 'right'
    case 'b':
      return 'bottom'
    case 'l':
      return 'left'
    default:
      return undefined
  }
}

function buildChartSeries(chart: WorkbookChartSnapshot, exportSourceSheetName: string): ChartSeriesRefs[] {
  const source = XLSX.utils.decode_range(`${chart.source.startAddress}:${chart.source.endAddress}`)
  const orientation = chart.seriesOrientation ?? 'columns'
  const firstRowAsHeaders = chart.firstRowAsHeaders === true
  const firstColumnAsLabels = chart.firstColumnAsLabels === true
  const series: ChartSeriesRefs[] = []

  if (orientation === 'rows') {
    const dataStartRow = source.s.r + (firstRowAsHeaders ? 1 : 0)
    const dataStartCol = source.s.c + (firstColumnAsLabels ? 1 : 0)
    const category =
      firstRowAsHeaders && dataStartCol <= source.e.c
        ? rangeFromIndexes(exportSourceSheetName, source.s.r, dataStartCol, source.s.r, source.e.c)
        : undefined
    for (let row = dataStartRow; row <= source.e.r; row += 1) {
      series.push({
        ...(firstColumnAsLabels ? { name: rangeFromIndexes(exportSourceSheetName, row, source.s.c, row, source.s.c) } : {}),
        ...(category ? { category } : {}),
        value: rangeFromIndexes(exportSourceSheetName, row, dataStartCol, row, source.e.c),
      })
    }
    return series.length > 0 ? series : [{ value: { ...chart.source, sheetName: exportSourceSheetName } }]
  }

  const dataStartRow = source.s.r + (firstRowAsHeaders ? 1 : 0)
  const dataStartCol = source.s.c + (firstColumnAsLabels ? 1 : 0)
  const category =
    firstColumnAsLabels && dataStartRow <= source.e.r
      ? rangeFromIndexes(exportSourceSheetName, dataStartRow, source.s.c, source.e.r, source.s.c)
      : undefined
  for (let col = dataStartCol; col <= source.e.c; col += 1) {
    series.push({
      ...(firstRowAsHeaders ? { name: rangeFromIndexes(exportSourceSheetName, source.s.r, col, source.s.r, col) } : {}),
      ...(category ? { category } : {}),
      value: rangeFromIndexes(exportSourceSheetName, dataStartRow, col, source.e.r, col),
    })
  }
  return series.length > 0 ? series : [{ value: { ...chart.source, sheetName: exportSourceSheetName } }]
}

function chartRefXml(kind: 'cat' | 'name' | 'val' | 'xVal' | 'yVal', formula: string): string {
  if (kind === 'name') {
    return `<c:tx><c:strRef><c:f>${escapeXml(formula)}</c:f></c:strRef></c:tx>`
  }
  const refTag = kind === 'cat' ? 'strRef' : 'numRef'
  return `<c:${kind}><c:${refTag}><c:f>${escapeXml(formula)}</c:f></c:${refTag}></c:${kind}>`
}

function buildSeriesXml(chart: WorkbookChartSnapshot, exportSourceSheetName: string): string {
  return buildChartSeries(chart, exportSourceSheetName)
    .map((series, index) => {
      const name = series.name ? chartRefXml('name', formulaForRange(exportSourceSheetName, series.name)) : ''
      if (chart.chartType === 'scatter') {
        const xValue = series.category ?? series.value
        if (!xValue || !series.value) {
          return ''
        }
        return [
          `<c:ser><c:idx val="${String(index)}"/><c:order val="${String(index)}"/>`,
          name,
          chartRefXml('xVal', formulaForRange(exportSourceSheetName, xValue)),
          series.value ? chartRefXml('yVal', formulaForRange(exportSourceSheetName, series.value)) : '',
          '</c:ser>',
        ].join('')
      }
      return [
        `<c:ser><c:idx val="${String(index)}"/><c:order val="${String(index)}"/>`,
        name,
        series.category ? chartRefXml('cat', formulaForRange(exportSourceSheetName, series.category)) : '',
        series.value ? chartRefXml('val', formulaForRange(exportSourceSheetName, series.value)) : '',
        '</c:ser>',
      ].join('')
    })
    .join('')
}

function buildChartTitleXml(title: string | undefined): string {
  if (!title || title.trim().length === 0) {
    return ''
  }
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(
    title.trim(),
  )}</a:t></a:r></a:p></c:rich></c:tx></c:title>`
}

function buildLegendXml(position: WorkbookChartLegendPosition | undefined): string {
  const value = legendPositionValue(position)
  return value ? `<c:legend><c:legendPos val="${value}"/><c:layout/></c:legend>` : ''
}

function buildPlotAreaXml(chart: WorkbookChartSnapshot, exportSourceSheetName: string): string {
  const typeElement = chartTypeElement(chart.chartType)
  const barDirection =
    chart.chartType === 'bar' || chart.chartType === 'column' ? `<c:barDir val="${chart.chartType === 'bar' ? 'bar' : 'col'}"/>` : ''
  const grouping =
    chart.chartType === 'bar' || chart.chartType === 'column' || chart.chartType === 'line' || chart.chartType === 'area'
      ? '<c:grouping val="standard"/>'
      : ''
  const axisIds = chart.chartType === 'pie' ? '' : '<c:axId val="123456"/><c:axId val="123457"/>'
  const axes =
    chart.chartType === 'pie'
      ? ''
      : [
          '<c:catAx><c:axId val="123456"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/>',
          '<c:crossAx val="123457"/><c:tickLblPos val="nextTo"/></c:catAx>',
          '<c:valAx><c:axId val="123457"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/>',
          '<c:crossAx val="123456"/><c:tickLblPos val="nextTo"/></c:valAx>',
        ].join('')
  return [
    '<c:plotArea><c:layout/>',
    `<c:${typeElement}>`,
    barDirection,
    grouping,
    '<c:varyColors val="0"/>',
    buildSeriesXml(chart, exportSourceSheetName),
    axisIds,
    `</c:${typeElement}>`,
    axes,
    '</c:plotArea>',
  ].join('')
}

function buildChartXml(chart: WorkbookChartSnapshot, exportSourceSheetName: string): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<c:chart>',
    buildChartTitleXml(chart.title),
    buildPlotAreaXml(chart, exportSourceSheetName),
    buildLegendXml(chart.legendPosition),
    '<c:plotVisOnly val="1"/>',
    '</c:chart>',
    '</c:chartSpace>',
  ].join('')
}

function buildDrawingAnchorXml(chart: WorkbookChartSnapshot, relationshipId: string, anchorId: number): string {
  const decoded = XLSX.utils.decode_cell(chart.address)
  const endCol = decoded.c + Math.max(1, chart.cols)
  const endRow = decoded.r + Math.max(1, chart.rows)
  return [
    '<xdr:twoCellAnchor editAs="twoCell">',
    `<xdr:from><xdr:col>${String(decoded.c)}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${String(decoded.r)}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`,
    `<xdr:to><xdr:col>${String(endCol)}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${String(endRow)}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`,
    '<xdr:graphicFrame macro="">',
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="${String(anchorId)}" name="${escapeXml(
      chart.id,
    )}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`,
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>',
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">',
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relationshipId}"/>`,
    '</a:graphicData></a:graphic>',
    '</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>',
  ].join('')
}

function buildDrawingXml(anchors: readonly string[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    ...anchors,
    '</xdr:wsDr>',
  ].join('')
}

function buildRelationshipsXml(relationships: readonly ParsedRelationship[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    ...relationships.map(
      (relationship) =>
        `<Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(
          relationship.target,
        )}"/>`,
    ),
    '</Relationships>',
  ].join('')
}

function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(xml)
  return asArray(recordChild(parsed, 'Relationships')?.['Relationship']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['Id'] !== 'string' || typeof entry['Target'] !== 'string' || typeof entry['Type'] !== 'string') {
      return []
    }
    return [{ id: entry['Id'], target: entry['Target'], type: entry['Type'] }]
  })
}

function nextRelationshipId(relationships: readonly ParsedRelationship[]): string {
  let next = 1
  for (const relationship of relationships) {
    const match = /^rId(\d+)$/u.exec(relationship.id)
    if (match) {
      next = Math.max(next, Number(match[1]) + 1)
    }
  }
  return `rId${String(next)}`
}

function ensureWorksheetRelationshipNamespace(sheetXml: string): string {
  if (/xmlns:r=/u.test(sheetXml)) {
    return sheetXml
  }
  return sheetXml.replace(
    /<worksheet\b([^>]*)>/u,
    `<worksheet$1 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`,
  )
}

function addWorksheetDrawing(sheetXml: string, relationshipId: string): string {
  const withNamespace = ensureWorksheetRelationshipNamespace(sheetXml)
  if (/<drawing\b/u.test(withNamespace)) {
    return withNamespace.replace(/<drawing\b[^>]*\/>/u, `<drawing r:id="${relationshipId}"/>`)
  }
  return withNamespace.replace('</worksheet>', `<drawing r:id="${relationshipId}"/></worksheet>`)
}

function addContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (contentTypesXml.includes(`PartName="${partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`)
}

function resolveTargetPath(basePartPath: string, target: string): string {
  const parts = basePartPath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return parts.join('/')
}

export function addExportChartsToXlsxBytes(
  bytes: Uint8Array,
  snapshot: WorkbookSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): Uint8Array {
  const charts = snapshot.workbook.metadata?.charts ?? []
  if (charts.length === 0) {
    return bytes
  }
  const zip = unzipSync(bytes)
  let nextDrawingIndex = nextPartIndex(zip, 'xl/drawings/drawing', '.xml')
  let nextChartIndex = nextPartIndex(zip, 'xl/charts/chart', '.xml')
  let contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetCharts = charts.filter((chart) => chart.sheetName === sheet.name)
      if (sheetCharts.length === 0) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }

      const drawingIndex = nextDrawingIndex
      nextDrawingIndex += 1
      const drawingPath = `xl/drawings/drawing${String(drawingIndex)}.xml`
      const drawingRelsPath = `xl/drawings/_rels/drawing${String(drawingIndex)}.xml.rels`
      const anchors: string[] = []
      const drawingRelationships: ParsedRelationship[] = []

      sheetCharts.forEach((chart, chartIndex) => {
        const chartPartIndex = nextChartIndex
        nextChartIndex += 1
        const chartPath = `xl/charts/chart${String(chartPartIndex)}.xml`
        const relationshipId = `rId${String(chartIndex + 1)}`
        const exportSourceSheetName = exportSheetNamesByOriginalName.get(chart.source.sheetName) ?? chart.source.sheetName
        setZipText(zip, chartPath, buildChartXml(chart, exportSourceSheetName))
        anchors.push(buildDrawingAnchorXml(chart, relationshipId, chartIndex + 2))
        drawingRelationships.push({
          id: relationshipId,
          type: chartRelationshipType,
          target: `../charts/chart${String(chartPartIndex)}.xml`,
        })
        contentTypesXml = addContentTypeOverride(contentTypesXml, `/${chartPath}`, chartContentType)
      })

      setZipText(zip, drawingPath, buildDrawingXml(anchors))
      setZipText(zip, drawingRelsPath, buildRelationshipsXml(drawingRelationships))
      contentTypesXml = addContentTypeOverride(contentTypesXml, `/${drawingPath}`, drawingContentType)

      const sheetRelsPath = `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
      const sheetRelationships = parseRelationships(getZipText(zip, sheetRelsPath))
      const drawingRelationshipId = nextRelationshipId(sheetRelationships)
      sheetRelationships.push({
        id: drawingRelationshipId,
        type: worksheetDrawingRelationshipType,
        target: `../drawings/drawing${String(drawingIndex)}.xml`,
      })
      setZipText(zip, sheetRelsPath, buildRelationshipsXml(sheetRelationships))
      setZipText(zip, sheetPath, addWorksheetDrawing(sheetXml, drawingRelationshipId))
    })

  if (contentTypesXml.length > 0) {
    setZipText(zip, '[Content_Types].xml', contentTypesXml)
  }
  return zipSync(zip)
}

function parseSheetRangeFormula(formula: string): CellRangeRef | null {
  const match = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$/u.exec(formula.trim())
  if (!match) {
    return null
  }
  const sheetName = (match[1] ?? match[2] ?? '').replaceAll("''", "'")
  const startAddress = `${match[3]}${match[4]}`
  const endAddress = `${match[5] ?? match[3]}${match[6] ?? match[4]}`
  return { sheetName, startAddress, endAddress }
}

function readReferenceFormula(value: unknown): CellRangeRef | null {
  if (!isRecord(value)) {
    return null
  }
  const numRef = recordChild(value, 'numRef')
  const strRef = recordChild(value, 'strRef')
  return parseSheetRangeFormula(stringChild(numRef, 'f') ?? stringChild(strRef, 'f') ?? '')
}

function readSeriesRefs(series: unknown): ChartSeriesRefs | null {
  if (!isRecord(series)) {
    return null
  }
  const name = readReferenceFormula(recordChild(series, 'tx'))
  const category = readReferenceFormula(recordChild(series, 'cat')) ?? readReferenceFormula(recordChild(series, 'xVal'))
  const value = readReferenceFormula(recordChild(series, 'val')) ?? readReferenceFormula(recordChild(series, 'yVal'))
  return value ? { ...(name ? { name } : {}), ...(category ? { category } : {}), value } : null
}

function unionRanges(ranges: readonly CellRangeRef[]): CellRangeRef | null {
  const [first] = ranges
  if (!first) {
    return null
  }
  let start = XLSX.utils.decode_cell(first.startAddress)
  let end = XLSX.utils.decode_cell(first.endAddress)
  for (const range of ranges.slice(1)) {
    if (range.sheetName !== first.sheetName) {
      continue
    }
    const nextStart = XLSX.utils.decode_cell(range.startAddress)
    const nextEnd = XLSX.utils.decode_cell(range.endAddress)
    start = { r: Math.min(start.r, nextStart.r), c: Math.min(start.c, nextStart.c) }
    end = { r: Math.max(end.r, nextEnd.r), c: Math.max(end.c, nextEnd.c) }
  }
  return {
    sheetName: first.sheetName,
    startAddress: XLSX.utils.encode_cell(start),
    endAddress: XLSX.utils.encode_cell(end),
  }
}

function inferSeriesOrientation(series: readonly ChartSeriesRefs[]): WorkbookChartSeriesOrientation | undefined {
  const valueRanges = series.flatMap((entry) => (entry.value ? [entry.value] : []))
  if (valueRanges.length === 0) {
    return undefined
  }
  const decoded = valueRanges.map((range) => ({
    start: XLSX.utils.decode_cell(range.startAddress),
    end: XLSX.utils.decode_cell(range.endAddress),
  }))
  const columns = decoded.every((range) => range.start.c === range.end.c)
  const rows = decoded.every((range) => range.start.r === range.end.r)
  if (columns) {
    return 'columns'
  }
  if (rows) {
    return 'rows'
  }
  return undefined
}

function rangeWithinSource(source: CellRangeRef, range: CellRangeRef): boolean {
  const sourceStart = XLSX.utils.decode_cell(source.startAddress)
  const sourceEnd = XLSX.utils.decode_cell(source.endAddress)
  const rangeStart = XLSX.utils.decode_cell(range.startAddress)
  const rangeEnd = XLSX.utils.decode_cell(range.endAddress)
  return (
    range.sheetName === source.sheetName &&
    rangeStart.r >= sourceStart.r &&
    rangeEnd.r <= sourceEnd.r &&
    rangeStart.c >= sourceStart.c &&
    rangeEnd.c <= sourceEnd.c
  )
}

function isTopRowHeaderRange(source: CellRangeRef, range: CellRangeRef): boolean {
  if (!rangeWithinSource(source, range)) {
    return false
  }
  const sourceStart = XLSX.utils.decode_cell(source.startAddress)
  const sourceEnd = XLSX.utils.decode_cell(source.endAddress)
  const rangeStart = XLSX.utils.decode_cell(range.startAddress)
  const rangeEnd = XLSX.utils.decode_cell(range.endAddress)
  return sourceStart.r < sourceEnd.r && rangeStart.r === sourceStart.r && rangeEnd.r === sourceStart.r
}

function isFirstColumnLabelRange(source: CellRangeRef, range: CellRangeRef): boolean {
  if (!rangeWithinSource(source, range)) {
    return false
  }
  const sourceStart = XLSX.utils.decode_cell(source.startAddress)
  const sourceEnd = XLSX.utils.decode_cell(source.endAddress)
  const rangeStart = XLSX.utils.decode_cell(range.startAddress)
  const rangeEnd = XLSX.utils.decode_cell(range.endAddress)
  return sourceStart.c < sourceEnd.c && rangeStart.c === sourceStart.c && rangeEnd.c === sourceStart.c
}

function rangeEquals(left: CellRangeRef | undefined, right: CellRangeRef | undefined): boolean {
  return (
    left !== undefined &&
    right !== undefined &&
    left.sheetName === right.sheetName &&
    left.startAddress === right.startAddress &&
    left.endAddress === right.endAddress
  )
}

function nonValueSeriesRange(range: CellRangeRef | undefined, series: ChartSeriesRefs): CellRangeRef[] {
  return range && !rangeEquals(range, series.value) ? [range] : []
}

function inferFirstRowAsHeaders(
  source: CellRangeRef,
  series: readonly ChartSeriesRefs[],
  orientation: WorkbookChartSeriesOrientation | undefined,
): boolean | undefined {
  const candidates = series.flatMap((entry) => {
    if (orientation === 'rows') {
      return nonValueSeriesRange(entry.category, entry)
    }
    if (orientation === 'columns') {
      return nonValueSeriesRange(entry.name, entry)
    }
    return [...nonValueSeriesRange(entry.name, entry), ...nonValueSeriesRange(entry.category, entry)]
  })
  return candidates.some((range) => isTopRowHeaderRange(source, range)) ? true : undefined
}

function inferFirstColumnAsLabels(
  source: CellRangeRef,
  series: readonly ChartSeriesRefs[],
  orientation: WorkbookChartSeriesOrientation | undefined,
): boolean | undefined {
  const candidates = series.flatMap((entry) => {
    if (orientation === 'rows') {
      return nonValueSeriesRange(entry.name, entry)
    }
    if (orientation === 'columns') {
      return nonValueSeriesRange(entry.category, entry)
    }
    return [...nonValueSeriesRange(entry.name, entry), ...nonValueSeriesRange(entry.category, entry)]
  })
  return candidates.some((range) => isFirstColumnLabelRange(source, range)) ? true : undefined
}

function textValues(value: unknown): string[] {
  if (typeof value === 'string') {
    return [value]
  }
  if (!isRecord(value)) {
    return []
  }
  return Object.entries(value).flatMap(([key, child]) => (key === 't' ? textValues(child) : textValues(child)))
}

function readChartTitle(chart: Record<string, unknown>): string | undefined {
  const values = textValues(recordChild(chart, 'title'))
  const title = values.join('').trim()
  return title.length > 0 ? title : undefined
}

function chartRecord(plotArea: Record<string, unknown>): { type: WorkbookChartType; record: Record<string, unknown> } | null {
  const lineChart = recordChild(plotArea, 'lineChart')
  if (lineChart) {
    return { type: 'line', record: lineChart }
  }
  const areaChart = recordChild(plotArea, 'areaChart')
  if (areaChart) {
    return { type: 'area', record: areaChart }
  }
  const pieChart = recordChild(plotArea, 'pieChart')
  if (pieChart) {
    return { type: 'pie', record: pieChart }
  }
  const scatterChart = recordChild(plotArea, 'scatterChart')
  if (scatterChart) {
    return { type: 'scatter', record: scatterChart }
  }
  const barChart = recordChild(plotArea, 'barChart')
  if (barChart) {
    return { type: recordChild(barChart, 'barDir')?.['val'] === 'bar' ? 'bar' : 'column', record: barChart }
  }
  return null
}

function parseChartXml(chartXml: string): {
  readonly chartType: WorkbookChartType
  readonly source: CellRangeRef
  readonly seriesOrientation?: WorkbookChartSeriesOrientation
  readonly firstRowAsHeaders?: boolean
  readonly firstColumnAsLabels?: boolean
  readonly title?: string
  readonly legendPosition?: WorkbookChartLegendPosition
} | null {
  const parsed: unknown = xmlParser.parse(chartXml)
  const chart = recordChild(recordChild(parsed, 'chartSpace'), 'chart')
  const plotArea = recordChild(chart, 'plotArea')
  if (!chart || !plotArea) {
    return null
  }
  const typedChart = chartRecord(plotArea)
  if (!typedChart) {
    return null
  }
  const series = asArray(typedChart.record['ser']).flatMap((entry) => {
    const refs = readSeriesRefs(entry)
    return refs ? [refs] : []
  })
  const source = unionRanges(
    series.flatMap((entry) => [entry.name, entry.category, entry.value].filter((range): range is CellRangeRef => Boolean(range))),
  )
  if (!source) {
    return null
  }
  const seriesOrientation = inferSeriesOrientation(series)
  const title = readChartTitle(chart)
  const legendPosition = parseLegendPosition(recordChild(recordChild(chart, 'legend'), 'legendPos')?.['val'])
  return {
    chartType: typedChart.type,
    source,
    ...(seriesOrientation !== undefined ? { seriesOrientation } : {}),
    ...(inferFirstRowAsHeaders(source, series, seriesOrientation) ? { firstRowAsHeaders: true } : {}),
    ...(inferFirstColumnAsLabels(source, series, seriesOrientation) ? { firstColumnAsLabels: true } : {}),
    ...(title !== undefined ? { title } : {}),
    ...(legendPosition !== undefined ? { legendPosition } : {}),
  }
}

function readDrawingRelationshipId(anchor: unknown): string | null {
  const chart = recordChild(recordChild(recordChild(recordChild(anchor, 'graphicFrame'), 'graphic'), 'graphicData'), 'chart')
  return typeof chart?.['id'] === 'string' ? chart['id'] : null
}

function readAnchorChartId(anchor: unknown, fallback: string): string {
  return stringChild(recordChild(recordChild(anchor, 'graphicFrame'), 'nvGraphicFramePr')?.['cNvPr'], 'name') ?? fallback
}

function readAnchorNumber(anchor: unknown, key: 'from' | 'to', field: 'col' | 'row'): number | null {
  const raw = recordChild(anchor, key)?.[field]
  const number = typeof raw === 'number' ? raw : typeof raw === 'string' ? Number(raw) : Number.NaN
  return Number.isFinite(number) ? number : null
}

export function readImportedWorkbookCharts(source: XlsxZipSource, sheetNames: readonly string[]): WorkbookChartSnapshot[] | undefined {
  const zip = readXlsxZipEntries(source)
  const charts: WorkbookChartSnapshot[] = []

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const sheetXml = getZipText(zip, sheetPath)
    const drawingRelationshipId = /<drawing\b[^>]*\br:id="([^"]+)"/u.exec(sheetXml ?? '')?.[1]
    if (!drawingRelationshipId) {
      return
    }
    const sheetRelationships = parseRelationships(getZipText(zip, `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`))
    const drawingRelationship = sheetRelationships.find(
      (relationship) => relationship.id === drawingRelationshipId && relationship.type === worksheetDrawingRelationshipType,
    )
    if (!drawingRelationship) {
      return
    }
    const drawingPath = resolveTargetPath(sheetPath, drawingRelationship.target)
    const drawingXml = getZipText(zip, drawingPath)
    if (!drawingXml) {
      return
    }
    const drawingRelationships = parseRelationships(
      getZipText(
        zip,
        `${drawingPath.slice(0, drawingPath.lastIndexOf('/'))}/_rels/${drawingPath.slice(drawingPath.lastIndexOf('/') + 1)}.rels`,
      ),
    )
    const parsedDrawing: unknown = xmlParser.parse(drawingXml)
    const anchors = asArray(recordChild(parsedDrawing, 'wsDr')?.['twoCellAnchor'])
    anchors.forEach((anchor, anchorIndex) => {
      const chartRelationshipId = readDrawingRelationshipId(anchor)
      const fromCol = readAnchorNumber(anchor, 'from', 'col')
      const fromRow = readAnchorNumber(anchor, 'from', 'row')
      const toCol = readAnchorNumber(anchor, 'to', 'col')
      const toRow = readAnchorNumber(anchor, 'to', 'row')
      if (!chartRelationshipId || fromCol === null || fromRow === null || toCol === null || toRow === null) {
        return
      }
      const chartRelationship = drawingRelationships.find(
        (relationship) => relationship.id === chartRelationshipId && relationship.type === chartRelationshipType,
      )
      if (!chartRelationship) {
        return
      }
      const chartPath = resolveTargetPath(drawingPath, chartRelationship.target)
      const chartMetadata = parseChartXml(getZipText(zip, chartPath) ?? '')
      if (!chartMetadata) {
        return
      }
      charts.push({
        id: readAnchorChartId(anchor, `xlsx-chart:${sheetName}:${String(anchorIndex + 1)}`),
        sheetName,
        address: XLSX.utils.encode_cell({ r: fromRow, c: fromCol }),
        rows: Math.max(1, toRow - fromRow),
        cols: Math.max(1, toCol - fromCol),
        ...chartMetadata,
      })
    })
  })

  return charts.length > 0 ? charts.toSorted((left, right) => left.id.localeCompare(right.id)) : undefined
}
