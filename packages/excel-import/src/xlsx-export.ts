import * as XLSX from 'xlsx'
import { unzipSync, zipSync } from 'fflate'

import type {
  CellStyleRecord,
  LiteralInput,
  WorkbookAxisEntrySnapshot,
  WorkbookMacroPayloadSnapshot,
  WorkbookMergeRangeSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { addExportArrayFormulasToXlsxBytes } from './xlsx-array-formulas.js'
import { addExportCellMetadataToXlsxBytes } from './xlsx-cell-metadata.js'
import { addMissingBlankCells, addMissingFormattedCells } from './xlsx-cell-insertion.js'
import { addExportCalculationSettingsToXlsxBytes } from './xlsx-calculation-settings.js'
import { addExportChartArtifactsToXlsxBytes } from './xlsx-chart-artifacts.js'
import { addExportChartsToXlsxBytes } from './xlsx-charts.js'
import { addExportLegacyCommentVmlToXlsxBytes } from './xlsx-comment-vml.js'
import { addExportCommentsToWorksheet } from './xlsx-comments.js'
import { addExportConditionalFormatsToXlsxBytes } from './xlsx-conditional-formats.js'
import { addExportControlArtifactsToXlsxBytes } from './xlsx-control-artifacts.js'
import { addExportDataModelArtifactsToXlsxBytes } from './xlsx-data-model-artifacts.js'
import { addExportDataTableFormulasToXlsxBytes } from './xlsx-data-table-formulas.js'
import { buildExportDefinedNames } from './xlsx-defined-names.js'
import { addExportDrawingArtifactsToXlsxBytes } from './xlsx-drawing-artifacts.js'
import { addExportExternalLinkArtifactsToXlsxBytes } from './xlsx-external-link-artifacts.js'
import { addExportWorksheetDimensionsToXlsxBytes, applyExportWorksheetDimensionsToWorksheetXml } from './xlsx-dimensions.js'
import { addExportFiltersToXlsxBytes } from './xlsx-filters.js'
import { addExportFreezePanesToXlsxBytes } from './xlsx-freeze-panes.js'
import { addExportPivotsToXlsxBytes } from './xlsx-pivot-export.js'
import { addExportPrintPageSetupToXlsxBytes } from './xlsx-print-page-setup.js'
import { addExportProtectedRangesToXlsxBytes } from './xlsx-protected-ranges.js'
import { addExportRichTextArtifactsToXlsxBytes } from './xlsx-rich-text-artifacts.js'
import { addExportSheetProtectionsToXlsxBytes } from './xlsx-sheet-protection.js'
import { addExportSortsToXlsxBytes } from './xlsx-sorts.js'
import { addExportSheetTabColorsToXlsxBytes } from './xlsx-tab-colors.js'
import { addExportTablesToXlsxBytes } from './xlsx-tables.js'
import { addExportThreadedCommentArtifactsToXlsxBytes } from './xlsx-threaded-comment-artifacts.js'
import { addExportThemeArtifactToXlsxBytes } from './xlsx-theme-artifacts.js'
import { addExportDataValidationsToXlsxBytes } from './xlsx-validations.js'
import { addExportViewStateToXlsxBytes } from './xlsx-view-state.js'
import { addExportWorkbookProtectionToXlsxBytes } from './xlsx-workbook-protection.js'
import { addExportWorkbookPropertiesToXlsxBytes } from './xlsx-workbook-properties.js'
import { addExportIgnoredErrorsToXlsxBytes } from './xlsx-ignored-errors.js'
import { decodePreservedVbaProjectPayload } from './xlsx-macros.js'
import { addExportPrinterSettingsToXlsxBytes } from './xlsx-printer-settings.js'
import { addExportWorksheetPropertiesToXlsxBytes } from './xlsx-sheet-properties.js'
import { applyExportSheetVisibilitiesToWorkbook } from './xlsx-sheet-visibility.js'
import { addExportSlicerConnectionArtifactsToXlsxBytes } from './xlsx-slicer-connection-artifacts.js'
import { addExportSparklinesToXlsxBytes } from './xlsx-sparklines.js'
import { addExportHyperlinksToWorksheet, hasExportHyperlinks } from './xlsx-hyperlinks.js'
import { preserveSnapshotNumberFormats } from './xlsx-export-number-formats.js'
import { escapeXmlAttribute, getZipText, setXmlAttribute, setZipText } from './xlsx-export-xml.js'
import { encodeExcelEscapedText } from './xlsx-escaped-text.js'
import {
  appendCustomCellXfsToStylesXml,
  readCellXfs,
  updateXmlElementCount,
  worksheetCellElementPattern,
  worksheetCellOpeningTagPattern,
} from './xlsx-style-xml.js'

function buildExportColumns(columns: readonly WorkbookAxisEntrySnapshot[] | undefined): XLSX.ColInfo[] | undefined {
  if (!columns || columns.length === 0) {
    return undefined
  }
  const maxIndex = columns.reduce((max, column) => Math.max(max, column.index), -1)
  if (maxIndex < 0) {
    return undefined
  }
  const output = Array.from({ length: maxIndex + 1 }, (): XLSX.ColInfo => ({}))
  for (const column of columns) {
    const target = output[column.index]
    if (!target) {
      continue
    }
    if (typeof column.size === 'number' && Number.isFinite(column.size) && column.size > 0) {
      target.wpx = column.size
    }
    if (column.hidden === true) {
      target.hidden = true
    }
  }
  return output.some((column) => Object.keys(column).length > 0) ? output : undefined
}

function buildExportRows(rows: readonly WorkbookAxisEntrySnapshot[] | undefined): XLSX.RowInfo[] | undefined {
  if (!rows || rows.length === 0) {
    return undefined
  }
  const maxIndex = rows.reduce((max, row) => Math.max(max, row.index), -1)
  if (maxIndex < 0) {
    return undefined
  }
  const output = Array.from({ length: maxIndex + 1 }, (): XLSX.RowInfo => ({}))
  for (const row of rows) {
    const target = output[row.index]
    if (!target) {
      continue
    }
    if (typeof row.size === 'number' && Number.isFinite(row.size) && row.size > 0) {
      target.hpx = row.size
    }
    if (row.hidden === true) {
      target.hidden = true
    }
  }
  return output.some((row) => Object.keys(row).length > 0) ? output : undefined
}

function sheetCellFormats(sheet: WorkbookSnapshot['sheets'][number]): Map<string, string> {
  const formats = new Map<string, string>()
  for (const cell of sheet.cells) {
    const format = cell.format?.trim()
    if (format && format !== 'General') {
      formats.set(cell.address, format)
    }
  }
  return formats
}

function decodeExportRange(startAddress: string, endAddress: string): XLSX.Range {
  const decoded = XLSX.utils.decode_range(`${startAddress}:${endAddress}`)
  return {
    s: {
      r: Math.min(decoded.s.r, decoded.e.r),
      c: Math.min(decoded.s.c, decoded.e.c),
    },
    e: {
      r: Math.max(decoded.s.r, decoded.e.r),
      c: Math.max(decoded.s.c, decoded.e.c),
    },
  }
}

function addRangeNumberFormats(
  formats: Map<string, string>,
  sheet: WorkbookSnapshot['sheets'][number],
  formatCodesById: ReadonlyMap<string, string>,
): void {
  for (const formatRange of sheet.metadata?.formatRanges ?? []) {
    if (formatRange.range.sheetName !== sheet.name) {
      continue
    }
    const format = formatCodesById.get(formatRange.formatId)?.trim()
    if (!format || format === 'General') {
      continue
    }
    const range = decodeExportRange(formatRange.range.startAddress, formatRange.range.endAddress)
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        formats.set(XLSX.utils.encode_cell({ r: row, c: col }), format)
      }
    }
  }
}

function buildSheetCellFormats(
  sheet: WorkbookSnapshot['sheets'][number],
  formatCodesById: ReadonlyMap<string, string>,
): Map<string, string> {
  const formats = new Map<string, string>()
  addRangeNumberFormats(formats, sheet, formatCodesById)
  for (const [address, format] of sheetCellFormats(sheet)) {
    formats.set(address, format)
  }
  return formats
}

function normalizeRgbColor(value: string | undefined): string | null {
  if (!value) {
    return null
  }
  const normalized = value.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{6}$/u.test(normalized)) {
    return `FF${normalized.toUpperCase()}`
  }
  if (/^[0-9a-fA-F]{8}$/u.test(normalized)) {
    return normalized.toUpperCase()
  }
  return null
}

function fastStyleFontXml(style: CellStyleRecord): string | null {
  const font = style.font
  if (!font) {
    return null
  }
  const children: string[] = []
  if (font.family) {
    children.push(`<name val="${escapeXmlAttribute(font.family)}"/>`)
  }
  if (font.size && font.size > 0) {
    children.push(`<sz val="${String(font.size)}"/>`)
  }
  const color = normalizeRgbColor(font.color)
  if (color) {
    children.push(`<color rgb="${color}"/>`)
  }
  if (font.bold === true) {
    children.push('<b/>')
  }
  if (font.italic === true) {
    children.push('<i/>')
  }
  if (font.underline === true) {
    children.push('<u/>')
  }
  return children.length > 0 ? `<font>${children.join('')}</font>` : null
}

function fastStyleFillXml(style: CellStyleRecord): string | null {
  const color = normalizeRgbColor(style.fill?.backgroundColor)
  return color ? `<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>` : null
}

function fastBorderStyle(value: NonNullable<CellStyleRecord['borders']>[keyof NonNullable<CellStyleRecord['borders']>]): string | null {
  if (!value) {
    return null
  }
  if (value.style === 'dashed') {
    return value.weight === 'medium' ? 'mediumDashed' : 'dashed'
  }
  if (value.style === 'dotted') {
    return 'dotted'
  }
  if (value.style === 'double') {
    return 'thick'
  }
  return value.weight
}

function fastBorderSideXml(
  name: 'left' | 'right' | 'top' | 'bottom',
  value: NonNullable<CellStyleRecord['borders']>[keyof NonNullable<CellStyleRecord['borders']>],
): string {
  const borderStyle = fastBorderStyle(value)
  if (!borderStyle || !value) {
    return `<${name}/>`
  }
  const color = normalizeRgbColor(value.color) ?? 'FF000000'
  return `<${name} style="${borderStyle}"><color rgb="${color}"/></${name}>`
}

function fastStyleBorderXml(style: CellStyleRecord): string | null {
  const borders = style.borders
  if (!borders) {
    return null
  }
  return `<border>${fastBorderSideXml('left', borders.left)}${fastBorderSideXml('right', borders.right)}${fastBorderSideXml(
    'top',
    borders.top,
  )}${fastBorderSideXml('bottom', borders.bottom)}<diagonal/></border>`
}

function fastStyleAlignmentXml(style: CellStyleRecord): string {
  const alignment = style.alignment
  if (!alignment) {
    return ''
  }
  const attributes = [
    alignment.horizontal ? `horizontal="${alignment.horizontal}"` : null,
    alignment.vertical ? `vertical="${alignment.vertical === 'middle' ? 'center' : alignment.vertical}"` : null,
    alignment.wrap === true ? 'wrapText="1"' : null,
    alignment.indent !== undefined && alignment.indent >= 0 ? `indent="${String(alignment.indent)}"` : null,
    alignment.shrinkToFit === true ? 'shrinkToFit="1"' : null,
    alignment.readingOrder !== undefined ? `readingOrder="${String(alignment.readingOrder)}"` : null,
    alignment.textRotation !== undefined ? `textRotation="${String(alignment.textRotation)}"` : null,
    alignment.justifyLastLine === true ? 'justifyLastLine="1"' : null,
  ].filter((entry): entry is string => Boolean(entry))
  return attributes.length > 0 ? `<alignment ${attributes.join(' ')}/>` : ''
}

function fastStyleProtectionXml(style: CellStyleRecord): string {
  if (style.protection === undefined) {
    return ''
  }
  const attributes = [
    style.protection.locked !== undefined ? `locked="${style.protection.locked ? '1' : '0'}"` : null,
    style.protection.hidden !== undefined ? `hidden="${style.protection.hidden ? '1' : '0'}"` : null,
  ].filter((entry): entry is string => Boolean(entry))
  return attributes.length > 0 ? `<protection ${attributes.join(' ')}/>` : '<protection/>'
}

function appendXmlChildren(
  xml: string,
  elementName: 'fonts' | 'fills' | 'borders',
  children: readonly string[],
): { xml: string; startIndex: number } {
  if (children.length === 0) {
    const current = new RegExp(`<${elementName}\\b[^>]*>([\\s\\S]*?)</${elementName}>`, 'u').exec(xml)?.[1] ?? ''
    return { xml, startIndex: Array.from(current.matchAll(new RegExp(`<${elementName.slice(0, -1)}\\b`, 'gu'))).length }
  }
  const pattern = new RegExp(`<${elementName}\\b([^>]*)>([\\s\\S]*?)</${elementName}>`, 'u')
  const match = pattern.exec(xml)
  if (!match) {
    return { xml, startIndex: 0 }
  }
  const body = match[2] ?? ''
  const childName = elementName.slice(0, -1)
  const startIndex = Array.from(body.matchAll(new RegExp(`<${childName}\\b`, 'gu'))).length
  const nextXml = xml.replace(pattern, (_tag, attributes: string, existingBody: string) => {
    const count = startIndex + children.length
    return `<${elementName}${updateXmlElementCount(attributes, count)}>${existingBody}${children.join('')}</${elementName}>`
  })
  return { xml: nextXml, startIndex }
}

function appendFastStyleXfs(
  stylesXml: string,
  styles: readonly CellStyleRecord[],
): { stylesXml: string; styleIndexById: Map<string, number> } {
  let output = stylesXml
  const fonts = styles.map(fastStyleFontXml)
  const fills = styles.map(fastStyleFillXml)
  const borders = styles.map(fastStyleBorderXml)
  const appendedFonts = appendXmlChildren(
    output,
    'fonts',
    fonts.filter((entry): entry is string => Boolean(entry)),
  )
  output = appendedFonts.xml
  const appendedFills = appendXmlChildren(
    output,
    'fills',
    fills.filter((entry): entry is string => Boolean(entry)),
  )
  output = appendedFills.xml
  const appendedBorders = appendXmlChildren(
    output,
    'borders',
    borders.filter((entry): entry is string => Boolean(entry)),
  )
  output = appendedBorders.xml

  let nextFontIndex = appendedFonts.startIndex
  let nextFillIndex = appendedFills.startIndex
  let nextBorderIndex = appendedBorders.startIndex
  const styleXfs: string[] = []
  const styleIndexById = new Map<string, number>()
  const baseStyleCount = readCellXfs(output).length
  styles.forEach((style, index) => {
    const fontXml = fonts[index]
    const fillXml = fills[index]
    const borderXml = borders[index]
    const fontId = fontXml ? nextFontIndex++ : 0
    const fillId = fillXml ? nextFillIndex++ : 0
    const borderId = borderXml ? nextBorderIndex++ : 0
    const alignmentXml = fastStyleAlignmentXml(style)
    const protectionXml = fastStyleProtectionXml(style)
    const childXml = `${alignmentXml}${protectionXml}`
    const attributes = [
      'numFmtId="0"',
      `fontId="${String(fontId)}"`,
      `fillId="${String(fillId)}"`,
      `borderId="${String(borderId)}"`,
      'xfId="0"',
      fontXml ? 'applyFont="1"' : null,
      fillXml ? 'applyFill="1"' : null,
      borderXml ? 'applyBorder="1"' : null,
      alignmentXml ? 'applyAlignment="1"' : null,
      protectionXml ? 'applyProtection="1"' : null,
    ].filter((entry): entry is string => Boolean(entry))
    styleIndexById.set(style.id, baseStyleCount + styleXfs.length)
    styleXfs.push(childXml ? `<xf ${attributes.join(' ')}>${childXml}</xf>` : `<xf ${attributes.join(' ')}/>`)
  })

  return {
    stylesXml: appendCustomCellXfsToStylesXml(output, styleXfs),
    styleIndexById,
  }
}

function applyStyleIndexesToSheetXml(
  sheetXml: string,
  sheet: WorkbookSnapshot['sheets'][number],
  styleIndexById: ReadonlyMap<string, number>,
): string {
  const stylesByAddress = new Map<string, number>()
  const rawStyleArtifactAddresses = new Set([
    ...(sheet.metadata?.styleArtifacts?.cellStyleIndexes ?? []).map((entry) => entry.address),
    ...(sheet.metadata?.styleArtifacts?.blankCellAddresses ?? []),
  ])
  for (const styleRange of sheet.metadata?.styleRanges ?? []) {
    if (styleRange.range.sheetName !== sheet.name) {
      continue
    }
    const styleIndex = styleIndexById.get(styleRange.styleId)
    if (styleIndex === undefined) {
      continue
    }
    const range = decodeExportRange(styleRange.range.startAddress, styleRange.range.endAddress)
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col })
        if (!rawStyleArtifactAddresses.has(address)) {
          stylesByAddress.set(address, styleIndex)
        }
      }
    }
  }
  if (stylesByAddress.size === 0) {
    return sheetXml
  }
  const handledAddresses = new Set<string>()
  let output = sheetXml.replace(worksheetCellElementPattern, (cellXml) => {
    const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
    const address = openingTag ? /\br="([^"]+)"/u.exec(openingTag)?.[1] : undefined
    const styleIndex = address ? stylesByAddress.get(address) : undefined
    if (!openingTag || !address || styleIndex === undefined) {
      return cellXml
    }
    handledAddresses.add(address)
    return cellXml.replace(openingTag, setXmlAttribute(openingTag, 's', String(styleIndex)))
  })
  const missingCells = [...stylesByAddress.entries()]
    .filter(([address]) => !handledAddresses.has(address))
    .map(([address, styleIndex]) => ({ address, styleIndex }))
  return missingCells.length > 0 ? addMissingFormattedCells(output, missingCells) : output
}

function applyStyleArtifactIndexesToSheetXml(
  sheetXml: string,
  styleArtifacts: NonNullable<WorkbookSnapshot['sheets'][number]['metadata']>['styleArtifacts'],
): string {
  const cellStyleIndexes = styleArtifacts?.cellStyleIndexes ?? []
  const blankCellAddresses = styleArtifacts?.blankCellAddresses ?? []
  if (cellStyleIndexes.length === 0 && blankCellAddresses.length === 0) {
    return sheetXml
  }
  const stylesByAddress = new Map(cellStyleIndexes.map((entry) => [entry.address, entry.styleIndex]))
  const handledAddresses = new Set<string>()
  const presentAddresses = new Set<string>()
  let output = sheetXml.replace(worksheetCellElementPattern, (cellXml) => {
    const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
    const address = openingTag ? /\br="([^"]+)"/u.exec(openingTag)?.[1] : undefined
    const styleIndex = address ? stylesByAddress.get(address) : undefined
    if (!openingTag || !address) {
      return cellXml
    }
    presentAddresses.add(address)
    if (styleIndex === undefined) {
      return cellXml
    }
    handledAddresses.add(address)
    return cellXml.replace(openingTag, setXmlAttribute(openingTag, 's', String(styleIndex)))
  })
  const missingCells = [...stylesByAddress.entries()]
    .filter(([address]) => !handledAddresses.has(address))
    .map(([address, styleIndex]) => ({ address, styleIndex }))
  if (missingCells.length > 0) {
    output = addMissingFormattedCells(output, missingCells)
  }
  const missingBlankCellAddresses = blankCellAddresses.filter((address) => !handledAddresses.has(address) && !presentAddresses.has(address))
  return missingBlankCellAddresses.length > 0 ? addMissingBlankCells(output, missingBlankCellAddresses) : output
}

function preserveSnapshotStyles(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const styles = snapshot.workbook.metadata?.styles ?? []
  const hasStyleRanges = snapshot.sheets.some((sheet) => (sheet.metadata?.styleRanges?.length ?? 0) > 0)
  if (styles.length === 0 || !hasStyleRanges) {
    return bytes
  }
  const zip = unzipSync(bytes)
  const stylesXml = getZipText(zip, 'xl/styles.xml')
  if (!stylesXml) {
    return bytes
  }
  const { stylesXml: nextStylesXml, styleIndexById } = appendFastStyleXfs(stylesXml, styles)
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (sheetXml) {
        setZipText(
          zip,
          sheetPath,
          applyExportWorksheetDimensionsToWorksheetXml(applyStyleIndexesToSheetXml(sheetXml, sheet, styleIndexById), sheet.metadata),
        )
      }
    })
  setZipText(zip, 'xl/styles.xml', nextStylesXml)
  return zipSync(zip)
}

function addExportStyleArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const stylesXml = snapshot.workbook.metadata?.styleArtifacts?.stylesXml
  if (!stylesXml) {
    return bytes
  }
  const zip = unzipSync(bytes)
  setZipText(zip, 'xl/styles.xml', stylesXml)
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const styleArtifacts = sheet.metadata?.styleArtifacts
      if (!styleArtifacts || (styleArtifacts.cellStyleIndexes.length === 0 && (styleArtifacts.blankCellAddresses?.length ?? 0) === 0)) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (sheetXml) {
        setZipText(zip, sheetPath, applyStyleArtifactIndexesToSheetXml(sheetXml, styleArtifacts))
      }
    })
  return zipSync(zip)
}

function buildExportMerges(merges: readonly WorkbookMergeRangeSnapshot[] | undefined): XLSX.Range[] | undefined {
  if (!merges || merges.length === 0) {
    return undefined
  }
  return merges.map((merge) => XLSX.utils.decode_range(`${merge.startAddress}:${merge.endAddress}`))
}

function updateWorksheetBounds(bounds: XLSX.Range | null, address: string): XLSX.Range {
  const decoded = XLSX.utils.decode_cell(address)
  if (!bounds) {
    return {
      s: { r: decoded.r, c: decoded.c },
      e: { r: decoded.r, c: decoded.c },
    }
  }
  return {
    s: {
      r: Math.min(bounds.s.r, decoded.r),
      c: Math.min(bounds.s.c, decoded.c),
    },
    e: {
      r: Math.max(bounds.e.r, decoded.r),
      c: Math.max(bounds.e.c, decoded.c),
    },
  }
}

function inferExportWorksheetRange(sheet: WorkbookSnapshot['sheets'][number]): string | undefined {
  let bounds: XLSX.Range | null = null
  for (const cell of sheet.cells) {
    bounds = updateWorksheetBounds(bounds, cell.address)
  }
  for (const merge of sheet.metadata?.merges ?? []) {
    bounds = updateWorksheetBounds(bounds, merge.startAddress)
    bounds = updateWorksheetBounds(bounds, merge.endAddress)
  }
  for (const thread of sheet.metadata?.commentThreads ?? []) {
    bounds = updateWorksheetBounds(bounds, thread.address)
  }
  for (const hyperlink of sheet.metadata?.hyperlinks ?? []) {
    bounds = updateWorksheetBounds(bounds, hyperlink.address)
  }
  for (const styleRange of sheet.metadata?.styleRanges ?? []) {
    bounds = updateWorksheetBounds(bounds, styleRange.range.startAddress)
    bounds = updateWorksheetBounds(bounds, styleRange.range.endAddress)
  }
  for (const formatRange of sheet.metadata?.formatRanges ?? []) {
    bounds = updateWorksheetBounds(bounds, formatRange.range.startAddress)
    bounds = updateWorksheetBounds(bounds, formatRange.range.endAddress)
  }
  return bounds ? XLSX.utils.encode_range(bounds) : undefined
}

function cellTypeForLiteral(value: LiteralInput): XLSX.ExcelDataType | undefined {
  if (typeof value === 'number') {
    return 'n'
  }
  if (typeof value === 'boolean') {
    return 'b'
  }
  if (typeof value === 'string') {
    return 's'
  }
  return undefined
}

function buildExportCell(cell: WorkbookSnapshot['sheets'][number]['cells'][number]): XLSX.CellObject | null {
  const output: XLSX.CellObject = { t: 'z' }
  if (cell.value !== undefined && cell.value !== null) {
    const type = cellTypeForLiteral(cell.value)
    if (type) {
      output.t = type
      output.v = typeof cell.value === 'string' ? encodeExcelEscapedText(cell.value) : cell.value
    }
  }
  if (typeof cell.formula === 'string' && cell.formula.trim().length > 0) {
    output.f = cell.formula.replace(/^=/, '')
    if (output.v === undefined) {
      output.t = 'e'
    } else {
      output.t = output.t ?? 'n'
    }
  }
  return output.v !== undefined || output.f !== undefined ? output : null
}

const invalidExportSheetNameCharacters = ['[', ']', ':', '*', '?', '/', '\\'] as const

function normalizeExportSheetName(name: string, order: number, usedNames: Set<string>): string {
  let sanitized = name
  for (const character of invalidExportSheetNameCharacters) {
    sanitized = sanitized.split(character).join(' ')
  }
  const baseName = sanitized.length > 0 ? sanitized : `Sheet${order + 1}`
  let candidate = baseName.slice(0, 31)
  candidate = candidate.length > 0 ? candidate : `Sheet${order + 1}`
  let suffix = 1
  while (usedNames.has(candidate)) {
    const suffixText = ` ${String(suffix)}`
    candidate = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`
    suffix += 1
  }
  usedNames.add(candidate)
  return candidate
}

function toUint8Array(value: unknown): Uint8Array {
  if (value instanceof Uint8Array) {
    return new Uint8Array(value)
  }
  if (value instanceof ArrayBuffer) {
    return new Uint8Array(value)
  }
  throw new Error('XLSX writer returned unsupported output bytes')
}

function applyMacroCodeNamesToWorkbook(
  workbook: XLSX.WorkBook,
  macroPayload: WorkbookMacroPayloadSnapshot | undefined,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): void {
  if (!macroPayload?.workbookCodeName && (!macroPayload?.sheetCodeNames || macroPayload.sheetCodeNames.length === 0)) {
    return
  }

  const workbookMetadata = {
    ...workbook.Workbook,
  }
  if (macroPayload.workbookCodeName) {
    workbookMetadata.WBProps = {
      ...workbookMetadata.WBProps,
      CodeName: macroPayload.workbookCodeName,
    }
  }

  if (macroPayload.sheetCodeNames && macroPayload.sheetCodeNames.length > 0) {
    const codeNamesByExportSheetName = new Map<string, string>()
    for (const entry of macroPayload.sheetCodeNames) {
      const exportSheetName = exportSheetNamesByOriginalName.get(entry.sheetName) ?? entry.sheetName
      codeNamesByExportSheetName.set(exportSheetName, entry.codeName)
    }
    const existingSheets = workbookMetadata.Sheets ?? []
    workbookMetadata.Sheets = workbook.SheetNames.map((sheetName, index) => {
      const codeName = codeNamesByExportSheetName.get(sheetName) ?? existingSheets[index]?.CodeName
      return {
        ...existingSheets[index],
        name: sheetName,
        ...(codeName ? { CodeName: codeName } : {}),
      }
    })
  }

  workbook.Workbook = workbookMetadata
}

export function exportXlsx(snapshot: WorkbookSnapshot): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const usedNames = new Set<string>()
  const exportSheetNamesByOriginalName = new Map<string, string>()
  const exportSheetIndexesByOriginalName = new Map<string, number>()
  const formatCodesById = new Map((snapshot.workbook.metadata?.formats ?? []).map((format) => [format.id, format.code]))
  const exportSheetFormats = snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .map((sheet) => buildSheetCellFormats(sheet, formatCodesById))

  for (const sheet of snapshot.sheets.toSorted((left, right) => left.order - right.order)) {
    const worksheet: XLSX.WorkSheet = {}
    for (const cell of sheet.cells) {
      const exportCell = buildExportCell(cell)
      if (exportCell) {
        worksheet[cell.address] = exportCell
      }
    }

    const ref = inferExportWorksheetRange(sheet)
    if (ref) {
      worksheet['!ref'] = ref
    }
    const columns = buildExportColumns(sheet.metadata?.columns)
    if (columns) {
      worksheet['!cols'] = columns
    }
    const rows = buildExportRows(sheet.metadata?.rows)
    if (rows) {
      worksheet['!rows'] = rows
    }
    const merges = buildExportMerges(sheet.metadata?.merges)
    if (merges) {
      worksheet['!merges'] = merges
    }
    addExportCommentsToWorksheet(worksheet, sheet.metadata?.commentThreads)
    if (hasExportHyperlinks(sheet.metadata)) {
      addExportHyperlinksToWorksheet(worksheet, sheet)
    }

    const exportSheetName = normalizeExportSheetName(sheet.name, sheet.order, usedNames)
    exportSheetNamesByOriginalName.set(sheet.name, exportSheetName)
    exportSheetIndexesByOriginalName.set(sheet.name, exportSheetIndexesByOriginalName.size)
    XLSX.utils.book_append_sheet(workbook, worksheet, exportSheetName)
  }

  const definedNames = buildExportDefinedNames(
    snapshot.workbook.metadata?.definedNames,
    exportSheetNamesByOriginalName,
    exportSheetIndexesByOriginalName,
  )
  if (definedNames) {
    workbook.Workbook = {
      ...workbook.Workbook,
      Names: definedNames,
    }
  }

  const macroPayload = snapshot.workbook.metadata?.macroPayloads?.[0]
  const preservedVbaProject = decodePreservedVbaProjectPayload(macroPayload)
  if (preservedVbaProject) {
    const macroWorkbook = workbook as { vbaraw?: Uint8Array }
    macroWorkbook.vbaraw = preservedVbaProject
    applyMacroCodeNamesToWorkbook(workbook, macroPayload, exportSheetNamesByOriginalName)
  }
  applyExportSheetVisibilitiesToWorkbook(workbook, snapshot)

  const bytes = toUint8Array(
    XLSX.write(workbook, {
      bookType: preservedVbaProject ? 'xlsm' : 'xlsx',
      type: 'buffer',
      bookVBA: Boolean(preservedVbaProject),
    }) as unknown,
  )
  const pivotBytes = addExportPivotsToXlsxBytes(
    addExportTablesToXlsxBytes(
      addExportDataValidationsToXlsxBytes(
        addExportConditionalFormatsToXlsxBytes(
          addExportSortsToXlsxBytes(
            addExportFiltersToXlsxBytes(
              addExportProtectedRangesToXlsxBytes(
                addExportSheetProtectionsToXlsxBytes(
                  addExportViewStateToXlsxBytes(
                    addExportFreezePanesToXlsxBytes(
                      addExportWorksheetPropertiesToXlsxBytes(
                        addExportSheetTabColorsToXlsxBytes(
                          addExportCalculationSettingsToXlsxBytes(
                            addExportWorkbookProtectionToXlsxBytes(addExportWorkbookPropertiesToXlsxBytes(bytes, snapshot), snapshot),
                            snapshot,
                          ),
                          snapshot,
                        ),
                        snapshot,
                      ),
                      snapshot,
                    ),
                    snapshot,
                  ),
                  snapshot,
                ),
                snapshot,
              ),
              snapshot,
            ),
            snapshot,
          ),
          snapshot,
        ),
        snapshot,
        exportSheetNamesByOriginalName,
      ),
      snapshot,
      exportSheetNamesByOriginalName,
    ),
    snapshot,
    exportSheetNamesByOriginalName,
  )
  const chartArtifactBytes = addExportChartArtifactsToXlsxBytes(pivotBytes, snapshot)
  const enrichedBytes = addExportChartsToXlsxBytes(chartArtifactBytes, snapshot, exportSheetNamesByOriginalName)
  const richTextArtifactBytes = addExportRichTextArtifactsToXlsxBytes(enrichedBytes, snapshot)
  const artifactStyledBytes = addExportStyleArtifactsToXlsxBytes(richTextArtifactBytes, snapshot)
  const styledBytes = preserveSnapshotStyles(artifactStyledBytes, snapshot)
  const formattedBytes = preserveSnapshotNumberFormats(styledBytes, exportSheetFormats)
  const themeArtifactBytes = addExportThemeArtifactToXlsxBytes(formattedBytes, snapshot)
  const dimensionedBytes = addExportWorksheetDimensionsToXlsxBytes(themeArtifactBytes, snapshot)
  const drawingArtifactBytes = addExportDrawingArtifactsToXlsxBytes(dimensionedBytes, snapshot)
  const ignoredErrorsBytes = addExportIgnoredErrorsToXlsxBytes(drawingArtifactBytes, snapshot)
  const sparklineBytes = addExportSparklinesToXlsxBytes(ignoredErrorsBytes, snapshot)
  const controlArtifactBytes = addExportControlArtifactsToXlsxBytes(sparklineBytes, snapshot)
  const dataTableFormulaBytes = addExportDataTableFormulasToXlsxBytes(controlArtifactBytes, snapshot)
  const arrayFormulaBytes = addExportArrayFormulasToXlsxBytes(dataTableFormulaBytes, snapshot)
  const dataModelArtifactBytes = addExportDataModelArtifactsToXlsxBytes(arrayFormulaBytes, snapshot)
  const externalLinkArtifactBytes = addExportExternalLinkArtifactsToXlsxBytes(dataModelArtifactBytes, snapshot)
  const slicerConnectionArtifactBytes = addExportSlicerConnectionArtifactsToXlsxBytes(externalLinkArtifactBytes, snapshot)
  const threadedCommentArtifactBytes = addExportThreadedCommentArtifactsToXlsxBytes(slicerConnectionArtifactBytes, snapshot)
  const printPageSetupBytes = addExportPrintPageSetupToXlsxBytes(
    addExportLegacyCommentVmlToXlsxBytes(threadedCommentArtifactBytes, snapshot),
    snapshot,
  )
  return addExportCellMetadataToXlsxBytes(addExportPrinterSettingsToXlsxBytes(printPageSetupBytes, snapshot), snapshot)
}
