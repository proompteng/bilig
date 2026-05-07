import * as XLSX from 'xlsx'
import * as XLSXStyle from 'xlsx-js-style'
import { unzipSync, zipSync } from 'fflate'

import type {
  CellStyleRecord,
  LiteralInput,
  WorkbookAxisEntrySnapshot,
  WorkbookMacroPayloadSnapshot,
  WorkbookMergeRangeSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { addExportCalculationSettingsToXlsxBytes } from './xlsx-calculation-settings.js'
import { addExportChartsToXlsxBytes } from './xlsx-charts.js'
import { addExportCommentsToWorksheet } from './xlsx-comments.js'
import { addExportConditionalFormatsToXlsxBytes } from './xlsx-conditional-formats.js'
import { buildExportDefinedNames } from './xlsx-defined-names.js'
import { addExportFiltersToXlsxBytes } from './xlsx-filters.js'
import { addExportFreezePanesToXlsxBytes } from './xlsx-freeze-panes.js'
import { addExportPivotsToXlsxBytes } from './xlsx-pivots.js'
import { addExportProtectedRangesToXlsxBytes } from './xlsx-protected-ranges.js'
import { addExportSheetProtectionsToXlsxBytes } from './xlsx-sheet-protection.js'
import { addExportSortsToXlsxBytes } from './xlsx-sorts.js'
import { addExportStylesToWorksheet } from './xlsx-styles.js'
import { addExportTablesToXlsxBytes } from './xlsx-tables.js'
import { addExportDataValidationsToXlsxBytes } from './xlsx-validations.js'
import { addExportWorkbookPropertiesToXlsxBytes } from './xlsx-workbook-properties.js'
import { decodePreservedVbaProjectPayload } from './xlsx-macros.js'
import {
  addCustomNumberFormatsToStylesXml,
  customNumberFormatStartId,
  escapeXmlAttribute,
  getZipText,
  readXmlNumberAttribute,
  repairLeadingZeroNumberFormatIds,
  setXmlAttribute,
  setZipText,
} from './xlsx-export-xml.js'

const largeSimpleExportCellCountThreshold = 100_000

function isRecord(value: unknown): value is Record<PropertyKey, unknown> {
  return typeof value === 'object' && value !== null
}

function isXlsxStyleRuntime(value: unknown): value is typeof XLSXStyle {
  if (!isRecord(value)) {
    return false
  }
  const utils = value['utils']
  return (
    typeof value['write'] === 'function' &&
    isRecord(utils) &&
    typeof utils['book_new'] === 'function' &&
    typeof utils['book_append_sheet'] === 'function'
  )
}

function resolveXlsxStyleRuntime(moduleValue: typeof XLSXStyle): typeof XLSXStyle {
  if (isXlsxStyleRuntime(moduleValue)) {
    return moduleValue
  }

  // Node ESM exposes this CommonJS package behind `default`; Bun and Vite expose
  // the named runtime surface directly.
  if (isRecord(moduleValue) && isXlsxStyleRuntime(moduleValue['default'])) {
    return moduleValue['default']
  }

  throw new Error('xlsx-js-style did not expose the expected workbook writer API')
}

const XLSXStyleRuntime = resolveXlsxStyleRuntime(XLSXStyle)

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

function readCellXfs(stylesXml: string): readonly string[] {
  const match = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/u.exec(stylesXml)
  if (!match) {
    return []
  }
  const body = match[1] ?? ''
  const entries: string[] = []
  let cursor = 0
  while (cursor < body.length) {
    const start = body.indexOf('<xf', cursor)
    if (start < 0) {
      break
    }
    const openingEnd = body.indexOf('>', start)
    if (openingEnd < 0) {
      break
    }
    if (body[openingEnd - 1] === '/') {
      entries.push(body.slice(start, openingEnd + 1))
      cursor = openingEnd + 1
      continue
    }
    const closingStart = body.indexOf('</xf>', openingEnd + 1)
    if (closingStart < 0) {
      break
    }
    entries.push(body.slice(start, closingStart + '</xf>'.length))
    cursor = closingStart + '</xf>'.length
  }
  return entries
}

function updateElementCount(openingAttributes: string, count: number): string {
  return /\scount="[^"]*"/u.test(openingAttributes)
    ? openingAttributes.replace(/\scount="[^"]*"/u, ` count="${String(count)}"`)
    : `${openingAttributes} count="${String(count)}"`
}

function appendCustomCellXfsToStylesXml(stylesXml: string, xfs: readonly string[]): string {
  if (xfs.length === 0) {
    return stylesXml
  }
  return stylesXml.replace(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/u, (_match, attributes: string, body: string) => {
    const count = Array.from(body.matchAll(/<xf\b/gu)).length + xfs.length
    return `<cellXfs${updateElementCount(attributes, count)}>${body}${xfs.join('')}</cellXfs>`
  })
}

function styleXfWithNumberFormat(xf: string, numberFormatId: number): string {
  return xf.replace(/<xf\b[^>]*(?:\/>|>)/u, (openingTag) =>
    setXmlAttribute(
      setXmlAttribute(setXmlAttribute(openingTag, 'numFmtId', String(numberFormatId)), 'applyNumberFormat', '1'),
      'xfId',
      '0',
    ),
  )
}

class ExportNumberFormatRegistry {
  private readonly baseXfs: readonly string[]
  private readonly numberFormatIdsByCode = new Map<string, number>()
  private readonly styleIndexesByKey = new Map<string, number>()
  private readonly addedFormatIdsByCode = new Map<string, number>()
  private readonly addedXfs: string[] = []
  private nextNumberFormatId: number

  constructor(stylesXml: string) {
    this.baseXfs = readCellXfs(stylesXml)
    const usedIds = [...stylesXml.matchAll(/\bnumFmtId="([0-9]+)"/gu)].map((match) => Number(match[1])).filter(Number.isSafeInteger)
    this.nextNumberFormatId = Math.max(customNumberFormatStartId, ...usedIds.map((id) => id + 1))
  }

  styleIndexFor(baseStyleIndex: number, formatCode: string): number {
    const key = `${String(baseStyleIndex)}\u0000${formatCode}`
    const existingStyleIndex = this.styleIndexesByKey.get(key)
    if (existingStyleIndex !== undefined) {
      return existingStyleIndex
    }
    const numberFormatId = this.numberFormatIdFor(formatCode)
    const baseXf = this.baseXfs[baseStyleIndex] ?? '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    const styleIndex = this.baseXfs.length + this.addedXfs.length
    this.addedXfs.push(styleXfWithNumberFormat(baseXf, numberFormatId))
    this.styleIndexesByKey.set(key, styleIndex)
    return styleIndex
  }

  apply(stylesXml: string): string {
    return appendCustomCellXfsToStylesXml(addCustomNumberFormatsToStylesXml(stylesXml, this.addedFormatIdsByCode), this.addedXfs)
  }

  private numberFormatIdFor(formatCode: string): number {
    const existingId = this.numberFormatIdsByCode.get(formatCode)
    if (existingId !== undefined) {
      return existingId
    }
    const id = this.nextNumberFormatId
    this.nextNumberFormatId += 1
    this.numberFormatIdsByCode.set(formatCode, id)
    this.addedFormatIdsByCode.set(formatCode, id)
    return id
  }
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

function addMissingFormattedCells(sheetXml: string, cells: readonly { readonly address: string; readonly styleIndex: number }[]): string {
  let output = sheetXml
  const byRow = new Map<number, string[]>()
  for (const cell of cells) {
    const rowNumber = XLSX.utils.decode_cell(cell.address).r + 1
    const cellXml = `<c r="${cell.address}" s="${String(cell.styleIndex)}" t="z"/>`
    byRow.set(rowNumber, [...(byRow.get(rowNumber) ?? []), cellXml])
  }
  for (const [rowNumber, rowCells] of byRow) {
    const rowPattern = new RegExp(`<row\\b(?=[^>]*\\br="${String(rowNumber)}")[^>]*(?:/>|>[\\s\\S]*?</row>)`, 'u')
    if (rowPattern.test(output)) {
      output = output.replace(rowPattern, (rowXml) =>
        rowXml.endsWith('/>')
          ? rowXml.replace(/\/>$/u, `>${rowCells.join('')}</row>`)
          : rowXml.replace('</row>', `${rowCells.join('')}</row>`),
      )
    } else {
      output = output.replace('</sheetData>', `<row r="${String(rowNumber)}">${rowCells.join('')}</row></sheetData>`)
    }
  }
  return output
}

function applyNumberFormatsToSheetXml(
  sheetXml: string,
  formats: ReadonlyMap<string, string>,
  registry: ExportNumberFormatRegistry,
): string {
  if (formats.size === 0) {
    return sheetXml
  }
  const handledAddresses = new Set<string>()
  let output = sheetXml.replace(/<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/gu, (cellXml) => {
    const openingTag = /<c\b[^>]*(?:\/>|>)/u.exec(cellXml)?.[0]
    const address = openingTag ? /\br="([^"]+)"/u.exec(openingTag)?.[1] : undefined
    const format = address ? formats.get(address) : undefined
    if (!openingTag || !address || !format) {
      return cellXml
    }
    handledAddresses.add(address)
    const baseStyleIndex = readXmlNumberAttribute(openingTag, 's') ?? 0
    const styleIndex = registry.styleIndexFor(baseStyleIndex, format)
    return cellXml.replace(openingTag, setXmlAttribute(openingTag, 's', String(styleIndex)))
  })
  const missingCells = [...formats.entries()]
    .filter(([address]) => !handledAddresses.has(address))
    .map(([address, format]) => ({
      address,
      styleIndex: registry.styleIndexFor(0, format),
    }))
  if (missingCells.length > 0) {
    output = addMissingFormattedCells(output, missingCells)
  }
  return output
}

function preserveSnapshotNumberFormats(bytes: Uint8Array, sheetFormats: readonly ReadonlyMap<string, string>[]): Uint8Array {
  if (sheetFormats.every((formats) => formats.size === 0)) {
    return repairLeadingZeroNumberFormatIds(bytes)
  }
  const zip = unzipSync(repairLeadingZeroNumberFormatIds(bytes))
  const stylesXml = getZipText(zip, 'xl/styles.xml')
  if (!stylesXml) {
    return zipSync(zip)
  }
  const registry = new ExportNumberFormatRegistry(stylesXml)
  sheetFormats.forEach((formats, sheetIndex) => {
    if (formats.size === 0) {
      return
    }
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const sheetXml = getZipText(zip, sheetPath)
    if (!sheetXml) {
      return
    }
    setZipText(zip, sheetPath, applyNumberFormatsToSheetXml(sheetXml, formats, registry))
  })
  setZipText(zip, 'xl/styles.xml', registry.apply(stylesXml))
  return zipSync(zip)
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
    alignment.horizontal && alignment.horizontal !== 'general' ? `horizontal="${alignment.horizontal}"` : null,
    alignment.vertical ? `vertical="${alignment.vertical === 'middle' ? 'center' : alignment.vertical}"` : null,
    alignment.wrap === true ? 'wrapText="1"' : null,
    alignment.indent !== undefined && alignment.indent >= 0 ? `indent="${String(alignment.indent)}"` : null,
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
    return `<${elementName}${updateElementCount(attributes, count)}>${existingBody}${children.join('')}</${elementName}>`
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
        stylesByAddress.set(XLSX.utils.encode_cell({ r: row, c: col }), styleIndex)
      }
    }
  }
  if (stylesByAddress.size === 0) {
    return sheetXml
  }
  const handledAddresses = new Set<string>()
  let output = sheetXml.replace(/<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/gu, (cellXml) => {
    const openingTag = /<c\b[^>]*(?:\/>|>)/u.exec(cellXml)?.[0]
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

function preserveSnapshotStyles(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const styles = snapshot.workbook.metadata?.styles ?? []
  if (styles.length === 0 || snapshot.sheets.every((sheet) => (sheet.metadata?.styleRanges?.length ?? 0) === 0)) {
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
        setZipText(zip, sheetPath, applyStyleIndexesToSheetXml(sheetXml, sheet, styleIndexById))
      }
    })
  setZipText(zip, 'xl/styles.xml', nextStylesXml)
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

function updateWorksheetBoundsForAxis(bounds: XLSX.Range | null, axis: 'row' | 'column', index: number): XLSX.Range {
  const address = axis === 'row' ? XLSX.utils.encode_cell({ r: index, c: 0 }) : XLSX.utils.encode_cell({ r: 0, c: index })
  return updateWorksheetBounds(bounds, address)
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
  for (const styleRange of sheet.metadata?.styleRanges ?? []) {
    bounds = updateWorksheetBounds(bounds, styleRange.range.startAddress)
    bounds = updateWorksheetBounds(bounds, styleRange.range.endAddress)
  }
  for (const formatRange of sheet.metadata?.formatRanges ?? []) {
    bounds = updateWorksheetBounds(bounds, formatRange.range.startAddress)
    bounds = updateWorksheetBounds(bounds, formatRange.range.endAddress)
  }
  for (const row of sheet.metadata?.rows ?? []) {
    bounds = updateWorksheetBoundsForAxis(bounds, 'row', row.index)
  }
  for (const column of sheet.metadata?.columns ?? []) {
    bounds = updateWorksheetBoundsForAxis(bounds, 'column', column.index)
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

function buildExportCell(cell: WorkbookSnapshot['sheets'][number]['cells'][number]): XLSXStyle.CellObject | null {
  const output: XLSXStyle.CellObject = { t: 'z' }
  if (cell.value !== undefined && cell.value !== null) {
    const type = cellTypeForLiteral(cell.value)
    if (type) {
      output.t = type
      output.v = cell.value
    }
  }
  if (typeof cell.formula === 'string' && cell.formula.trim().length > 0) {
    output.f = cell.formula.trim().replace(/^=/, '')
    if (output.v === undefined) {
      output.t = 'n'
      output.v = 0
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
  const baseName = sanitized.trim().length > 0 ? sanitized : `Sheet${order + 1}`
  let candidate = baseName.slice(0, 31)
  if (candidate.trim().length === 0) {
    candidate = sanitized.trim().slice(0, 31) || `Sheet${order + 1}`
  }
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
  workbook: XLSXStyle.WorkBook,
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

function countSnapshotCells(snapshot: WorkbookSnapshot): number {
  return snapshot.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0)
}

function hasLargeSimpleExportShape(snapshot: WorkbookSnapshot): boolean {
  if (countSnapshotCells(snapshot) < largeSimpleExportCellCountThreshold) {
    return false
  }
  const metadata = snapshot.workbook.metadata
  if (
    (metadata?.macroPayloads?.length ?? 0) > 0 ||
    (metadata?.tables?.length ?? 0) > 0 ||
    (metadata?.charts?.length ?? 0) > 0 ||
    (metadata?.pivots?.length ?? 0) > 0 ||
    (metadata?.definedNames?.length ?? 0) > 0
  ) {
    return false
  }
  return snapshot.sheets.every((sheet) => {
    if (sheet.cells.some((cell) => cell.formula !== undefined)) {
      return false
    }
    const sheetMetadata = sheet.metadata
    return (
      (sheetMetadata?.commentThreads?.length ?? 0) === 0 &&
      (sheetMetadata?.validations?.length ?? 0) === 0 &&
      (sheetMetadata?.conditionalFormats?.length ?? 0) === 0 &&
      (sheetMetadata?.filters?.length ?? 0) === 0 &&
      (sheetMetadata?.sorts?.length ?? 0) === 0 &&
      (sheetMetadata?.protectedRanges?.length ?? 0) === 0 &&
      sheetMetadata?.sheetProtection === undefined &&
      sheetMetadata?.freezePane === undefined
    )
  })
}

export function exportXlsx(snapshot: WorkbookSnapshot): Uint8Array {
  const useLargeSimpleFastPath = hasLargeSimpleExportShape(snapshot)
  const workbook = useLargeSimpleFastPath ? XLSX.utils.book_new() : XLSXStyleRuntime.utils.book_new()
  const usedNames = new Set<string>()
  const exportSheetNamesByOriginalName = new Map<string, string>()
  const exportSheetIndexesByOriginalName = new Map<string, number>()
  const formatCodesById = new Map((snapshot.workbook.metadata?.formats ?? []).map((format) => [format.id, format.code]))
  const exportSheetFormats = snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .map((sheet) => buildSheetCellFormats(sheet, formatCodesById))

  for (const sheet of snapshot.sheets.toSorted((left, right) => left.order - right.order)) {
    const worksheet: XLSXStyle.WorkSheet = {}
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
    if (!useLargeSimpleFastPath) {
      addExportStylesToWorksheet(worksheet, sheet.metadata?.styleRanges, snapshot.workbook.metadata?.styles)
    }
    addExportCommentsToWorksheet(worksheet, sheet.metadata?.commentThreads)

    const exportSheetName = normalizeExportSheetName(sheet.name, sheet.order, usedNames)
    exportSheetNamesByOriginalName.set(sheet.name, exportSheetName)
    exportSheetIndexesByOriginalName.set(sheet.name, exportSheetIndexesByOriginalName.size)
    if (useLargeSimpleFastPath) {
      XLSX.utils.book_append_sheet(workbook, worksheet, exportSheetName)
    } else {
      XLSXStyleRuntime.utils.book_append_sheet(workbook, worksheet, exportSheetName)
    }
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

  const bytes = toUint8Array(
    useLargeSimpleFastPath
      ? (XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }) as unknown)
      : (XLSXStyleRuntime.write(workbook, {
          bookType: preservedVbaProject ? 'xlsm' : 'xlsx',
          type: 'buffer',
          cellStyles: true,
          bookVBA: Boolean(preservedVbaProject),
        }) as unknown),
  )
  const enrichedBytes = addExportChartsToXlsxBytes(
    addExportPivotsToXlsxBytes(
      addExportTablesToXlsxBytes(
        addExportDataValidationsToXlsxBytes(
          addExportConditionalFormatsToXlsxBytes(
            addExportSortsToXlsxBytes(
              addExportFiltersToXlsxBytes(
                addExportProtectedRangesToXlsxBytes(
                  addExportSheetProtectionsToXlsxBytes(
                    addExportFreezePanesToXlsxBytes(
                      addExportCalculationSettingsToXlsxBytes(addExportWorkbookPropertiesToXlsxBytes(bytes, snapshot), snapshot),
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
    ),
    snapshot,
    exportSheetNamesByOriginalName,
  )
  return preserveSnapshotNumberFormats(preserveSnapshotStyles(enrichedBytes, snapshot), exportSheetFormats)
}
