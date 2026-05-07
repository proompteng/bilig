import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'
import type * as XLSXStyle from 'xlsx-js-style'

import type {
  CellBorderSideSnapshot,
  CellBorderStyle,
  CellBorderWeight,
  CellHorizontalAlignment,
  CellStyleAlignmentSnapshot,
  CellStyleBordersSnapshot,
  CellStyleFillSnapshot,
  CellStyleFontSnapshot,
  CellStyleRecord,
  CellVerticalAlignment,
  SheetStyleRangeSnapshot,
} from '@bilig/protocol'

type ImportedCellStyle = Omit<CellStyleRecord, 'id'>
type ExportCellAlignment = NonNullable<XLSXStyle.CellStyle['alignment']>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

function stringValue(value: unknown): string | null {
  return typeof value === 'string' ? value : null
}

function numberValue(value: unknown): number | null {
  if (typeof value !== 'string' || value.trim().length === 0) {
    return null
  }
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function getFileText(files: unknown, path: string): string | null {
  if (!isRecord(files)) {
    return null
  }
  const file = files[normalizeZipPath(path)]
  if (!isRecord(file)) {
    return null
  }
  const content = file['content']
  if (typeof content === 'string') {
    return content
  }
  if (content instanceof ArrayBuffer) {
    return new TextDecoder().decode(content)
  }
  if (ArrayBuffer.isView(content)) {
    return new TextDecoder().decode(content)
  }
  return null
}

function workbookRecord(workbook: XLSX.WorkBook): Record<string, unknown> | null {
  const value: unknown = workbook
  return isRecord(value) ? value : null
}

function workbookFiles(workbook: XLSX.WorkBook): unknown {
  return workbookRecord(workbook)?.['files']
}

function workbookSheetPaths(workbook: XLSX.WorkBook): string[] {
  const directory = workbookRecord(workbook)?.['Directory']
  if (!isRecord(directory)) {
    return []
  }
  return asArray(directory['sheets']).flatMap((entry) => (typeof entry === 'string' ? [entry] : []))
}

function workbookStylePath(workbook: XLSX.WorkBook): string | null {
  const directory = workbookRecord(workbook)?.['Directory']
  if (!isRecord(directory)) {
    return null
  }
  if (typeof directory['style'] === 'string') {
    return directory['style']
  }
  const firstStylePath = asArray(directory['styles']).find((entry) => typeof entry === 'string')
  return typeof firstStylePath === 'string' ? firstStylePath : null
}

function normalizeRgbColor(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null
  }
  const normalized = value.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return `#${normalized.toLowerCase()}`
  }
  if (/^[0-9a-fA-F]{8}$/.test(normalized)) {
    return `#${normalized.slice(2).toLowerCase()}`
  }
  return null
}

function toArgbColor(value: string): string | null {
  const normalized = normalizeRgbColor(value)
  return normalized ? `FF${normalized.slice(1).toUpperCase()}` : null
}

function readColorRecord(value: unknown): string | null {
  return normalizeRgbColor(
    recordChild(value, 'color')?.['rgb'] ?? recordChild(value, 'fgColor')?.['rgb'] ?? recordChild(value, 'bgColor')?.['rgb'],
  )
}

function readFont(font: unknown): CellStyleFontSnapshot | undefined {
  if (!isRecord(font)) {
    return undefined
  }
  const output: CellStyleFontSnapshot = {}
  const name = stringValue(recordChild(font, 'name')?.['val'])
  const size = numberValue(recordChild(font, 'sz')?.['val'])
  const color = readColorRecord(font)
  if (name && name.length > 0) {
    output.family = name
  }
  if (size !== null && size > 0) {
    output.size = size
  }
  if (font['b'] !== undefined) {
    output.bold = true
  }
  if (font['i'] !== undefined) {
    output.italic = true
  }
  if (font['u'] !== undefined) {
    output.underline = true
  }
  if (color) {
    output.color = color
  }
  return Object.keys(output).length > 0 ? output : undefined
}

function readFill(fill: unknown): CellStyleFillSnapshot | undefined {
  const patternFill = recordChild(fill, 'patternFill')
  if (!patternFill || patternFill['patternType'] !== 'solid') {
    return undefined
  }
  const backgroundColor = normalizeRgbColor(recordChild(patternFill, 'fgColor')?.['rgb'] ?? recordChild(patternFill, 'bgColor')?.['rgb'])
  return backgroundColor ? { backgroundColor } : undefined
}

function readBorderKind(value: unknown): { style: CellBorderStyle; weight: CellBorderWeight } | null {
  switch (value) {
    case 'hair':
    case 'thin':
      return { style: 'solid', weight: 'thin' }
    case 'medium':
      return { style: 'solid', weight: 'medium' }
    case 'thick':
      return { style: 'solid', weight: 'thick' }
    case 'dashed':
      return { style: 'dashed', weight: 'thin' }
    case 'mediumDashed':
      return { style: 'dashed', weight: 'medium' }
    case 'dotted':
      return { style: 'dotted', weight: 'thin' }
    default:
      return null
  }
}

function readBorderSide(side: unknown): CellBorderSideSnapshot | undefined {
  if (!isRecord(side)) {
    return undefined
  }
  const borderKind = readBorderKind(side['style'])
  if (!borderKind) {
    return undefined
  }
  return {
    ...borderKind,
    color: normalizeRgbColor(recordChild(side, 'color')?.['rgb']) ?? '#000000',
  }
}

function readBorders(border: unknown): CellStyleBordersSnapshot | undefined {
  if (!isRecord(border)) {
    return undefined
  }
  const top = readBorderSide(border['top'])
  const right = readBorderSide(border['right'])
  const bottom = readBorderSide(border['bottom'])
  const left = readBorderSide(border['left'])
  const borders: CellStyleBordersSnapshot = {}
  if (top) {
    borders.top = top
  }
  if (right) {
    borders.right = right
  }
  if (bottom) {
    borders.bottom = bottom
  }
  if (left) {
    borders.left = left
  }
  return Object.keys(borders).length > 0 ? borders : undefined
}

function readHorizontalAlignment(value: unknown): CellHorizontalAlignment | undefined {
  switch (value) {
    case 'general':
    case 'left':
    case 'center':
    case 'right':
      return value
    default:
      return undefined
  }
}

function readVerticalAlignment(value: unknown): CellVerticalAlignment | undefined {
  switch (value) {
    case 'top':
      return 'top'
    case 'center':
    case 'middle':
      return 'middle'
    case 'bottom':
      return 'bottom'
    default:
      return undefined
  }
}

function readAlignment(alignment: unknown): CellStyleAlignmentSnapshot | undefined {
  if (!isRecord(alignment)) {
    return undefined
  }
  const horizontal = readHorizontalAlignment(alignment['horizontal'])
  const vertical = readVerticalAlignment(alignment['vertical'])
  const indent = numberValue(alignment['indent'])
  const output: CellStyleAlignmentSnapshot = {
    ...(horizontal ? { horizontal } : {}),
    ...(vertical ? { vertical } : {}),
    ...(alignment['wrapText'] === 'true' || alignment['wrapText'] === '1' ? { wrap: true } : {}),
    ...(indent !== null && indent >= 0 ? { indent } : {}),
  }
  return Object.keys(output).length > 0 ? output : undefined
}

function parseWorkbookStyles(stylesXml: string): Map<number, ImportedCellStyle> {
  const parsed: unknown = xmlParser.parse(stylesXml)
  const styleSheet = recordChild(parsed, 'styleSheet')
  if (!styleSheet) {
    return new Map()
  }

  const fonts = asArray(recordChild(styleSheet, 'fonts')?.['font']).map(readFont)
  const fills = asArray(recordChild(styleSheet, 'fills')?.['fill']).map(readFill)
  const borders = asArray(recordChild(styleSheet, 'borders')?.['border']).map(readBorders)
  const cellXfs = asArray(recordChild(styleSheet, 'cellXfs')?.['xf'])
  const styles = new Map<number, ImportedCellStyle>()

  cellXfs.forEach((entry, index) => {
    if (!isRecord(entry)) {
      return
    }
    const fontId = numberValue(entry['fontId'])
    const fillId = numberValue(entry['fillId'])
    const borderId = numberValue(entry['borderId'])
    const font = fontId !== null ? fonts[fontId] : undefined
    const fill = fillId !== null ? fills[fillId] : undefined
    const bordersValue = borderId !== null ? borders[borderId] : undefined
    const alignment = readAlignment(entry['alignment'])
    const style: ImportedCellStyle = {
      ...(entry['applyFill'] === '1' && fill ? { fill } : {}),
      ...(entry['applyFont'] === '1' && font ? { font } : {}),
      ...(entry['applyAlignment'] === '1' && alignment ? { alignment } : {}),
      ...(entry['applyBorder'] === '1' && bordersValue ? { borders: bordersValue } : {}),
    }
    if (Object.keys(style).length > 0) {
      styles.set(index, style)
    }
  })

  return styles
}

function parseSheetStyleIndexes(sheetXml: string): Map<string, number> {
  const parsed: unknown = xmlParser.parse(sheetXml)
  const worksheet = recordChild(parsed, 'worksheet')
  const sheetData = recordChild(worksheet, 'sheetData')
  const output = new Map<string, number>()

  for (const row of asArray(sheetData?.['row'])) {
    for (const cell of asArray(recordChild(row, 'c') ?? (isRecord(row) ? row['c'] : undefined))) {
      if (!isRecord(cell)) {
        continue
      }
      const address = stringValue(cell['r'])
      const styleIndex = numberValue(cell['s'])
      if (address && styleIndex !== null) {
        output.set(address, styleIndex)
      }
    }
  }

  return output
}

export function readImportedWorkbookFileStyles(
  workbook: XLSX.WorkBook,
  sheetNames: readonly string[],
): Map<string, Map<string, ImportedCellStyle>> {
  const files = workbookFiles(workbook)
  const stylePath = workbookStylePath(workbook)
  const styleXml = stylePath ? getFileText(files, stylePath) : null
  if (!styleXml) {
    return new Map()
  }
  const stylesByIndex = parseWorkbookStyles(styleXml)
  if (stylesByIndex.size === 0) {
    return new Map()
  }

  const sheetPaths = workbookSheetPaths(workbook)
  const output = new Map<string, Map<string, ImportedCellStyle>>()
  sheetNames.forEach((sheetName, index) => {
    const sheetPath = sheetPaths[index]
    const sheetXml = sheetPath ? getFileText(files, sheetPath) : null
    if (!sheetXml) {
      return
    }
    const styleIndexes = parseSheetStyleIndexes(sheetXml)
    const cellStyles = new Map<string, ImportedCellStyle>()
    for (const [address, styleIndex] of styleIndexes) {
      const style = stylesByIndex.get(styleIndex)
      if (style) {
        cellStyles.set(address, style)
      }
    }
    if (cellStyles.size > 0) {
      output.set(sheetName, cellStyles)
    }
  })

  return output
}

function exportColor(value: string): XLSXStyle.CellStyleColor | undefined {
  const rgb = toArgbColor(value)
  return rgb ? { rgb } : undefined
}

function exportAlignment(alignment: CellStyleAlignmentSnapshot): ExportCellAlignment {
  return {
    ...(alignment.horizontal && alignment.horizontal !== 'general' ? { horizontal: alignment.horizontal } : {}),
    ...(alignment.vertical ? { vertical: alignment.vertical === 'middle' ? 'center' : alignment.vertical } : {}),
    ...(alignment.wrap === true ? { wrapText: true } : {}),
    ...(alignment.indent !== undefined && alignment.indent >= 0 ? { indent: alignment.indent } : {}),
  }
}

function exportBorderSide(side: CellBorderSideSnapshot): { color: XLSXStyle.CellStyleColor; style?: XLSXStyle.BorderType } {
  const color = exportColor(side.color) ?? { rgb: 'FF000000' }
  switch (side.style) {
    case 'solid':
      return { color, style: side.weight }
    case 'dashed':
      return { color, style: side.weight === 'thin' ? 'dashed' : 'mediumDashed' }
    case 'dotted':
      return { color, style: 'dotted' }
    case 'double':
      return { color, style: 'thick' }
  }
}

function exportStyle(style: CellStyleRecord): XLSXStyle.CellStyle {
  const output: XLSXStyle.CellStyle = {}
  const fillColor = style.fill?.backgroundColor ? exportColor(style.fill.backgroundColor) : undefined
  if (fillColor) {
    output.fill = { patternType: 'solid', fgColor: fillColor }
  }

  if (style.font) {
    const font: NonNullable<XLSXStyle.CellStyle['font']> = {}
    const fontColor = style.font.color ? exportColor(style.font.color) : undefined
    if (style.font.family) {
      font.name = style.font.family
    }
    if (style.font.size) {
      font.sz = style.font.size
    }
    if (style.font.bold === true) {
      font.bold = true
    }
    if (style.font.italic === true) {
      font.italic = true
    }
    if (style.font.underline === true) {
      font.underline = true
    }
    if (fontColor) {
      font.color = fontColor
    }
    if (Object.keys(font).length > 0) {
      output.font = font
    }
  }

  if (style.alignment) {
    const alignment = exportAlignment(style.alignment)
    if (Object.keys(alignment).length > 0) {
      output.alignment = alignment
    }
  }

  if (style.borders) {
    const border: NonNullable<XLSXStyle.CellStyle['border']> = {}
    if (style.borders.top) {
      border.top = exportBorderSide(style.borders.top)
    }
    if (style.borders.right) {
      border.right = exportBorderSide(style.borders.right)
    }
    if (style.borders.bottom) {
      border.bottom = exportBorderSide(style.borders.bottom)
    }
    if (style.borders.left) {
      border.left = exportBorderSide(style.borders.left)
    }
    if (Object.keys(border).length > 0) {
      output.border = border
    }
  }

  return output
}

function isStyledWorksheetCell(value: unknown): value is XLSXStyle.CellObject {
  return isRecord(value) && typeof value['t'] === 'string'
}

export function addExportStylesToWorksheet(
  worksheet: XLSXStyle.WorkSheet,
  styleRanges: readonly SheetStyleRangeSnapshot[] | undefined,
  styles: readonly CellStyleRecord[] | undefined,
): void {
  if (!styleRanges || styleRanges.length === 0 || !styles || styles.length === 0) {
    return
  }

  const stylesById = new Map(styles.map((style) => [style.id, style]))
  for (const styleRange of styleRanges) {
    const style = stylesById.get(styleRange.styleId)
    if (!style) {
      continue
    }
    const exportStyleValue = exportStyle(style)
    const range = XLSX.utils.decode_range(`${styleRange.range.startAddress}:${styleRange.range.endAddress}`)
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let column = range.s.c; column <= range.e.c; column += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: column })
        const existingCell = worksheet[address]
        const cell = isStyledWorksheetCell(existingCell) ? existingCell : ({ t: 'z' } satisfies XLSXStyle.CellObject)
        cell.s = exportStyleValue
        worksheet[address] = cell
      }
    }
  }
}
