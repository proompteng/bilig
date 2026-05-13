import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import { workbookDirectorySheetPaths, workbookSheetPath, workbookSheetPathsByName } from './xlsx-workbook-sheet-paths.js'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

interface ImportedWorkbookFileNumberFormatsOptions {
  formatCandidateAddressesBySheet?: ReadonlyMap<string, ReadonlySet<string>>
}

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

function workbookRecord(workbook: XLSX.WorkBook): Record<string, unknown> | null {
  const value: unknown = workbook
  return isRecord(value) ? value : null
}

function workbookFiles(workbook: XLSX.WorkBook): unknown {
  return workbookRecord(workbook)?.['files']
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

function getPackageText(files: unknown, sourceZip: XlsxZipEntries | null, path: string): string | null {
  return sourceZip ? getZipText(sourceZip, path) : getFileText(files, path)
}

function normalizeImportedNumberFormatCode(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  return trimmed.length > 0 && trimmed !== 'General' ? trimmed : undefined
}

function builtinNumberFormatCode(formatId: number): string | undefined {
  const ssf: unknown = XLSX.SSF
  if (!isRecord(ssf) || !isRecord(ssf['_table'])) {
    return undefined
  }
  return normalizeImportedNumberFormatCode(ssf['_table'][String(formatId)])
}

function extractStyleXmlElement(stylesXml: string, elementName: string): string | null {
  const qualifiedName = `(?:[A-Za-z_][\\w.-]*:)?${elementName}`
  const expanded = new RegExp(`<${qualifiedName}\\b[^>]*>[\\s\\S]*?<\\/${qualifiedName}>`, 'u').exec(stylesXml)
  if (expanded) {
    return expanded[0]
  }
  const selfClosing = new RegExp(`<${qualifiedName}\\b[^>]*\\/>`, 'u').exec(stylesXml)
  return selfClosing?.[0] ?? null
}

function parseWorkbookNumberFormats(stylesXml: string): Map<number, string> {
  const relevantStylesXml = ['numFmts', 'cellXfs']
    .flatMap((elementName) => {
      const elementXml = extractStyleXmlElement(stylesXml, elementName)
      return elementXml ? [elementXml] : []
    })
    .join('')
  if (relevantStylesXml.length === 0) {
    return new Map()
  }
  const parsed: unknown = xmlParser.parse(`<styleSheet>${relevantStylesXml}</styleSheet>`)
  const styleSheet = recordChild(parsed, 'styleSheet')
  if (!styleSheet) {
    return new Map()
  }

  const customFormats = new Map<number, string>()
  for (const entry of asArray(recordChild(styleSheet, 'numFmts')?.['numFmt'])) {
    if (!isRecord(entry)) {
      continue
    }
    const formatId = numberValue(entry['numFmtId'])
    const formatCode = normalizeImportedNumberFormatCode(entry['formatCode'])
    if (formatId !== null && Number.isSafeInteger(formatId) && formatCode) {
      customFormats.set(formatId, formatCode)
    }
  }

  const formatsByStyleIndex = new Map<number, string>()
  asArray(recordChild(styleSheet, 'cellXfs')?.['xf']).forEach((entry, index) => {
    if (!isRecord(entry)) {
      return
    }
    const formatId = numberValue(entry['numFmtId'])
    if (formatId === null || !Number.isSafeInteger(formatId)) {
      return
    }
    const formatCode = customFormats.get(formatId) ?? builtinNumberFormatCode(formatId)
    if (formatCode) {
      formatsByStyleIndex.set(index, formatCode)
    }
  })
  return formatsByStyleIndex
}

function readXmlAttribute(tag: string, name: string): string | null {
  const doubleQuoted = new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(tag)
  if (doubleQuoted) {
    return doubleQuoted[1] ?? null
  }
  const singleQuoted = new RegExp(`\\b${name}='([^']*)'`, 'u').exec(tag)
  return singleQuoted?.[1] ?? null
}

function parseSheetStyleIndexes(sheetXml: string, candidateAddresses?: ReadonlySet<string>): Map<string, number> {
  const output = new Map<string, number>()
  if (candidateAddresses?.size === 0) {
    return output
  }
  let remainingCandidateCount = candidateAddresses?.size ?? null

  for (const match of sheetXml.matchAll(/<c\b[^>]*>/gu)) {
    const cellTag = match[0]
    const address = readXmlAttribute(cellTag, 'r')
    if (!address) {
      continue
    }
    if (candidateAddresses) {
      if (!candidateAddresses.has(address)) {
        continue
      }
      remainingCandidateCount = remainingCandidateCount === null ? null : remainingCandidateCount - 1
    }
    const styleIndexValue = readXmlAttribute(cellTag, 's')
    if (styleIndexValue === null || styleIndexValue.trim().length === 0) {
      if (remainingCandidateCount === 0) {
        break
      }
      continue
    }
    const styleIndex = Number(styleIndexValue)
    if (Number.isSafeInteger(styleIndex)) {
      output.set(address, styleIndex)
    }
    if (remainingCandidateCount === 0) {
      break
    }
  }

  return output
}

export function readImportedWorkbookFileNumberFormats(
  workbook: XLSX.WorkBook,
  sheetNames: readonly string[],
  options: ImportedWorkbookFileNumberFormatsOptions = {},
  source?: XlsxZipSource,
): Map<string, Map<string, string>> {
  const files = workbookFiles(workbook)
  const sourceZip = source ? readXlsxZipEntries(source) : null
  const stylePath = workbookStylePath(workbook)
  const styleXml = stylePath ? getPackageText(files, sourceZip, stylePath) : null
  if (!styleXml) {
    return new Map()
  }
  const formatsByStyleIndex = parseWorkbookNumberFormats(styleXml)
  if (formatsByStyleIndex.size === 0) {
    return new Map()
  }

  const sheetPathsByName = workbookSheetPathsByName(workbook)
  const fallbackSheetPaths = workbookDirectorySheetPaths(workbook)
  const output = new Map<string, Map<string, string>>()
  sheetNames.forEach((sheetName, index) => {
    const candidateAddresses = options.formatCandidateAddressesBySheet?.get(sheetName)
    if (candidateAddresses?.size === 0) {
      return
    }
    const sheetPath = workbookSheetPath(sheetPathsByName, fallbackSheetPaths, sheetName, index)
    const sheetXml = sheetPath ? getPackageText(files, sourceZip, sheetPath) : null
    if (!sheetXml) {
      return
    }
    const styleIndexes = parseSheetStyleIndexes(sheetXml, candidateAddresses)
    const formats = new Map<string, string>()
    for (const [address, styleIndex] of styleIndexes) {
      const format = formatsByStyleIndex.get(styleIndex)
      if (format) {
        formats.set(address, format)
      }
    }
    if (formats.size > 0) {
      output.set(sheetName, formats)
    }
  })
  return output
}
