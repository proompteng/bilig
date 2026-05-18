import * as XLSX from 'xlsx'

import { translateFormulaReferences } from '@bilig/formula'
import { getZipText, type XlsxZipEntries } from './xlsx-zip.js'
import { workbookSheetPath } from './xlsx-workbook-sheet-paths.js'

interface SharedFormulaBase {
  readonly row: number
  readonly column: number
  readonly formula: string
}

export interface WorksheetFormulaCell {
  readonly address: string
  readonly row: number
  readonly column: number
  readonly formulaType: string | null
  readonly sharedIndex: string | null
  readonly ref: string | null
  readonly aca: boolean | null
  readonly bx: boolean | null
  readonly ca: boolean | null
  readonly xmlSpace: string | null
  readonly cellValueType: string | null
  readonly cachedValueRaw: string | null
  readonly cachedValue: string | number | boolean | null | undefined
  readonly rawFormulaXml: string
  readonly formula: string
}

const cellElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?c)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const cellOpeningTagPattern = /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u
const formulaElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?f\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?f)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/u
const formulaOpeningTagPattern = /<(?:[A-Za-z_][\w.-]*:)?f\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u
const formulaTextPattern = /<(?:[A-Za-z_][\w.-]*:)?f\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f>/u
const valueTextPattern = /<(?:[A-Za-z_][\w.-]*:)?v\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?v>/u
const cellAddressPattern = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$/u
const lessThanCode = '<'.charCodeAt(0)
const colonCode = ':'.charCodeAt(0)
const lowerFCode = 'f'.charCodeAt(0)
const upperFCode = 'F'.charCodeAt(0)

function isFormulaNameByte(value: number | undefined): boolean {
  return value === lowerFCode || value === upperFCode
}

function hasFormulaElement(bytes: Uint8Array | undefined): boolean {
  if (!bytes) {
    return false
  }
  for (let index = 0; index < bytes.byteLength - 1; index += 1) {
    const value = bytes[index]
    if (value === lessThanCode && isFormulaNameByte(bytes[index + 1])) {
      return true
    }
    if (value === colonCode && isFormulaNameByte(bytes[index + 1])) {
      return true
    }
  }
  return false
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

function normalizeAddress(address: string): string | null {
  if (!cellAddressPattern.test(address)) {
    return null
  }
  try {
    return XLSX.utils.encode_cell(XLSX.utils.decode_cell(address.replaceAll('$', '')))
  } catch {
    return null
  }
}

function readFormulaXml(cellXml: string): string | null {
  return formulaElementPattern.exec(cellXml)?.[0] ?? null
}

function readFormulaText(formulaXml: string): string {
  return decodeXmlText(formulaTextPattern.exec(formulaXml)?.[1] ?? '')
}

function readBooleanAttribute(openingTag: string, attributeName: string): boolean | null {
  const raw = readXmlAttribute(openingTag, attributeName)
  if (raw === null) {
    return null
  }
  return raw === '1' || raw.toLowerCase() === 'true'
}

function readCachedValueRaw(cellXml: string): string | null {
  const raw = valueTextPattern.exec(cellXml)?.[1]
  return raw === undefined ? null : decodeXmlText(raw)
}

function readCachedValue(raw: string | null, cellValueType: string | null): string | number | boolean | null | undefined {
  if (raw === null) {
    return undefined
  }
  const normalizedType = cellValueType?.trim().toLowerCase()
  if (normalizedType === 'b') {
    return raw === '1' || raw.toLowerCase() === 'true'
  }
  if (normalizedType === 'str' || normalizedType === 's' || normalizedType === 'inlinestr' || normalizedType === 'e') {
    return raw
  }
  const numericValue = Number(raw)
  if (Number.isFinite(numericValue)) {
    return numericValue
  }
  return raw
}

export function readWorksheetFormulaCells(sheetXml: string | null): WorksheetFormulaCell[] {
  if (!sheetXml) {
    return []
  }
  const cells: WorksheetFormulaCell[] = []
  for (const match of sheetXml.matchAll(cellElementPattern)) {
    const cellXml = match[0]
    const cellOpeningTag = cellOpeningTagPattern.exec(cellXml)?.[0]
    const rawAddress = cellOpeningTag ? readXmlAttribute(cellOpeningTag, 'r') : null
    const address = rawAddress ? normalizeAddress(rawAddress) : null
    if (!address) {
      continue
    }
    const formulaXml = readFormulaXml(cellXml)
    const formulaOpeningTag = formulaXml ? formulaOpeningTagPattern.exec(formulaXml)?.[0] : null
    if (!formulaXml || !formulaOpeningTag) {
      continue
    }
    const decodedAddress = XLSX.utils.decode_cell(address)
    const formulaType = readXmlAttribute(formulaOpeningTag, 't')
    const sharedIndex = readXmlAttribute(formulaOpeningTag, 'si')
    const cellValueType = cellOpeningTag ? readXmlAttribute(cellOpeningTag, 't') : null
    const cachedValueRaw = readCachedValueRaw(cellXml)
    cells.push({
      address,
      row: decodedAddress.r,
      column: decodedAddress.c,
      formulaType,
      sharedIndex,
      ref: readXmlAttribute(formulaOpeningTag, 'ref'),
      aca: readBooleanAttribute(formulaOpeningTag, 'aca'),
      bx: readBooleanAttribute(formulaOpeningTag, 'bx'),
      ca: readBooleanAttribute(formulaOpeningTag, 'ca'),
      xmlSpace: readXmlAttribute(formulaOpeningTag, 'xml:space'),
      cellValueType,
      cachedValueRaw,
      cachedValue: readCachedValue(cachedValueRaw, cellValueType),
      rawFormulaXml: formulaXml,
      formula: readFormulaText(formulaXml),
    })
  }
  return cells
}

function readSharedFormulasFromCells(cells: readonly WorksheetFormulaCell[]): Map<string, string> {
  const sharedBases = new Map<string, SharedFormulaBase>()
  const formulas = new Map<string, string>()

  for (const cell of cells) {
    if (cell.formulaType !== 'shared' || cell.sharedIndex === null || cell.formula.trim().length === 0) {
      continue
    }
    sharedBases.set(cell.sharedIndex, {
      row: cell.row,
      column: cell.column,
      formula: cell.formula,
    })
    formulas.set(cell.address, cell.formula)
  }

  for (const cell of cells) {
    if (cell.formulaType !== 'shared' || cell.sharedIndex === null || cell.formula.trim().length > 0) {
      continue
    }
    const base = sharedBases.get(cell.sharedIndex)
    if (!base) {
      continue
    }
    try {
      formulas.set(cell.address, translateFormulaReferences(base.formula, cell.row - base.row, cell.column - base.column))
    } catch {
      // SheetJS has already provided a best-effort formula for unsupported syntax.
    }
  }

  return formulas
}

export function readWorksheetSharedFormulas(sheetXml: string | null): Map<string, string> {
  return readSharedFormulasFromCells(readWorksheetFormulaCells(sheetXml))
}

export function readWorksheetFormulaCellManifest(sheetXml: string | null): Map<string, WorksheetFormulaCell> {
  const cells = readWorksheetFormulaCells(sheetXml)
  const sharedFormulas = readSharedFormulasFromCells(cells)
  const manifest = new Map<string, WorksheetFormulaCell>()
  for (const cell of cells) {
    manifest.set(cell.address, {
      ...cell,
      formula: sharedFormulas.get(cell.address) ?? cell.formula,
    })
  }
  return manifest
}

export function readImportedWorksheetFormulas(
  zip: XlsxZipEntries,
  sheetNames: readonly string[],
  sheetPathsByName: ReadonlyMap<string, string>,
  fallbackSheetPaths: readonly string[],
): Map<string, Map<string, string>> {
  const formulasBySheet = new Map<string, Map<string, string>>()
  sheetNames.forEach((sheetName, index) => {
    const sheetPath = workbookSheetPath(sheetPathsByName, fallbackSheetPaths, sheetName, index)
    if (!sheetPath) {
      return
    }
    if (!hasFormulaElement(zip[sheetPath])) {
      return
    }
    const formulas = new Map(
      [...readWorksheetFormulaCellManifest(getZipText(zip, sheetPath)).entries()]
        .filter((entry) => entry[1].formula.trim().length > 0)
        .map((entry) => [entry[0], entry[1].formula]),
    )
    if (formulas.size > 0) {
      formulasBySheet.set(sheetName, formulas)
    }
  })
  return formulasBySheet
}

export function readImportedWorksheetFormulaManifests(
  zip: XlsxZipEntries,
  sheetNames: readonly string[],
  sheetPathsByName: ReadonlyMap<string, string>,
  fallbackSheetPaths: readonly string[],
): Map<string, Map<string, WorksheetFormulaCell>> {
  const manifestsBySheet = new Map<string, Map<string, WorksheetFormulaCell>>()
  sheetNames.forEach((sheetName, index) => {
    const sheetPath = workbookSheetPath(sheetPathsByName, fallbackSheetPaths, sheetName, index)
    if (!sheetPath) {
      return
    }
    if (!hasFormulaElement(zip[sheetPath])) {
      return
    }
    const manifest = readWorksheetFormulaCellManifest(getZipText(zip, sheetPath))
    if (manifest.size > 0) {
      manifestsBySheet.set(sheetName, manifest)
    }
  })
  return manifestsBySheet
}
