import * as XLSX from 'xlsx'

import type { LiteralInput, WorkbookDefinedNameSnapshot, WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'

interface ImportedSheetBounds {
  readonly endRow: number
  readonly endCol: number
}

function normalizeA1Address(value: string): string | null {
  const normalized = value.trim().replace(/\$/g, '').toUpperCase()
  if (!/^[A-Z]+[1-9][0-9]*$/.test(normalized)) {
    return null
  }
  try {
    return XLSX.utils.encode_cell(XLSX.utils.decode_cell(normalized))
  } catch {
    return null
  }
}

function absoluteA1Address(value: string): string | null {
  const normalized = normalizeA1Address(value)
  if (!normalized) {
    return null
  }
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(normalized)
  return match ? `$${match[1]}$${match[2]}` : null
}

function parseQuotedSheetReference(value: string): { sheetName: string; reference: string } | null {
  if (!value.startsWith("'")) {
    return null
  }
  let sheetName = ''
  for (let index = 1; index < value.length; index += 1) {
    const character = value[index]
    if (character === "'" && value[index + 1] === "'") {
      sheetName += "'"
      index += 1
      continue
    }
    if (character === "'" && value[index + 1] === '!') {
      const reference = value.slice(index + 2).trim()
      return sheetName.trim().length > 0 && reference.length > 0 ? { sheetName, reference } : null
    }
    sheetName += character
  }
  return null
}

function parseSheetReference(value: string): { sheetName: string; reference: string } | null {
  const quoted = parseQuotedSheetReference(value)
  if (quoted) {
    return quoted
  }
  const separatorIndex = value.indexOf('!')
  if (separatorIndex <= 0) {
    return null
  }
  const sheetName = value.slice(0, separatorIndex).trim()
  const reference = value.slice(separatorIndex + 1).trim()
  return sheetName.length > 0 && reference.length > 0 ? { sheetName, reference } : null
}

function normalizeColumnReference(value: string): string | null {
  const normalized = value.trim().replace(/\$/g, '').toUpperCase()
  if (!/^[A-Z]+$/.test(normalized)) {
    return null
  }
  try {
    return XLSX.utils.encode_col(XLSX.utils.decode_col(normalized))
  } catch {
    return null
  }
}

function normalizeRowReference(value: string): number | null {
  const normalized = value.trim().replace(/\$/g, '')
  if (!/^[1-9][0-9]*$/.test(normalized)) {
    return null
  }
  const row = Number.parseInt(normalized, 10)
  return Number.isSafeInteger(row) && row > 0 ? row : null
}

function readImportedSheetBounds(workbook: XLSX.WorkBook): ReadonlyMap<string, ImportedSheetBounds> {
  const bounds = new Map<string, ImportedSheetBounds>()
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet?.['!ref']) {
      continue
    }
    try {
      const decoded = XLSX.utils.decode_range(sheet['!ref'])
      bounds.set(sheetName, {
        endRow: Math.max(decoded.e.r, 0),
        endCol: Math.max(decoded.e.c, 0),
      })
    } catch {
      // Invalid worksheet dimensions should not prevent formula-preserving name import.
    }
  }
  return bounds
}

function parseBoundedWholeColumnReference(
  sheetName: string,
  start: string,
  end: string,
  sheetBoundsByName: ReadonlyMap<string, ImportedSheetBounds>,
): WorkbookDefinedNameValueSnapshot | null {
  const startColumn = normalizeColumnReference(start)
  const endColumn = normalizeColumnReference(end)
  const sheetBounds = sheetBoundsByName.get(sheetName)
  if (!startColumn || !endColumn || !sheetBounds) {
    return null
  }
  return {
    kind: 'range-ref',
    sheetName,
    startAddress: `${startColumn}1`,
    endAddress: XLSX.utils.encode_cell({ r: sheetBounds.endRow, c: XLSX.utils.decode_col(endColumn) }),
  }
}

function parseBoundedWholeRowReference(
  sheetName: string,
  start: string,
  end: string,
  sheetBoundsByName: ReadonlyMap<string, ImportedSheetBounds>,
): WorkbookDefinedNameValueSnapshot | null {
  const startRow = normalizeRowReference(start)
  const endRow = normalizeRowReference(end)
  const sheetBounds = sheetBoundsByName.get(sheetName)
  if (!startRow || !endRow || !sheetBounds) {
    return null
  }
  return {
    kind: 'range-ref',
    sheetName,
    startAddress: XLSX.utils.encode_cell({ r: startRow - 1, c: 0 }),
    endAddress: XLSX.utils.encode_cell({ r: endRow - 1, c: sheetBounds.endCol }),
  }
}

function parseDefinedNameReferenceValue(
  sheetName: string,
  reference: string,
  sheetBoundsByName: ReadonlyMap<string, ImportedSheetBounds>,
): WorkbookDefinedNameValueSnapshot | null {
  const parts = reference.split(':')
  if (parts.length === 1) {
    const address = normalizeA1Address(parts[0] ?? '')
    return address ? { kind: 'cell-ref', sheetName, address } : null
  }
  if (parts.length === 2) {
    const startAddress = normalizeA1Address(parts[0] ?? '')
    const endAddress = normalizeA1Address(parts[1] ?? '')
    if (startAddress && endAddress) {
      return { kind: 'range-ref', sheetName, startAddress, endAddress }
    }
    return (
      parseBoundedWholeColumnReference(sheetName, parts[0] ?? '', parts[1] ?? '', sheetBoundsByName) ??
      parseBoundedWholeRowReference(sheetName, parts[0] ?? '', parts[1] ?? '', sheetBoundsByName)
    )
  }
  return null
}

function parseDefinedNameScalarValue(value: string): WorkbookDefinedNameValueSnapshot | null {
  const trimmed = value.trim()
  if (/^[+-]?(?:(?:[0-9]+(?:\.[0-9]*)?)|(?:\.[0-9]+))(?:[eE][+-]?[0-9]+)?$/.test(trimmed)) {
    const numberValue = Number(trimmed)
    return Number.isFinite(numberValue) ? { kind: 'scalar', value: numberValue } : null
  }
  if (/^TRUE$/i.test(trimmed)) {
    return { kind: 'scalar', value: true }
  }
  if (/^FALSE$/i.test(trimmed)) {
    return { kind: 'scalar', value: false }
  }
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return { kind: 'scalar', value: trimmed.slice(1, -1).replace(/""/g, '"') }
  }
  return null
}

function parseImportedDefinedNameValue(
  ref: string,
  sheetBoundsByName: ReadonlyMap<string, ImportedSheetBounds>,
): WorkbookDefinedNameValueSnapshot | null {
  const trimmed = ref.trim()
  if (trimmed.length === 0) {
    return null
  }
  const expression = trimmed.startsWith('=') ? trimmed.slice(1).trim() : trimmed
  const sheetReference = parseSheetReference(expression)
  if (sheetReference && !sheetReference.sheetName.startsWith('[')) {
    const parsedReference = parseDefinedNameReferenceValue(sheetReference.sheetName, sheetReference.reference, sheetBoundsByName)
    if (parsedReference) {
      return parsedReference
    }
  }
  const scalar = parseDefinedNameScalarValue(expression)
  if (scalar) {
    return scalar
  }
  return { kind: 'formula', formula: trimmed.startsWith('=') ? trimmed : `=${trimmed}` }
}

export function readImportedDefinedNames(workbook: XLSX.WorkBook): {
  definedNames: WorkbookDefinedNameSnapshot[] | undefined
  ignoredCount: number
} {
  const entries = workbook.Workbook?.Names
  if (!Array.isArray(entries) || entries.length === 0) {
    return { definedNames: undefined, ignoredCount: 0 }
  }

  const definedNamesByNormalizedName = new Map<string, WorkbookDefinedNameSnapshot>()
  const sheetBoundsByName = readImportedSheetBounds(workbook)
  let ignoredCount = 0
  for (const entry of entries) {
    const name = typeof entry.Name === 'string' ? entry.Name.trim() : ''
    const ref = typeof entry.Ref === 'string' ? entry.Ref.trim() : ''
    if (name.length === 0 || ref.length === 0 || typeof entry.Sheet === 'number') {
      ignoredCount += 1
      continue
    }
    const value = parseImportedDefinedNameValue(ref, sheetBoundsByName)
    if (!value) {
      ignoredCount += 1
      continue
    }
    definedNamesByNormalizedName.set(name.toUpperCase(), { name, value })
  }

  const definedNames = [...definedNamesByNormalizedName.values()].toSorted((left, right) => left.name.localeCompare(right.name))
  return {
    definedNames: definedNames.length > 0 ? definedNames : undefined,
    ignoredCount,
  }
}

function quoteDefinedNameSheetName(sheetName: string): string {
  return `'${sheetName.replace(/'/g, "''")}'`
}

function exportDefinedNameSheetReference(
  sheetName: string,
  reference: string,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): string | null {
  const exportSheetName = exportSheetNamesByOriginalName.get(sheetName)
  return exportSheetName ? `${quoteDefinedNameSheetName(exportSheetName)}!${reference}` : null
}

function formatExportLiteralDefinedNameValue(value: LiteralInput): string | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? String(value) : null
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  if (typeof value === 'string') {
    if (value.startsWith('=')) {
      return value.slice(1).trim()
    }
    return `"${value.replace(/"/g, '""')}"`
  }
  return '""'
}

function formatExportDefinedNameValue(
  value: WorkbookDefinedNameValueSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): string | null {
  if (typeof value !== 'object' || value === null) {
    return formatExportLiteralDefinedNameValue(value)
  }
  switch (value.kind) {
    case 'scalar':
      return formatExportLiteralDefinedNameValue(value.value)
    case 'cell-ref': {
      const address = absoluteA1Address(value.address)
      return address ? exportDefinedNameSheetReference(value.sheetName, address, exportSheetNamesByOriginalName) : null
    }
    case 'range-ref': {
      const startAddress = absoluteA1Address(value.startAddress)
      const endAddress = absoluteA1Address(value.endAddress)
      return startAddress && endAddress
        ? exportDefinedNameSheetReference(value.sheetName, `${startAddress}:${endAddress}`, exportSheetNamesByOriginalName)
        : null
    }
    case 'formula':
      return value.formula.startsWith('=') ? value.formula.slice(1).trim() : value.formula.trim()
    case 'structured-ref':
      return value.columnName.trim().length > 0 ? `${value.tableName}[${value.columnName}]` : value.tableName
  }
}

export function buildExportDefinedNames(
  definedNames: readonly WorkbookDefinedNameSnapshot[] | undefined,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): XLSX.DefinedName[] | undefined {
  if (!definedNames || definedNames.length === 0) {
    return undefined
  }
  const output: XLSX.DefinedName[] = []
  for (const definedName of definedNames) {
    const name = definedName.name.trim()
    if (name.length === 0) {
      continue
    }
    const ref = formatExportDefinedNameValue(definedName.value, exportSheetNamesByOriginalName)
    if (ref && ref.length > 0) {
      output.push({ Name: name, Ref: ref })
    }
  }
  return output.length > 0 ? output : undefined
}
