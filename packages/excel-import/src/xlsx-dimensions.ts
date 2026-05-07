import { unzipSync, zipSync } from 'fflate'

import type { WorkbookAxisEntrySnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { escapeXmlAttribute, getZipText, setXmlAttribute, setZipText } from './xlsx-export-xml.js'

interface ExportRowMetadata {
  readonly rowNumber: number
  readonly size?: number
  readonly hidden?: true
}

function normalizeExportRowMetadata(rows: readonly WorkbookAxisEntrySnapshot[] | undefined): ExportRowMetadata[] {
  if (!rows || rows.length === 0) {
    return []
  }
  return rows
    .flatMap((row) => {
      if (!Number.isSafeInteger(row.index) || row.index < 0) {
        return []
      }
      const size = typeof row.size === 'number' && Number.isFinite(row.size) && row.size > 0 ? row.size : undefined
      if (size === undefined && row.hidden !== true) {
        return []
      }
      return [
        {
          rowNumber: row.index + 1,
          ...(size !== undefined ? { size } : {}),
          ...(row.hidden === true ? { hidden: true as const } : {}),
        },
      ]
    })
    .toSorted((left, right) => left.rowNumber - right.rowNumber)
}

function formatRowSize(size: number): string {
  return Number.isInteger(size) ? String(size) : String(Number(size.toFixed(4)))
}

function rowOpeningTagPattern(rowNumber: number): RegExp {
  return new RegExp(`<row\\b(?=[^>]*\\br="${String(rowNumber)}"(?:\\s|/|>))[^>]*>`, 'u')
}

function readRowNumber(rowTag: string): number | null {
  const match = /\br="([0-9]+)"/u.exec(rowTag)
  if (!match) {
    return null
  }
  const rowNumber = Number(match[1])
  return Number.isSafeInteger(rowNumber) && rowNumber > 0 ? rowNumber : null
}

function applyRowMetadata(rowTag: string, row: ExportRowMetadata): string {
  let nextTag = rowTag
  if (row.size !== undefined) {
    nextTag = setXmlAttribute(setXmlAttribute(nextTag, 'ht', formatRowSize(row.size)), 'customHeight', '1')
  }
  if (row.hidden === true) {
    nextTag = setXmlAttribute(nextTag, 'hidden', '1')
  }
  return nextTag
}

function buildEmptyRowXml(row: ExportRowMetadata): string {
  let rowTag = `<row r="${escapeXmlAttribute(String(row.rowNumber))}"/>`
  rowTag = applyRowMetadata(rowTag, row)
  return rowTag
}

function insertRowIntoSheetData(sheetDataXml: string, rowXml: string, rowNumber: number): string {
  const closeIndex = sheetDataXml.lastIndexOf('</sheetData>')
  if (closeIndex < 0) {
    return sheetDataXml
  }
  let insertIndex = closeIndex
  for (const match of sheetDataXml.matchAll(/<row\b[^>]*>/gu)) {
    const existingRowNumber = readRowNumber(match[0])
    if (existingRowNumber !== null && existingRowNumber > rowNumber) {
      insertIndex = match.index
      break
    }
  }
  return `${sheetDataXml.slice(0, insertIndex)}${rowXml}${sheetDataXml.slice(insertIndex)}`
}

function upsertWorksheetRowMetadata(sheetXml: string, rows: readonly ExportRowMetadata[]): string {
  let nextXml = sheetXml
  for (const row of rows) {
    const existingRowPattern = rowOpeningTagPattern(row.rowNumber)
    if (existingRowPattern.test(nextXml)) {
      nextXml = nextXml.replace(existingRowPattern, (rowTag) => applyRowMetadata(rowTag, row))
      continue
    }
    const rowXml = buildEmptyRowXml(row)
    if (/<sheetData\b[^>]*\/>/u.test(nextXml)) {
      nextXml = nextXml.replace(/<sheetData\b([^>]*)\/>/u, (_match, attributes: string) => `<sheetData${attributes}>${rowXml}</sheetData>`)
      continue
    }
    const sheetDataMatch = /<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/u.exec(nextXml)
    if (!sheetDataMatch) {
      continue
    }
    const sheetDataXml = sheetDataMatch[0]
    const updatedSheetDataXml = insertRowIntoSheetData(sheetDataXml, rowXml, row.rowNumber)
    nextXml = `${nextXml.slice(0, sheetDataMatch.index)}${updatedSheetDataXml}${nextXml.slice(sheetDataMatch.index + sheetDataXml.length)}`
  }
  return nextXml
}

export function hasExportRowMetadata(snapshot: WorkbookSnapshot): boolean {
  return snapshot.sheets.some((sheet) => normalizeExportRowMetadata(sheet.metadata?.rows).length > 0)
}

export function applyExportRowMetadataToWorksheetXml(sheetXml: string, rows: readonly WorkbookAxisEntrySnapshot[] | undefined): string {
  const normalizedRows = normalizeExportRowMetadata(rows)
  return normalizedRows.length > 0 ? upsertWorksheetRowMetadata(sheetXml, normalizedRows) : sheetXml
}

export function addExportRowMetadataToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!hasExportRowMetadata(snapshot)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const updatedSheetXml = applyExportRowMetadataToWorksheetXml(sheetXml, sheet.metadata?.rows)
      if (updatedSheetXml === sheetXml) {
        return
      }
      setZipText(zip, sheetPath, updatedSheetXml)
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}
