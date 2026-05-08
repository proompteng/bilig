import * as XLSX from 'xlsx'

import type { SheetMetadataSnapshot, WorkbookHyperlinkSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { worksheetCellRecords } from './xlsx-worksheet-cells.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readNonEmptyString(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  return trimmed.length > 0 ? trimmed : undefined
}

function readHyperlinkTarget(link: Record<string, unknown>): string | undefined {
  const target = readNonEmptyString(link['Target'])
  if (target) {
    return target
  }
  const location = readNonEmptyString(link['location'])
  return location ? `#${location}` : undefined
}

export function readImportedSheetHyperlinks(sheetName: string, sheet: XLSX.WorkSheet): WorkbookHyperlinkSnapshot[] | undefined {
  const hyperlinks: WorkbookHyperlinkSnapshot[] = []
  for (const { address, cell } of worksheetCellRecords(sheet)) {
    const link = cell['l']
    if (!isRecord(link)) {
      continue
    }
    const target = readHyperlinkTarget(link)
    if (!target) {
      continue
    }
    const tooltip = readNonEmptyString(link['Tooltip'])
    const display = readNonEmptyString(link['display'])
    hyperlinks.push({
      sheetName,
      address,
      target,
      ...(tooltip ? { tooltip } : {}),
      ...(display ? { display } : {}),
    })
  }
  return hyperlinks.length > 0
    ? hyperlinks.toSorted(
        (left, right) =>
          XLSX.utils.decode_cell(left.address).r - XLSX.utils.decode_cell(right.address).r || left.address.localeCompare(right.address),
      )
    : undefined
}

export function addExportHyperlinksToWorksheet(worksheet: XLSX.WorkSheet, sheet: WorkbookSnapshot['sheets'][number]): void {
  for (const hyperlink of sheet.metadata?.hyperlinks ?? []) {
    if (hyperlink.sheetName !== sheet.name || hyperlink.target.trim().length === 0) {
      continue
    }
    const existingCell: unknown = worksheet[hyperlink.address]
    const cell = isRecord(existingCell) ? existingCell : { t: 'z' }
    cell['l'] = {
      Target: hyperlink.target,
      ...(hyperlink.tooltip ? { Tooltip: hyperlink.tooltip } : {}),
      ...(hyperlink.display ? { display: hyperlink.display } : {}),
    }
    worksheet[hyperlink.address] = cell
  }
}

export function hasExportHyperlinks(metadata: SheetMetadataSnapshot | undefined): boolean {
  return (metadata?.hyperlinks?.length ?? 0) > 0
}
