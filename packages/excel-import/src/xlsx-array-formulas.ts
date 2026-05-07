import * as XLSX from 'xlsx'

import type { WorkbookSpillSnapshot } from '@bilig/protocol'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isWorksheetCellAddress(value: string): boolean {
  return /^[A-Z]{1,3}[1-9][0-9]*$/u.test(value)
}

function decodeArrayFormulaRange(value: string): XLSX.Range | undefined {
  try {
    return XLSX.utils.decode_range(value)
  } catch {
    return undefined
  }
}

export function readImportedArrayFormulaSpills(sheetName: string, sheet: XLSX.WorkSheet): WorkbookSpillSnapshot[] | undefined {
  const spills: WorkbookSpillSnapshot[] = []
  for (const address in sheet) {
    const cell: unknown = sheet[address]
    if (!isWorksheetCellAddress(address) || !isRecord(cell)) {
      continue
    }
    const formula = cell['f']
    const arrayRangeText = cell['F']
    if (typeof formula !== 'string' || formula.trim().length === 0 || typeof arrayRangeText !== 'string') {
      continue
    }
    const range = decodeArrayFormulaRange(arrayRangeText.trim())
    if (!range) {
      continue
    }
    const owner = XLSX.utils.decode_cell(address)
    if (range.s.r !== owner.r || range.s.c !== owner.c) {
      continue
    }
    const rows = range.e.r - range.s.r + 1
    const cols = range.e.c - range.s.c + 1
    if (rows <= 1 && cols <= 1) {
      continue
    }
    spills.push({
      sheetName,
      address: XLSX.utils.encode_cell(range.s),
      rows,
      cols,
    })
  }
  return spills.length > 0 ? spills : undefined
}
