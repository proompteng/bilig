import * as XLSX from 'xlsx'

import type { WorkbookMergeRangeSnapshot } from '@bilig/protocol'

export function buildMergeEntries(sheetName: string, merges: readonly XLSX.Range[] | undefined): WorkbookMergeRangeSnapshot[] | undefined {
  if (!Array.isArray(merges) || merges.length === 0) {
    return undefined
  }
  const entries = merges.flatMap((range) =>
    range.s.r === range.e.r && range.s.c === range.e.c
      ? []
      : [
          {
            sheetName,
            startAddress: XLSX.utils.encode_cell(range.s),
            endAddress: XLSX.utils.encode_cell(range.e),
          },
        ],
  )
  return entries.length > 0 ? entries : undefined
}
