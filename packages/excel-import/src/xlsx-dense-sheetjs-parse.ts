import type { Unzipped } from 'fflate'
import * as XLSX from 'xlsx'

const textDecoder = new TextDecoder()

export function shouldUseDenseSheetJsParse(
  data: Uint8Array,
  workbookZip: Unzipped | null,
  options: {
    readonly minByteLength: number
    readonly maxColumnCount: number
  },
): boolean {
  if (!workbookZip || data.byteLength < options.minByteLength) {
    return false
  }
  let sawWorksheetDimension = false
  for (const path of Object.keys(workbookZip)) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/u.test(path)) {
      continue
    }
    const bytes = workbookZip[path]
    if (!bytes) {
      continue
    }
    const dimensionRef = readWorksheetDimensionRef(bytes)
    if (!dimensionRef) {
      continue
    }
    sawWorksheetDimension = true
    const range = XLSX.utils.decode_range(dimensionRef.includes(':') ? dimensionRef : `${dimensionRef}:${dimensionRef}`)
    if (range.e.c + 1 > options.maxColumnCount) {
      return false
    }
  }
  return sawWorksheetDimension
}

function readWorksheetDimensionRef(bytes: Uint8Array): string | null {
  const headerXml = textDecoder.decode(bytes.subarray(0, Math.min(bytes.byteLength, 4096)))
  return /<dimension\b[^>]*\bref="([^"]+)"/u.exec(headerXml)?.[1] ?? null
}
