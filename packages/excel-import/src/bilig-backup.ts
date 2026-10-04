import { parseCellAddress } from '@bilig/formula'
import { isWorkbookSnapshot } from '@bilig/protocol'
import { createSheetPreview } from './workbook-import-helpers.js'
import { createWorkbookPreview } from './workbook-import-preview.js'
import type { ImportedWorkbook } from './workbook-import-result.js'
import { BILIG_CONTENT_TYPE } from './workbook-import-content-types.js'

export function importBiligBackup(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const snapshot: unknown = JSON.parse(new TextDecoder().decode(data))
  if (!isWorkbookSnapshot(snapshot) || snapshot.sheets.length === 0 || !snapshot.workbook.name.trim()) {
    throw new Error('Invalid Bilig workbook backup')
  }
  const sheets = snapshot.sheets.map((sheet) => {
    let rowCount = 0
    let columnCount = 0
    const previewCells = new Map<string, string>()
    for (const cell of sheet.cells) {
      const { row, col } = parseCellAddress(cell.address, sheet.name)
      rowCount = Math.max(rowCount, row + 1)
      columnCount = Math.max(columnCount, col + 1)
      if (row < 8 && col < 6) {
        previewCells.set(`${row}:${col}`, cell.formula !== undefined ? `=${cell.formula.replace(/^=/, '')}` : String(cell.value ?? ''))
      }
    }
    return createSheetPreview({
      name: sheet.name,
      rowCount,
      columnCount,
      nonEmptyCellCount: sheet.cells.length,
      readCellText: (row, col) => previewCells.get(`${row}:${col}`) ?? '',
    })
  })
  return {
    snapshot,
    workbookName: snapshot.workbook.name,
    sheetNames: snapshot.sheets.map((sheet) => sheet.name),
    warnings: [],
    preview: createWorkbookPreview({
      contentType: BILIG_CONTENT_TYPE,
      fileName,
      fileSizeBytes: data.byteLength,
      workbookName: snapshot.workbook.name,
      sheets,
      warnings: [],
    }),
  }
}
