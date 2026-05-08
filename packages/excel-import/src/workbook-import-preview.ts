import type { ImportedWorkbookSheetPreview } from './workbook-import-helpers.js'
import type { WorkbookImportContentType } from './workbook-import-content-types.js'

export interface ImportedWorkbookPreview {
  fileName: string
  contentType: WorkbookImportContentType
  fileSizeBytes: number
  workbookName: string
  sheetCount: number
  sheets: readonly ImportedWorkbookSheetPreview[]
  warnings: readonly string[]
}

export function createWorkbookPreview(input: {
  contentType: WorkbookImportContentType
  fileName: string
  fileSizeBytes: number
  workbookName: string
  sheets: readonly ImportedWorkbookSheetPreview[]
  warnings: readonly string[]
}): ImportedWorkbookPreview {
  return {
    fileName: input.fileName,
    contentType: input.contentType,
    fileSizeBytes: input.fileSizeBytes,
    workbookName: input.workbookName,
    sheetCount: input.sheets.length,
    sheets: input.sheets,
    warnings: [...input.warnings],
  }
}
