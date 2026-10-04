export const BILIG_CONTENT_TYPE = 'application/vnd.bilig.workbook+json'
export const XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
export const XLSM_CONTENT_TYPE = 'application/vnd.ms-excel.sheet.macroenabled.12'
export const XLSB_CONTENT_TYPE = 'application/vnd.ms-excel.sheet.binary.macroenabled.12'
export const LEGACY_XLS_CONTENT_TYPE = 'application/vnd.ms-excel'
export const CSV_CONTENT_TYPE = 'text/csv'
export const WORKBOOK_IMPORT_CONTENT_TYPES = [
  XLSX_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSB_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  CSV_CONTENT_TYPE,
  BILIG_CONTENT_TYPE,
] as const

export type WorkbookImportContentType = (typeof WORKBOOK_IMPORT_CONTENT_TYPES)[number]

export function normalizeWorkbookImportContentType(contentType: string): WorkbookImportContentType | null {
  const mediaType = contentType.split(';', 1)[0]?.trim().toLowerCase() ?? ''
  return WORKBOOK_IMPORT_CONTENT_TYPES.find((candidate) => candidate === mediaType) ?? null
}
