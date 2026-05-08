export const XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
export const XLSM_CONTENT_TYPE = 'application/vnd.ms-excel.sheet.macroenabled.12'
export const XLSB_CONTENT_TYPE = 'application/vnd.ms-excel.sheet.binary.macroenabled.12'
export const LEGACY_XLS_CONTENT_TYPE = 'application/vnd.ms-excel'
export const CSV_CONTENT_TYPE = 'text/csv'
export const EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES = [
  XLSX_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSB_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
] as const
export type ExcelWorkbookImportContentType = (typeof EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES)[number]
export const WORKBOOK_IMPORT_CONTENT_TYPES = [...EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES, CSV_CONTENT_TYPE] as const
export type WorkbookImportContentType = (typeof WORKBOOK_IMPORT_CONTENT_TYPES)[number]

export function normalizeWorkbookImportContentType(contentType: string): WorkbookImportContentType | null {
  const mediaType = contentType.split(';', 1)[0]?.trim().toLowerCase() ?? ''
  switch (mediaType) {
    case XLSX_CONTENT_TYPE:
      return XLSX_CONTENT_TYPE
    case XLSM_CONTENT_TYPE:
      return XLSM_CONTENT_TYPE
    case XLSB_CONTENT_TYPE:
      return XLSB_CONTENT_TYPE
    case LEGACY_XLS_CONTENT_TYPE:
      return LEGACY_XLS_CONTENT_TYPE
    case CSV_CONTENT_TYPE:
      return CSV_CONTENT_TYPE
    default:
      return null
  }
}
