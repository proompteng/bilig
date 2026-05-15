declare module 'xlsx-calc' {
  interface XlsxCalc {
    (workbook: unknown, options?: { continue_after_error?: boolean; log_error?: boolean }): void
    import_functions(functions: Record<string, (...args: unknown[]) => unknown>): void
  }

  const xlsxCalc: XlsxCalc
  export default xlsxCalc
}
