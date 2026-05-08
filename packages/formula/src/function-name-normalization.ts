const EXCEL_COMPATIBILITY_FUNCTION_PREFIXES = ['_XLFN.', '_XLWS.'] as const

export function normalizeFormulaFunctionName(name: string): string {
  let normalized = name.trim().toUpperCase()
  let changed = true

  while (changed) {
    changed = false
    for (const prefix of EXCEL_COMPATIBILITY_FUNCTION_PREFIXES) {
      if (normalized.startsWith(prefix)) {
        normalized = normalized.slice(prefix.length)
        changed = true
        break
      }
    }
  }

  return normalized
}
