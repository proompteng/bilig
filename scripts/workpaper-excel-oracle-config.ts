export function resolveExcelOracleDisabled(env: { BILIG_EXCEL_ORACLE_DISABLE?: string | undefined }): boolean {
  return parseStrictBooleanFlag(env.BILIG_EXCEL_ORACLE_DISABLE, 'BILIG_EXCEL_ORACLE_DISABLE', false)
}

function parseStrictBooleanFlag(value: string | undefined, name: string, fallback: boolean): boolean {
  if (value === undefined || value.length === 0) {
    return fallback
  }
  if (value === '1' || value === 'true') {
    return true
  }
  if (value === '0' || value === 'false') {
    return false
  }
  throw new Error(`${name} must be "1", "true", "0", or "false" when set, got ${value}`)
}
