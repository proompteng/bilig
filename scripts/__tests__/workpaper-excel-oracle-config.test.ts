import { describe, expect, it } from 'vitest'

import { resolveExcelOracleDisabled } from '../workpaper-excel-oracle-config.ts'

describe('WorkPaper Excel oracle config', () => {
  it('defaults to allowing Excel oracle automation checks', () => {
    expect(resolveExcelOracleDisabled({})).toBe(false)
    expect(resolveExcelOracleDisabled({ BILIG_EXCEL_ORACLE_DISABLE: '' })).toBe(false)
  })

  it('accepts explicit boolean values for disabling Excel oracle automation', () => {
    expect(resolveExcelOracleDisabled({ BILIG_EXCEL_ORACLE_DISABLE: '1' })).toBe(true)
    expect(resolveExcelOracleDisabled({ BILIG_EXCEL_ORACLE_DISABLE: 'true' })).toBe(true)
    expect(resolveExcelOracleDisabled({ BILIG_EXCEL_ORACLE_DISABLE: '0' })).toBe(false)
    expect(resolveExcelOracleDisabled({ BILIG_EXCEL_ORACLE_DISABLE: 'false' })).toBe(false)
  })

  it('rejects malformed Excel oracle disable flags', () => {
    expect(() => resolveExcelOracleDisabled({ BILIG_EXCEL_ORACLE_DISABLE: 'yes' })).toThrow(
      'BILIG_EXCEL_ORACLE_DISABLE must be "1", "true", "0", or "false" when set, got yes',
    )
  })
})
