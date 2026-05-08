import * as XLSX from 'xlsx'
import { describe, expect, it } from 'vitest'

import { importXlsx } from '../index.js'

describe('github issue #42 xlsx literal error cell import', () => {
  it('preserves literal Excel error cells as display text instead of numeric error codes', () => {
    const imported = importXlsx(buildLiteralErrorWorkbookBytes(), 'literal-error-cells.xlsx')
    const cells = new Map(imported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell.value]) ?? [])

    expect(cells.get('A2')).toBe('#N/A')
    expect(cells.get('B2')).toBe('#DIV/0!')
    expect(cells.get('C2')).toBe('#REF!')
    expect(cells.get('D2')).toBe('#VALUE!')
    expect(cells.get('E2')).toBe('#NAME?')
    expect(cells.get('F2')).toBe('#NUM!')

    expect(imported.warnings).toEqual([])
    expect(imported.preview.sheets[0]?.previewRows[1]).toEqual(['#N/A', '#DIV/0!', '#REF!', '#VALUE!', '#NAME?', '#NUM!'])
  })
})

function buildLiteralErrorWorkbookBytes(): Uint8Array {
  const sheet = XLSX.utils.aoa_to_sheet([['NA', 'DIV', 'REF', 'VALUE', 'NAME', 'NUM']])
  sheet['A2'] = { t: 'e', v: 42 }
  sheet['B2'] = { t: 'e', v: 7 }
  sheet['C2'] = { t: 'e', v: 23 }
  sheet['D2'] = { t: 'e', v: 15 }
  sheet['E2'] = { t: 'e', v: 29 }
  sheet['F2'] = { t: 'e', v: 36 }
  sheet['!ref'] = 'A1:F2'

  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, sheet, 'Errors')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}
