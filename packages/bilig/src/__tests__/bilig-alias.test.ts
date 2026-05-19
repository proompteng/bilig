import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { exportXlsx, importXlsx } from '../xlsx.js'

describe('bilig package alias', () => {
  it('exposes the WorkPaper runtime and XLSX helpers through the short package name', () => {
    const workbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })
    const inputs = workbook.getSheetId('Inputs')
    const summary = workbook.getSheetId('Summary')

    expect(inputs).toBeTypeOf('number')
    expect(summary).toBeTypeOf('number')
    workbook.setCellContents({ sheet: inputs!, row: 1, col: 1 }, 48)
    workbook.setCellContents({ sheet: inputs!, row: 2, col: 1 }, 1500)
    expect(readNumber(workbook.getCellValue({ sheet: summary!, row: 1, col: 1 }))).toBe(72_000)

    const exported = exportXlsx(workbook.exportSnapshot())
    const imported = importXlsx(exported, 'pricing.xlsx')
    expect(imported.sheetNames).toEqual(['Inputs', 'Summary'])

    workbook.dispose()
  })
})

function readNumber(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'number') {
    return value.value
  }
  throw new Error(`Expected numeric cell value, received ${JSON.stringify(value)}`)
}
