import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'

import { recalculateExceljsBuffer, recalculateExceljsWorkbook } from '../index.js'

describe('exceljs-formula-recalc', () => {
  it('mutates an ExcelJS workbook with recalculated formula results', async () => {
    const workbook = new ExcelJS.Workbook()
    const inputs = workbook.addWorksheet('Inputs')
    inputs.getCell('A1').value = 'Metric'
    inputs.getCell('B1').value = 'Value'
    inputs.getCell('A2').value = 'Units'
    inputs.getCell('B2').value = 40
    inputs.getCell('A3').value = 'Price'
    inputs.getCell('B3').value = 1200

    const summary = workbook.addWorksheet('Summary')
    summary.getCell('A1').value = 'Metric'
    summary.getCell('B1').value = 'Value'
    summary.getCell('A2').value = 'Revenue'
    summary.getCell('B2').value = {
      formula: 'Inputs!B2*Inputs!B3',
      result: 48_000,
    }

    const result = await recalculateExceljsWorkbook(workbook, {
      edits: [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ],
      reads: ['Summary!B2'],
    })

    expect(result.workbookMutated).toBe(true)
    expect(readNumber(result.reads['Summary!B2'])).toBe(72_000)
    expect(readExceljsFormulaResult(workbook.getWorksheet('Summary')?.getCell('B2').value)).toBe(72_000)
  })

  it('can recalculate bytes without mutating an ExcelJS workbook object', async () => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 2
    sheet.getCell('B1').value = 3
    sheet.getCell('C1').value = {
      formula: 'A1+B1',
      result: 5,
    }

    const input = await workbook.xlsx.writeBuffer()
    const result = recalculateExceljsBuffer(input, {
      edits: [{ target: 'Sheet1!A1', value: 9 }],
      reads: ['Sheet1!C1'],
    })

    expect(readNumber(result.reads['Sheet1!C1'])).toBe(12)
    expect(readExceljsFormulaResult(sheet.getCell('C1').value)).toBe(5)
  })
})

function readNumber(value: XlsxFormulaCellValue): number {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'number') {
    return value.value
  }
  throw new Error(`Expected numeric cell value, received ${JSON.stringify(value)}`)
}

function readExceljsFormulaResult(value: ExcelJS.CellValue | undefined): number {
  if (typeof value === 'object' && value !== null && 'result' in value && typeof value.result === 'number') {
    return value.result
  }
  throw new Error(`Expected ExcelJS formula result, received ${JSON.stringify(value)}`)
}

type XlsxFormulaCellValue = Awaited<ReturnType<typeof recalculateExceljsWorkbook>>['reads'][string]
