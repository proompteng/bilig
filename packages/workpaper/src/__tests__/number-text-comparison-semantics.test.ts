import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, ref: string): CellValue {
  const address = workbook.simpleCellAddressFromString(ref)
  if (!address) {
    throw new Error(`Expected ${ref} to resolve`)
  }
  return workbook.getCellValue(address)
}

function expectString(value: CellValue, expected: string): void {
  expect(value).toMatchObject({ tag: ValueTag.String, value: expected })
}

function buildForecastWorkbook(useColumnIndex: boolean): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Forecast: [
        ['Label', 'Value', 'Horizon', 'Result'],
        ['Literal number compared with text blank', 10, ' ', '=IF(B2<" ","beyond horizon","calculate")'],
        ['Cell number compared with cell text blank', 10, ' ', '=IF(B3<C3,"beyond horizon","calculate")'],
        ['Forecast horizon blank suppresses future rate', 10, ' ', '=IF(B4<C4," ",0.1092458427510494)'],
        ['Future amount should stay blank after horizon', 10, ' ', '=IF(B5<C5," ",117010.0245753175)'],
        ['Numeric-looking text still sorts after numbers', 999, '1', '=IF(B6<C6,"beyond horizon","calculate")'],
        ['Empty text remains distinct from numeric zero', 0, '', '=IF(B7=C7,"blank","not blank")'],
      ] satisfies TestCell[][],
    },
    { maxRows: 12, maxColumns: 8, useColumnIndex },
  )
}

describe('number-vs-text comparison semantics', () => {
  it.each([false, true])('keeps forecast-horizon blanks when numbers compare with text using useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildForecastWorkbook(useColumnIndex)

    expectString(cellValue(workbook, 'Forecast!D2'), 'beyond horizon')
    expectString(cellValue(workbook, 'Forecast!D3'), 'beyond horizon')
    expectString(cellValue(workbook, 'Forecast!D4'), ' ')
    expectString(cellValue(workbook, 'Forecast!D5'), ' ')
    expectString(cellValue(workbook, 'Forecast!D6'), 'beyond horizon')
    expectString(cellValue(workbook, 'Forecast!D7'), 'not blank')

    workbook.dispose()
  })
})
