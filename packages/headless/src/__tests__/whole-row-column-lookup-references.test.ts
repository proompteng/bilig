import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

function display(workpaper: WorkPaper, ref: string): string {
  const address = workpaper.simpleCellAddressFromString(ref)
  if (!address) {
    throw new Error(`Expected ${ref} to resolve`)
  }
  return workpaper.getCellDisplayValue(address)
}

describe('whole-row and whole-column lookup references', () => {
  it('evaluates imported MATCH and INDEX formulas using whole-row headers and whole-column data windows', () => {
    const workpaper = WorkPaper.buildFromSheets({
      Data: [
        ['Type', 'Created (UTC)'],
        ['charge', '2026-01-31'],
      ],
      Out: [
        [
          '=MATCH("Type",Data!$1:$1,0)',
          '=MATCH("Type",Data!$A$1:$ZZ$1,0)',
          '=INDEX(Data!$A:$ZZ,2,1)',
          '=INDEX(Data!$A$1:$ZZ$2,2,1)',
          '=IFERROR(INDEX(Data!$A:$ZZ,2,IFERROR(MATCH("Type",Data!$1:$1,0),MATCH("type",Data!$1:$1,0))),"")',
        ],
      ],
    })

    expect(display(workpaper, 'Out!A1')).toBe('1')
    expect(display(workpaper, 'Out!B1')).toBe('1')
    expect(display(workpaper, 'Out!C1')).toBe('charge')
    expect(display(workpaper, 'Out!D1')).toBe('charge')
    expect(display(workpaper, 'Out!E1')).toBe('charge')

    workpaper.dispose()
  })
})
