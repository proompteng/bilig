import { FormulaMode } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { compileFormula } from '../compiler.js'

describe('AVERAGE axis aggregate compilation', () => {
  it('binds whole-column AVERAGE references to the native aggregate path', () => {
    const sameSheet = compileFormula('AVERAGE(V:V)')

    expect(sameSheet.mode).toBe(FormulaMode.WasmFastPath)
    expect(sameSheet.symbolicRanges).toEqual(['V:V'])
    expect(sameSheet.parsedSymbolicRanges?.[0]).toMatchObject({
      address: 'V:V',
      refKind: 'cols',
      startAddress: 'V',
      endAddress: 'V',
      startCol: 21,
      endCol: 21,
    })

    const crossSheet = compileFormula('AVERAGE(Data!$A:$A)')

    expect(crossSheet.mode).toBe(FormulaMode.WasmFastPath)
    expect(crossSheet.symbolicRanges).toEqual(['Data!A:A'])
    expect(crossSheet.parsedSymbolicRanges?.[0]).toMatchObject({
      address: 'Data!$A:$A',
      sheetName: 'Data',
      refKind: 'cols',
      startAddress: 'A',
      endAddress: 'A',
      startCol: 0,
      endCol: 0,
      startColAbsolute: true,
      endColAbsolute: true,
    })
  })

  it('keeps bounded AVERAGE ranges on the JS path that ignores blanks', () => {
    const bounded = compileFormula('ROUND(AVERAGE(E2:E315),2)')

    expect(bounded.mode).toBe(FormulaMode.JsOnly)
    expect(bounded.symbolicRanges).toEqual([])
  })
})
