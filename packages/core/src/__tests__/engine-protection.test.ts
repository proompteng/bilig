import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'

describe('SpreadsheetEngine protections', () => {
  it('roundtrips sheet and range protections through snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'protection-roundtrip' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setSheetProtection({ sheetName: 'Sheet1', hideFormulas: true })
    engine.setRangeProtection({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      },
      hideFormulas: true,
    })

    const snapshot = engine.exportSnapshot()
    const metadata = snapshot.sheets.find((sheet) => sheet.name === 'Sheet1')?.metadata
    expect(metadata?.sheetProtection).toEqual({
      sheetName: 'Sheet1',
      hideFormulas: true,
    })
    expect(metadata?.protectedRanges).toEqual([
      {
        id: 'protect-a1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        hideFormulas: true,
      },
    ])

    const restored = new SpreadsheetEngine({ workbookName: 'protection-roundtrip-restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)
    expect(restored.getSheetProtection('Sheet1')).toEqual({
      sheetName: 'Sheet1',
      hideFormulas: true,
    })
    expect(restored.getRangeProtections('Sheet1')).toEqual([
      {
        id: 'protect-a1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        hideFormulas: true,
      },
    ])
  })

  it('blocks writes to protected sheets and ranges while allowing explicit unprotect', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'protection-enforcement' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeProtection({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      },
    })

    expect(() => engine.setCellValue('Sheet1', 'A1', 7)).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.setCellValue('Sheet1', 'C3', 7)).not.toThrow()

    engine.setSheetProtection({ sheetName: 'Sheet1', hideFormulas: true })
    expect(() => engine.setCellValue('Sheet1', 'D4', 9)).toThrow(/Workbook protection blocks this change/)

    expect(engine.clearSheetProtection('Sheet1')).toBe(true)
    expect(engine.deleteRangeProtection('protect-a1')).toBe(true)
    expect(() => engine.setCellValue('Sheet1', 'A1', 7)).not.toThrow()
  })

  it('reports missing range protections after deletion through the public API', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'protection-delete-missing' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeProtection({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      },
      hideFormulas: true,
    })

    expect(engine.getRangeProtection('protect-a1')).toMatchObject({ id: 'protect-a1' })
    expect(engine.deleteRangeProtection('protect-a1')).toBe(true)
    expect(engine.getRangeProtection('protect-a1')).toBeUndefined()
    expect(engine.getRangeProtections('Sheet1')).toEqual([])
    expect(engine.deleteRangeProtection('protect-a1')).toBe(false)
  })
})
