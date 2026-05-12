import { describe, expect, it } from 'vitest'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

describe('engine defined-name implicit intersection', () => {
  it('intersects row-vector defined names before scalar logical tests', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'defined-name-logical-implicit-intersection',
        metadata: {
          definedNames: [
            { name: 'root1', value: { kind: 'range-ref', sheetName: 'Main', startAddress: 'B51', endAddress: 'D51' } },
            { name: 'root2', value: { kind: 'range-ref', sheetName: 'Main', startAddress: 'B52', endAddress: 'D52' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Main',
          order: 0,
          cells: [
            { address: 'B51', value: 1 },
            { address: 'C51', value: 2 },
            { address: 'D51', value: 0.5 },
            { address: 'B52', value: 10 },
            { address: 'C52', value: 20 },
            { address: 'D52', value: 30 },
            { address: 'B54', formula: 'IF(AND(root1<=1,root1>=0),root1,root2)' },
            { address: 'C54', formula: 'IF(AND(root1<=1,root1>=0),root1,root2)' },
            { address: 'D54', formula: 'IF(AND(root1<=1,root1>=0),root1,root2)' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'defined-name-logical-implicit-intersection' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Main', 'B54')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Main', 'C54')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Main', 'D54')).toEqual({ tag: ValueTag.Number, value: 0.5 })

    engine.setCellValue('Main', 'C51', 0.25)
    expect(engine.getCellValue('Main', 'C54')).toEqual({ tag: ValueTag.Number, value: 0.25 })

    engine.setCellValue('Main', 'D51', -1)
    expect(engine.getCellValue('Main', 'D54')).toEqual({ tag: ValueTag.Number, value: 30 })
  })

  it('intersects row-vector lookup names and treats matched blanks as zero', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'defined-name-lookup-blank-return',
        metadata: {
          definedNames: [
            { name: 'KeyRate_last', value: { kind: 'range-ref', sheetName: 'Main', startAddress: 'B1', endAddress: 'D1' } },
            { name: 'tbl_last', value: { kind: 'range-ref', sheetName: 'Main', startAddress: 'B4', endAddress: 'F6' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Main',
          order: 0,
          cells: [
            { address: 'B1', value: 1 },
            { address: 'C1', value: 2 },
            { address: 'D1', value: 3 },
            { address: 'B4', value: 1 },
            { address: 'F4', value: 10 },
            { address: 'B5', value: 2 },
            { address: 'B6', value: 3 },
            { address: 'F6', value: 30 },
            { address: 'C10', formula: 'VLOOKUP(KeyRate_last,tbl_last,5,0)' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'defined-name-lookup-blank-return' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Main', 'C10')).toEqual({ tag: ValueTag.Number, value: 0 })

    engine.setCellValue('Main', 'F5', 42)
    expect(engine.getCellValue('Main', 'C10')).toEqual({ tag: ValueTag.Number, value: 42 })

    engine.setCellValue('Main', 'C1', 3)
    expect(engine.getCellValue('Main', 'C10')).toEqual({ tag: ValueTag.Number, value: 30 })
  })
})
