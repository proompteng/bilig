import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'

describe('SpreadsheetEngine data validations', () => {
  it('roundtrips data validation metadata through snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-roundtrip' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setDataValidation({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B4',
      },
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: false,
      showDropdown: true,
      errorStyle: 'stop',
      errorTitle: 'Status required',
      errorMessage: 'Pick Draft or Final.',
    })

    const snapshot = engine.exportSnapshot()
    expect(snapshot.sheets.find((sheet) => sheet.name === 'Sheet1')?.metadata?.validations).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
        showDropdown: true,
        errorStyle: 'stop',
        errorTitle: 'Status required',
        errorMessage: 'Pick Draft or Final.',
      },
    ])

    const restored = new SpreadsheetEngine({ workbookName: 'validation-roundtrip-restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
        showDropdown: true,
        errorStyle: 'stop',
        errorTitle: 'Status required',
        errorMessage: 'Pick Draft or Final.',
      },
    ])
  })

  it('rewrites and restores data validation ranges across structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-structural' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setDataValidation({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B4',
      },
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: false,
    })

    engine.insertRows('Sheet1', 1, 1)
    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B3',
          endAddress: 'B5',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
      },
    ])

    engine.deleteColumns('Sheet1', 0, 1)
    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A3',
          endAddress: 'A5',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
      },
    ])

    expect(engine.undo()).toBe(true)
    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B3',
          endAddress: 'B5',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
      },
    ])

    expect(engine.undo()).toBe(true)
    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
      },
    ])
  })
})
