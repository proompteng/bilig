import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { selectCellSnapshot, selectMetrics, selectSelectionState, selectViewportCells } from '../selectors.js'

describe('core selectors', () => {
  it('reads cell, metrics, selection, and viewport snapshots from the engine', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'selectors' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setSelection('Sheet1', 'A1')

    expect(selectCellSnapshot(engine, 'Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })
    expect(selectMetrics(engine).batchId).toBeGreaterThan(0)
    expect(selectSelectionState(engine)).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
      anchorAddress: 'A1',
      range: { startAddress: 'A1', endAddress: 'A1' },
      editMode: 'idle',
    })
    expect(
      selectViewportCells(engine, 'Sheet1', {
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      }),
    ).toEqual([
      expect.objectContaining({
        sheetName: 'Sheet1',
        address: 'A1',
      }),
    ])
  })

  it('returns no viewport cells when the sheet does not exist', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'selectors-missing-sheet' })
    await engine.ready()

    expect(
      selectViewportCells(engine, 'Missing', {
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      }),
    ).toEqual([])
  })
})
