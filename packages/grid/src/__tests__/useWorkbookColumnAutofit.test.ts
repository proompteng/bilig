import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import { measureWorkbookColumnAutofit, type WorkbookColumnTextMeasurer } from '../useWorkbookColumnAutofit.js'

function createSnapshot(address: string, value: string): CellSnapshot {
  return {
    address,
    flags: 0,
    sheetName: 'Sheet1',
    value: { tag: ValueTag.String, value },
    version: 1,
  }
}

describe('measureWorkbookColumnAutofit', () => {
  it('measures sheet cells and optimistic editor seeds outside the render-state hook', () => {
    const engine: GridEngineLike = {
      workbook: {
        getSheet: () => ({
          grid: {
            forEachCellEntry(listener) {
              listener(1, 0, 2)
              listener(2, 1, 2)
              listener(3, 1, 3)
            },
          },
        }),
      },
      getCell: (_sheetName, address) => createSnapshot(address, address === 'C1' ? 'short' : 'medium'),
      getCellStyle: () => undefined,
      subscribeCells: () => () => undefined,
    }
    const measurer: WorkbookColumnTextMeasurer = {
      font: '',
      measureText: (text) => ({ width: text.length * 10 }),
    }

    expect(
      measureWorkbookColumnAutofit({
        columnIndex: 2,
        editorFontSize: '12px',
        engine,
        freezeRows: 0,
        getCellEditorSeed: (_sheetName, address) => (address === 'C2' ? 'optimistic wider value' : undefined),
        getVisibleRegion: () => ({
          freezeCols: 0,
          freezeRows: 0,
          range: { height: 10, width: 10, x: 0, y: 0 },
          tx: 0,
          ty: 0,
        }),
        headerFontStyle: '600 11px sans-serif',
        measurer,
        selectedCell: { col: 2, row: 1 },
        selectedCellSnapshot: createSnapshot('C2', 'selected'),
        sheetName: 'Sheet1',
      }),
    ).toBe('optimistic wider value'.length * 10 + 28)
  })
})
