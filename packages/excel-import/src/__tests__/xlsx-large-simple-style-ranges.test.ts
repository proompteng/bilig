import type { CellStyleRecord } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { buildLargeSimpleStyleRanges } from '../xlsx-large-simple-style-ranges.js'
import { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena, type ImportedWorksheetCellScan } from '../xlsx-large-simple-arena.js'

describe('large simple style range materialization', () => {
  it('streams row-major style records directly into coalesced ranges', () => {
    const styleIndexes = new ImportedWorksheetStyleIndexArena()
    styleIndexes.add(0, 0, 1)
    styleIndexes.add(0, 1, 1)
    styleIndexes.add(1, 0, 1)
    const styleCatalog = new Map<string, CellStyleRecord>()
    const styleRanges = buildLargeSimpleStyleRanges('Data', scanWithStyleIndexes(styleIndexes), stylesByIndex(), styleCatalog)
    const styleId = styleRanges[0]?.styleId

    expect(styleRanges).toEqual([
      { range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B1' }, styleId },
      { range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'A2' }, styleId },
    ])
    expect([...styleCatalog.values()]).toEqual([{ id: styleId, fill: { backgroundColor: '#ffcc00' } }])
  })

  it('sorts and compacts ranges when style records arrive out of row order', () => {
    const styleIndexes = new ImportedWorksheetStyleIndexArena()
    styleIndexes.add(1, 0, 1)
    styleIndexes.add(1, 1, 1)
    styleIndexes.add(0, 0, 1)
    styleIndexes.add(0, 1, 1)
    const styleRanges = buildLargeSimpleStyleRanges('Data', scanWithStyleIndexes(styleIndexes), stylesByIndex(), new Map())
    const styleId = styleRanges[0]?.styleId

    expect(styleRanges).toEqual([{ range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' }, styleId }])
  })
})

function stylesByIndex(): Map<number, Omit<CellStyleRecord, 'id'>> {
  return new Map([[1, { fill: { backgroundColor: '#ffcc00' } }]])
}

function scanWithStyleIndexes(styleIndexes: ImportedWorksheetStyleIndexArena): ImportedWorksheetCellScan {
  return {
    arena: new ImportedWorkbookArena(),
    sheetIndex: 0,
    richTextCells: [],
    styleIndexes,
    blankStyleCellCount: 0,
    cellCount: 0,
    valueCellCount: 0,
    formulaCellCount: 0,
    mergeCount: 0,
    conditionalFormatCount: 0,
    tableCount: 0,
    rowCount: 0,
    columnCount: 0,
    usedRange: null,
  }
}
