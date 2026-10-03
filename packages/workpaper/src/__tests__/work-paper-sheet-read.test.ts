import { describe, expect, it } from 'vitest'
import { readRuntimeSnapshot } from '@bilig/core'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { WorkPaper } from '../index.js'
import {
  buildWorkPaperDenseRange,
  collectSerializedWorkPaperSheets,
  collectWorkPaperSheetsByName,
  readWorkPaperSheetRange,
  workPaperRangeForSheetDimensions,
} from '../work-paper-sheet-read.js'

describe('work paper sheet read helpers', () => {
  it('builds dense ranges by visiting each address', () => {
    expect(
      buildWorkPaperDenseRange(
        {
          start: { sheet: 2, row: 3, col: 4 },
          end: { sheet: 2, row: 4, col: 6 },
        },
        (address) => `${address.sheet}:${address.row}:${address.col}`,
      ),
    ).toEqual([
      ['2:3:4', '2:3:5', '2:3:6'],
      ['2:4:4', '2:4:5', '2:4:6'],
    ])
  })

  it('maps non-empty sheet dimensions to a full-sheet range', () => {
    expect(workPaperRangeForSheetDimensions(3, { width: 2, height: 4 })).toEqual({
      start: { sheet: 3, row: 0, col: 0 },
      end: { sheet: 3, row: 3, col: 1 },
    })
    expect(workPaperRangeForSheetDimensions(3, { width: 0, height: 4 })).toBeUndefined()
    expect(readWorkPaperSheetRange({ sheetId: 3, dimensions: { width: 0, height: 0 }, readRange: () => [[1]] })).toEqual([])
  })

  it('preserves blank cells inside a sparse dense sheet range', () => {
    expect(
      readWorkPaperSheetRange({
        sheetId: 3,
        dimensions: { width: 3, height: 2 },
        readRange: (range) => {
          expect(range).toEqual({
            start: { sheet: 3, row: 0, col: 0 },
            end: { sheet: 3, row: 1, col: 2 },
          })
          return [
            [1, null, 3],
            [null, 5, null],
          ]
        },
      }),
    ).toEqual([
      [1, null, 3],
      [null, 5, null],
    ])
  })

  it('reads range values and serialized formulas after input edits', () => {
    const workbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Win rate', 0.25],
        ['Leads', 40],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Expected customers', '=Inputs!B2*Inputs!B3'],
      ],
    })
    const inputsId = workbook.getSheetId('Inputs')!
    const summaryId = workbook.getSheetId('Summary')!
    const summaryRange = {
      start: { sheet: summaryId, row: 0, col: 0 },
      end: { sheet: summaryId, row: 1, col: 1 },
    }

    expect(workbook.getRangeSerialized(summaryRange)).toEqual([
      ['Metric', 'Value'],
      ['Expected customers', '=Inputs!B2*Inputs!B3'],
    ])
    expect(workbook.getRangeValues(summaryRange)[1]?.[1]).toEqual({ tag: ValueTag.Number, value: 10 })

    workbook.setCellContents({ sheet: inputsId, row: 1, col: 1 }, 0.5)

    expect(workbook.getRangeSerialized(summaryRange)).toEqual([
      ['Metric', 'Value'],
      ['Expected customers', '=Inputs!B2*Inputs!B3'],
    ])
    expect(workbook.getRangeValues(summaryRange)[1]?.[1]).toEqual({ tag: ValueTag.Number, value: 20 })
  })

  it('collects sheet records by name and preserves runtime snapshots on serialized output', () => {
    const sheets = [
      { id: 1, name: 'Sheet1' },
      { id: 2, name: 'Sheet2' },
    ]
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'Book' },
      sheets: [],
    }

    expect(collectWorkPaperSheetsByName(sheets, (sheet) => sheet.id)).toEqual({ Sheet1: 1, Sheet2: 2 })
    const serialized = collectSerializedWorkPaperSheets({
      sheets,
      readSheet: (sheet) => [[sheet.name]],
      runtimeSnapshot: snapshot,
    })

    expect(serialized).toEqual({ Sheet1: [['Sheet1']], Sheet2: [['Sheet2']] })
    expect(readRuntimeSnapshot(serialized)).toBe(snapshot)
  })
})
