import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function cellChanges(changes: Array<{ kind: string; a1?: string }>): Set<string> {
  return new Set(changes.flatMap((change) => (change.kind === 'cell' && change.a1 ? [change.a1] : [])))
}

describe('WorkPaper INDEX/MATCH direct path', () => {
  it('updates exact text MATCH operands through the existing-literal direct path', () => {
    const rowCount = 128
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        ['Key', 'Value', '', 'key-2', `=MATCH(D1,A2:A${rowCount + 1},0)`],
        ...Array.from({ length: rowCount }, (_, index) => [`key-${index + 1}`, (index + 1) * 10]),
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 2 })

    workbook.resetPerformanceCounters()
    const changes = workbook.setCellContents(cell(sheetId, 0, 3), 'key-99')

    expect(cellChanges(changes)).toEqual(new Set(['D1', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 99 })
    expect(workbook.getStats().lastMetrics).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      changedCellPayloadsBuilt: 0,
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      exactIndexBuilds: 0,
      formulasBound: 0,
      topoRepairs: 0,
    })
  })

  it('updates approximate duplicate MATCH operands through the uniform direct path', () => {
    const rowCount = 128
    const lookupValue = Math.floor(rowCount / 5)
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        ['Key', 'Value', '', Math.floor(rowCount / 4), `=MATCH(D1,A2:A${rowCount + 1},1)`],
        ...Array.from({ length: rowCount }, (_, index) => {
          const key = Math.ceil((index + 1) / 2)
          return [key, key * 10]
        }),
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 4) * 2 })

    workbook.resetPerformanceCounters()
    const changes = workbook.setCellContents(cell(sheetId, 0, 3), lookupValue)

    expect(cellChanges(changes)).toEqual(new Set(['D1', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: lookupValue * 2 })
    expect(workbook.getStats().lastMetrics).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      approxIndexBuilds: 0,
      changedCellPayloadsBuilt: 0,
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      formulasBound: 0,
      lookupOwnerBuilds: 0,
      topoRepairs: 0,
    })
  })

  it('updates exact INDEX/MATCH operands without rebinding dynamic range dependencies', () => {
    const rowCount = 128
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [
          ['Key', 'Value', '', 'key-2', `=INDEX(B2:B${rowCount + 1},MATCH(D1,A2:A${rowCount + 1},0))`],
          ...Array.from({ length: rowCount }, (_, index) => [`key-${index + 1}`, (index + 1) * 10]),
        ],
      },
      { useColumnIndex: true },
    )
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 20 })

    workbook.resetPerformanceCounters()
    const lookupChanges = workbook.setCellContents(cell(sheetId, 0, 3), 'key-99')

    expect(cellChanges(lookupChanges)).toEqual(new Set(['D1', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 990 })
    expect(workbook.getStats().lastMetrics).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      cycleFormulaScans: 0,
      formulasBound: 0,
      topoRepairs: 0,
    })

    workbook.resetPerformanceCounters()
    const selectedValueChanges = workbook.setCellContents(cell(sheetId, 99, 1), 12_345)

    expect(cellChanges(selectedValueChanges)).toEqual(new Set(['B100', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 12_345 })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      formulasBound: 0,
      topoRepairs: 0,
    })
  })

  it('undoes suspended exact lookup-column numeric batches from the compact numeric inverse path', () => {
    const rowCount = 96
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [
          ['Key', 'Value', '', Math.floor(rowCount / 2), `=MATCH(D1,A2:A${rowCount + 1},0)`],
          ...Array.from({ length: rowCount }, (_, index) => [index + 1, (index + 1) * 10]),
        ],
      },
      { useColumnIndex: true },
    )
    const sheetId = workbook.getSheetId('Bench')!

    workbook.resetPerformanceCounters()
    workbook.suspendEvaluation()
    for (let index = 0; index < 40; index += 1) {
      const row = rowCount - index
      workbook.setCellContents(cell(sheetId, row, 0), row + 1_000)
    }
    const changes = workbook.resumeEvaluation()

    expect(changes).toHaveLength(40)
    expect(workbook.getPerformanceCounters()).toMatchObject({
      kernelSyncOnlyRecalcSkips: 1,
      formulasBound: 0,
      lookupOwnerBuilds: 0,
    })
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({
      tag: ValueTag.Number,
      value: Math.floor(rowCount / 2),
    })

    const undoChanges = workbook.undo()

    expect(undoChanges).toHaveLength(40)
    expect(workbook.getCellValue(cell(sheetId, rowCount, 0))).toEqual({ tag: ValueTag.Number, value: rowCount })
    expect(workbook.getCellValue(cell(sheetId, rowCount - 39, 0))).toEqual({ tag: ValueTag.Number, value: rowCount - 39 })
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({
      tag: ValueTag.Number,
      value: Math.floor(rowCount / 2),
    })
  })

  it('preserves first-match and missing-match INDEX/MATCH semantics on the direct path', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [
          ['Key', 'Value', '', 'dup', '=INDEX(B2:B6,MATCH(D1,A2:A6,0))'],
          ['alpha', 10],
          ['dup', 20],
          ['beta', 30],
          ['dup', 40],
          ['gamma', 50],
        ],
      },
      { useColumnIndex: true },
    )
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 20 })

    workbook.resetPerformanceCounters()
    const firstMatchChanges = workbook.setCellContents(cell(sheetId, 2, 1), 2_000)

    expect(cellChanges(firstMatchChanges)).toEqual(new Set(['B3', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 2_000 })
    expect(workbook.getStats().lastMetrics).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 1,
      wasmFormulaCount: 0,
    })

    workbook.resetPerformanceCounters()
    const laterDuplicateChanges = workbook.setCellContents(cell(sheetId, 4, 1), 4_000)

    expect(cellChanges(laterDuplicateChanges)).toEqual(new Set(['B5', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 2_000 })
    expect(workbook.getStats().lastMetrics).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 1,
      wasmFormulaCount: 0,
    })

    workbook.resetPerformanceCounters()
    const missingLookupChanges = workbook.setCellContents(cell(sheetId, 0, 3), 'missing')

    expect(cellChanges(missingLookupChanges)).toEqual(new Set(['D1', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Error, code: ErrorCode.NA })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      formulasBound: 0,
      topoRepairs: 0,
    })
  })
})

describe('WorkPaper approximate MATCH direct path', () => {
  it('binds uniform numeric approximate MATCH without building the sorted lookup index', () => {
    const rowCount = 128
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        ['Key', 'Value', '', Math.floor(rowCount / 2) + 0.5, `=MATCH(D1,A2:A${rowCount + 1},1)`],
        ...Array.from({ length: rowCount }, (_, index) => [index + 1, (index + 1) * 10]),
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      approxIndexBuilds: 0,
      lookupOwnerBuilds: 0,
    })

    workbook.resetPerformanceCounters()
    const changes = workbook.setCellContents(cell(sheetId, 0, 3), rowCount - 0.5)

    expect(cellChanges(changes)).toEqual(new Set(['D1', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: rowCount - 1 })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      approxIndexBuilds: 0,
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      lookupOwnerBuilds: 0,
    })
  })
})

describe('WorkPaper INDEX reference direct path', () => {
  it('updates row-index operands without rebinding dynamic INDEX range dependencies', () => {
    const rowCount = 128
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        ['Key', 'Value', '', 2, `=INDEX(A2:B${rowCount + 1},D1,2)`],
        ...Array.from({ length: rowCount }, (_, index) => [index + 1, (index + 1) * 10]),
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 20 })

    workbook.resetPerformanceCounters()
    const rowSelectorChanges = workbook.setCellContents(cell(sheetId, 0, 3), rowCount - 1)

    expect(cellChanges(rowSelectorChanges)).toEqual(new Set(['D1', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: (rowCount - 1) * 10 })
    expect(workbook.getStats().lastMetrics).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      changedCellPayloadsBuilt: 0,
      cycleFormulaScans: 0,
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      formulasBound: 0,
      topoRepairs: 0,
    })

    workbook.resetPerformanceCounters()
    const selectedValueChanges = workbook.setCellContents(cell(sheetId, rowCount - 1, 1), 98_765)

    expect(cellChanges(selectedValueChanges)).toEqual(new Set([`B${rowCount}`, 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 98_765 })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      formulasBound: 0,
      topoRepairs: 0,
    })
  })

  it('preserves INDEX reference scalar value and error semantics on the direct path', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        ['Key', 'Value', '', 2, '=INDEX(A2:B5,D1,2)'],
        [1, 'one'],
        [2, 'two'],
        [3, 'three'],
        [4, 'four'],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toMatchObject({ tag: ValueTag.String, value: 'two' })

    workbook.resetPerformanceCounters()
    const textResultChanges = workbook.setCellContents(cell(sheetId, 0, 3), 4)

    expect(cellChanges(textResultChanges)).toEqual(new Set(['D1', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toMatchObject({ tag: ValueTag.String, value: 'four' })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      formulasBound: 0,
      topoRepairs: 0,
    })

    workbook.resetPerformanceCounters()
    const outOfRangeChanges = workbook.setCellContents(cell(sheetId, 0, 3), 5)

    expect(cellChanges(outOfRangeChanges)).toEqual(new Set(['D1', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      changedCellPayloadsBuilt: 0,
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      formulasBound: 0,
      topoRepairs: 0,
    })

    workbook.resetPerformanceCounters()
    const nonNumericChanges = workbook.setCellContents(cell(sheetId, 0, 3), 'bad')

    expect(cellChanges(nonNumericChanges)).toEqual(new Set(['D1', 'E1']))
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      changedCellPayloadsBuilt: 0,
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      formulasBound: 0,
      topoRepairs: 0,
    })
  })
})
