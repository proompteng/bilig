import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { WorkPaper, type WorkPaperCellAddress, type WorkPaperChange } from '../index.js'
import { forceMaterializeTrackedIndexChanges, hasDeferredTrackedIndexChanges } from '../tracked-cell-index-changes.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function columnLabel(index: number): string {
  let value = index + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

function buildFormulaChainRow(downstreamCount: number): unknown[] {
  return Array.from({ length: downstreamCount + 1 }, (_value, col) => (col === 0 ? 1 : `=${columnLabel(col - 1)}1+1`))
}

function buildFormulaFanoutRow(downstreamCount: number): unknown[] {
  return Array.from({ length: downstreamCount + 1 }, (_value, col) => (col === 0 ? 1 : `=$A$1+${String(col)}`))
}

describe('existing numeric tracked fast path', () => {
  it('updates common scalar leaf formulas through the inline mutation path', () => {
    const cases = [
      {
        name: 'arithmetic',
        sheet: [[100, 20, null, '=A1+B1*2']],
        edit: { row: 0, col: 0, value: 101 },
        expected: { tag: ValueTag.Number, value: 141 },
      },
      {
        name: 'branching',
        sheet: [[100, null, null, '=IF(A1>0,"yes","no")']],
        edit: { row: 0, col: 0, value: -1 },
        expected: { tag: ValueTag.String, value: 'no' },
      },
      {
        name: 'financial',
        sheet: [[100, null, null, '=PMT(A1/12,A2,A3)'], [12], [1000]],
        edit: { row: 0, col: 0, value: 101 },
        expected: { tag: ValueTag.Number, value: -8416.66666668398 },
      },
      {
        name: 'math',
        sheet: [[100, null, null, '=ROUND(SQRT(A1),2)']],
        edit: { row: 0, col: 0, value: 121 },
        expected: { tag: ValueTag.Number, value: 11 },
      },
      {
        name: 'math-round-decimal-tie',
        sheet: [[0, 0.1, null, '=ROUND(A1/1000/1000/B1,0)*B1']],
        edit: { row: 0, col: 0, value: 16150000 },
        expected: { tag: ValueTag.Number, value: 16.2 },
      },
      {
        name: 'text-concat',
        sheet: [['foo', 'bar', null, '=CONCATENATE(A1,"-",B1)']],
        edit: { row: 0, col: 0, value: 'baz' },
        expected: { tag: ValueTag.String, value: 'baz-bar' },
      },
      {
        name: 'text-length',
        sheet: [['foo', 'bar', null, '=LEN(A1)+LEN(B1)']],
        edit: { row: 0, col: 1, value: 'quux' },
        expected: { tag: ValueTag.Number, value: 7 },
      },
      {
        name: 'minmax',
        sheet: [[100, 20, 5, '=MIN(A1,B1,C1)+MAX(A1,B1,C1)']],
        edit: { row: 0, col: 2, value: 200 },
        expected: { tag: ValueTag.Number, value: 220 },
      },
    ] as const

    for (const testCase of cases) {
      const workbook = WorkPaper.buildFromSheets({ Bench: testCase.sheet })
      const sheetId = workbook.getSheetId('Bench')!

      expect(Reflect.get(workbook, 'engineEventsAttached'), testCase.name).toBe(false)
      const changes = workbook.setCellContents(cell(sheetId, testCase.edit.row, testCase.edit.col), testCase.edit.value)
      const result = workbook.getCellValue(cell(sheetId, 0, 3))

      expect(changes, testCase.name).toHaveLength(2)
      expect(Reflect.get(workbook, 'engineEventsAttached'), testCase.name).toBe(false)
      expect(Reflect.get(workbook, 'setCellContentsRuntimeCache'), testCase.name).toBeUndefined()
      expect(Reflect.get(workbook, 'existingNumericFastPathRuntimeCache'), testCase.name).toBeUndefined()
      expect(workbook.getStats().lastMetrics, testCase.name).toMatchObject({
        dirtyFormulaCount: 0,
        jsFormulaCount: 0,
        wasmFormulaCount: 0,
      })
      if (testCase.expected.tag === ValueTag.Number) {
        expect(result.tag, testCase.name).toBe(ValueTag.Number)
        expect(result.tag === ValueTag.Number ? result.value : Number.NaN, testCase.name).toBeCloseTo(testCase.expected.value, 9)
      } else {
        expect(result, testCase.name).toMatchObject(testCase.expected)
      }
      workbook.dispose()
    }
  })

  it('returns tiny no-listener direct scalar changes eagerly', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[1, 2, '=A1+B1']],
    })
    const sheetId = workbook.getSheetId('Bench')!

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(changes).toHaveLength(2)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(false)
    expect(Reflect.get(workbook, 'engineEvents').hasPendingLazyChanges).toBe(false)
    expect(changes[0]).toMatchObject({
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 9 },
    })
    expect(changes[1]).toMatchObject({
      a1: 'C1',
      newValue: { tag: ValueTag.Number, value: 11 },
    })
  })

  it('keeps single-cell direct literal changes eager', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[1, 2]],
    })
    const sheetId = workbook.getSheetId('Bench')!

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(changes).toHaveLength(1)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(false)
    expect(changes[0]).toMatchObject({
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 9 },
    })
  })

  it('returns large no-listener direct changes lazily and detaches before later writes', () => {
    const downstreamCount = 320
    const workbook = WorkPaper.buildFromSheets({
      Bench: [buildFormulaChainRow(downstreamCount)],
    })
    const sheetId = workbook.getSheetId('Bench')!

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(Array.isArray(changes)).toBe(true)
    expect(changes).toHaveLength(downstreamCount + 1)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)

    workbook.setCellContents(cell(sheetId, 0, 0), 10)

    expect(changes[0]).toMatchObject({
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 9 },
    })
    expect(changes[downstreamCount]).toMatchObject({
      a1: `${columnLabel(downstreamCount)}1`,
      newValue: { tag: ValueTag.Number, value: 9 + downstreamCount },
    })
    expect(forceMaterializeTrackedIndexChanges(changes)).toBe(true)
  })

  it('keeps large direct scalar fanout tracked events lazy', () => {
    const downstreamCount = 320
    const workbook = WorkPaper.buildFromSheets({
      Bench: [buildFormulaFanoutRow(downstreamCount)],
    })
    const sheetId = workbook.getSheetId('Bench')!

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(changes).toHaveLength(downstreamCount + 1)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    expect(Reflect.get(workbook, 'engineEventsAttached')).toBe(false)
    expect(Reflect.get(workbook, 'engineEvents').hasPendingLazyChanges).toBe(true)

    workbook.setCellContents(cell(sheetId, 0, 0), 10)

    expect(changes[0]).toMatchObject({
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 9 },
    })
    expect(changes[downstreamCount]).toMatchObject({
      a1: `${columnLabel(downstreamCount)}1`,
      newValue: { tag: ValueTag.Number, value: 9 + downstreamCount },
    })
    expect(forceMaterializeTrackedIndexChanges(changes)).toBe(true)
  })

  it('keeps valuesUpdated payloads lazy for large direct changes', () => {
    const downstreamCount = 320
    const workbook = WorkPaper.buildFromSheets({
      Bench: [buildFormulaChainRow(downstreamCount)],
    })
    const sheetId = workbook.getSheetId('Bench')!
    const events: WorkPaperChange[][] = []
    workbook.on('valuesUpdated', (changes) => {
      events.push(changes)
    })

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(events).toHaveLength(1)
    expect(events[0]).toBe(changes)
    expect(changes).toHaveLength(downstreamCount + 1)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)

    workbook.setCellContents(cell(sheetId, 0, 0), 10)

    expect(changes[0]).toMatchObject({
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 9 },
    })
    expect(changes[downstreamCount]).toMatchObject({
      a1: `${columnLabel(downstreamCount)}1`,
      newValue: { tag: ValueTag.Number, value: 9 + downstreamCount },
    })
    expect(forceMaterializeTrackedIndexChanges(changes)).toBe(true)
  })

  it('returns cross-sheet direct scalar fanout changes lazily with a stable literal prefix', () => {
    const downstreamCount = 320
    const workbook = WorkPaper.buildFromSheets({
      Data: [[1]],
      Summary: Array.from({ length: downstreamCount }, (_value, row) => [`=Data!$A$1+${row + 1}`]),
    })
    const dataSheetId = workbook.getSheetId('Data')!
    const summarySheetId = workbook.getSheetId('Summary')!

    const changes = workbook.setCellContents(cell(dataSheetId, 0, 0), 9)

    expect(Array.isArray(changes)).toBe(true)
    expect(changes).toHaveLength(downstreamCount + 1)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)

    workbook.setCellContents(cell(dataSheetId, 0, 0), 10)

    expect(changes[0]).toMatchObject({
      address: { sheet: dataSheetId, row: 0, col: 0 },
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 9 },
    })
    expect(changes[1]).toMatchObject({
      address: { sheet: summarySheetId, row: 0, col: 0 },
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 10 },
    })
    expect(changes[downstreamCount]).toMatchObject({
      address: { sheet: summarySheetId, row: downstreamCount - 1, col: 0 },
      a1: `A${downstreamCount}`,
      newValue: { tag: ValueTag.Number, value: 9 + downstreamCount },
    })
    expect(forceMaterializeTrackedIndexChanges(changes)).toBe(true)
  })
})
