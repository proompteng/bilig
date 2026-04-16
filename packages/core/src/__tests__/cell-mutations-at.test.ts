import { describe, expect, it } from 'vitest'
import { WorkbookStore } from '../workbook-store.js'
import {
  cellMutationRefToEngineOp,
  cloneCellMutationAt,
  cloneCellMutationRef,
  countPotentialNewCellsForMutationRefs,
} from '../cell-mutations-at.js'

describe('cell mutations at', () => {
  it('clones all supported mutation shapes', () => {
    expect(
      cloneCellMutationAt({
        kind: 'setCellValue',
        row: 1,
        col: 2,
        value: 42,
      }),
    ).toEqual({
      kind: 'setCellValue',
      row: 1,
      col: 2,
      value: 42,
    })

    expect(
      cloneCellMutationAt({
        kind: 'setCellFormula',
        row: 3,
        col: 4,
        formula: 'A1*2',
      }),
    ).toEqual({
      kind: 'setCellFormula',
      row: 3,
      col: 4,
      formula: 'A1*2',
    })

    expect(
      cloneCellMutationAt({
        kind: 'clearCell',
        row: 5,
        col: 6,
      }),
    ).toEqual({
      kind: 'clearCell',
      row: 5,
      col: 6,
    })
  })

  it('clones mutation refs and converts them into engine ops', () => {
    const workbook = new WorkbookStore('cell-mutations')
    workbook.createSheet('Sheet1', 0)
    const sheetId = workbook.getSheet('Sheet1')!.id

    const valueRef = cloneCellMutationRef({
      sheetId,
      mutation: {
        kind: 'setCellValue',
        row: 0,
        col: 0,
        value: 7,
      },
    })
    const formulaRef = cloneCellMutationRef({
      sheetId,
      mutation: {
        kind: 'setCellFormula',
        row: 1,
        col: 1,
        formula: 'A1*2',
      },
    })
    const clearRef = cloneCellMutationRef({
      sheetId,
      mutation: {
        kind: 'clearCell',
        row: 2,
        col: 2,
      },
    })

    expect(valueRef).toEqual({
      sheetId,
      mutation: {
        kind: 'setCellValue',
        row: 0,
        col: 0,
        value: 7,
      },
    })
    expect(cellMutationRefToEngineOp(workbook, valueRef)).toEqual({
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'A1',
      value: 7,
    })
    expect(cellMutationRefToEngineOp(workbook, formulaRef)).toEqual({
      kind: 'setCellFormula',
      sheetName: 'Sheet1',
      address: 'B2',
      formula: 'A1*2',
    })
    expect(cellMutationRefToEngineOp(workbook, clearRef)).toEqual({
      kind: 'clearCell',
      sheetName: 'Sheet1',
      address: 'C3',
    })
  })

  it('counts only non-clear mutations and rejects unknown sheet ids', () => {
    const workbook = new WorkbookStore('cell-mutations')
    workbook.createSheet('Sheet1', 0)
    const sheetId = workbook.getSheet('Sheet1')!.id

    expect(
      countPotentialNewCellsForMutationRefs([
        { sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 1 } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1*2' } },
        { sheetId, mutation: { kind: 'clearCell', row: 0, col: 2 } },
      ]),
    ).toBe(2)

    expect(() =>
      cellMutationRefToEngineOp(workbook, {
        sheetId: 999,
        mutation: { kind: 'clearCell', row: 0, col: 0 },
      }),
    ).toThrow('Unknown sheet id: 999')
  })
})
