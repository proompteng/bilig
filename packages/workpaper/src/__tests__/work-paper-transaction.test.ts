import { afterEach, describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'

import { WorkPaper, WorkPaperOperationError, type WorkPaperCellAddress, type WorkPaperFunctionPluginDefinition } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function scalarMathPlugin(id: string, functionName: string, operation: (value: number) => number): WorkPaperFunctionPluginDefinition {
  return {
    id,
    implementedFunctions: {
      [functionName]: { method: functionName },
    },
    functions: {
      [functionName]: (value) => {
        if (value?.tag !== ValueTag.Number) {
          return { tag: ValueTag.Error, code: 3 }
        }
        return { tag: ValueTag.Number, value: operation(value.value) }
      },
    },
  }
}

afterEach(() => {
  WorkPaper.unregisterAllFunctions()
})

describe('WorkPaper transactions', () => {
  it('rolls back partial automation mutations when user code throws', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [[1, '=A1+1']],
    })
    const sheetId = workbook.getSheetId('Sheet1')!
    workbook.setCellContents(cell(sheetId, 0, 0), 5)
    workbook.copy({
      start: cell(sheetId, 0, 0),
      end: cell(sheetId, 0, 1),
    })
    const beforeSheets = workbook.getAllSheetsSerialized()
    const beforeNames = workbook.getAllNamedExpressionsSerialized()
    const sheetAdded = vi.fn()
    workbook.onDetailed('sheetAdded', sheetAdded)

    expect(() => {
      workbook.transaction(() => {
        workbook.setCellContents(cell(sheetId, 0, 0), 99)
        workbook.addNamedExpression('TemporaryRate', '=Sheet1!$A$1')
        workbook.addSheet('Scratch')
        workbook.copy({
          start: cell(sheetId, 0, 0),
          end: cell(sheetId, 0, 0),
        })
        throw new Error('automation failed')
      })
    }).toThrow('automation failed')

    expect(workbook.getAllSheetsSerialized()).toEqual(beforeSheets)
    expect(workbook.getAllNamedExpressionsSerialized()).toEqual(beforeNames)
    expect(workbook.getSheetNames()).toEqual(['Sheet1'])
    expect(workbook.isClipboardEmpty()).toBe(false)
    expect(sheetAdded).not.toHaveBeenCalled()

    const undoChanges = workbook.undo()
    expect(undoChanges).toHaveLength(2)
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
  })

  it('commits successful automation transactions as one undoable operation', () => {
    const workbook = WorkPaper.buildFromArray([[1, '=A1+1']])
    const sheetId = workbook.getSheetId('Sheet1')!

    const changes = workbook.transaction(() => {
      workbook.setCellContents(cell(sheetId, 0, 0), 10)
      workbook.setCellContents(cell(sheetId, 0, 1), '=A1*3')
    })

    expect(changes).toHaveLength(2)
    expect(
      workbook.getRangeSerialized({
        start: cell(sheetId, 0, 0),
        end: cell(sheetId, 0, 1),
      }),
    ).toEqual([[10, '=A1*3']])
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 30,
    })

    const undoChanges = workbook.undo()
    expect(undoChanges).toHaveLength(2)
    expect(
      workbook.getRangeSerialized({
        start: cell(sheetId, 0, 0),
        end: cell(sheetId, 0, 1),
      }),
    ).toEqual([[1, '=A1+1']])
  })

  it('rejects transactions inside an existing suppressed mutation scope', () => {
    const workbook = WorkPaper.buildFromArray([[1]])

    expect(() => {
      workbook.batch(() => {
        workbook.transaction(() => {})
      })
    }).toThrow(WorkPaperOperationError)
  })

  it('rolls back function plugin registry changes when config updates fail mid-transaction', () => {
    const basePlugin = scalarMathPlugin('transaction-base-plugin', 'BASEONLY', (value) => value + 1)
    const temporaryPlugin = scalarMathPlugin('transaction-temporary-plugin', 'TEMPONLY', (value) => value * 2)
    WorkPaper.registerFunctionPlugin(basePlugin)
    WorkPaper.registerFunctionPlugin(temporaryPlugin)

    const workbook = WorkPaper.buildFromArray([[2]], { functionPlugins: [basePlugin] })
    const sheetId = workbook.getSheetId('Sheet1')!

    workbook.setCellContents(cell(sheetId, 0, 1), '=BASEONLY(A1)')
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(workbook.getFunctionPlugin('BASEONLY')?.id).toBe(basePlugin.id)
    expect(workbook.getFunctionPlugin('TEMPONLY')).toBeUndefined()

    expect(() => {
      workbook.transaction(() => {
        workbook.updateConfig({ functionPlugins: [basePlugin, temporaryPlugin] })
        workbook.setCellContents(cell(sheetId, 0, 1), '=TEMPONLY(A1)')
        expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
          tag: ValueTag.Number,
          value: 4,
        })
        throw new Error('automation failed after plugin update')
      })
    }).toThrow('automation failed after plugin update')

    expect(workbook.getConfig().functionPlugins?.map((plugin) => plugin.id)).toEqual([basePlugin.id])
    expect(workbook.getFunctionPlugin('BASEONLY')?.id).toBe(basePlugin.id)
    expect(workbook.getFunctionPlugin('TEMPONLY')).toBeUndefined()
    expect(workbook.getCellFormula(cell(sheetId, 0, 1))).toBe('=BASEONLY(A1)')
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
  })
})
