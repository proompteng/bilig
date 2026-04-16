import { CellFlags, loadLiteralSheetIntoEmptySheet, makeCellKey, type SpreadsheetEngine, type EngineCellMutationRef } from '@bilig/core'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { WorkPaperCellAddress, WorkPaperSheet } from './work-paper-types.js'

export function tryLoadInitialLiteralSheet(engine: SpreadsheetEngine, sheetId: number, content: WorkPaperSheet): boolean {
  if (sheetContainsFormulaContent(content)) {
    return false
  }
  loadLiteralSheetIntoEmptySheet(engine.workbook, engine.strings, sheetId, content)
  return true
}

function sheetContainsFormulaContent(content: WorkPaperSheet): boolean {
  return content.some((row) => row.some((value) => typeof value === 'string' && value.trim().startsWith('=')))
}

export interface PreparedInitialMixedSheetLoad {
  formulaRefs: EngineCellMutationRef[]
  potentialNewCells: number
}

export function prepareInitialMixedSheetLoad(args: {
  engine: SpreadsheetEngine
  sheetId: number
  content: WorkPaperSheet
  rewriteFormula: (formula: string, destination: WorkPaperCellAddress) => string
}): PreparedInitialMixedSheetLoad {
  const sheet = args.engine.workbook.getSheetById(args.sheetId)
  if (!sheet) {
    throw new Error(`Unknown sheet id: ${args.sheetId}`)
  }

  let literalCount = 0
  let formulaCount = 0
  for (let rowIndex = 0; rowIndex < args.content.length; rowIndex += 1) {
    const row = args.content[rowIndex]!
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const raw = row[colIndex]!
      if (typeof raw === 'string' && raw.trim().startsWith('=')) {
        formulaCount += 1
        continue
      }
      if (raw !== null) {
        literalCount += 1
      }
    }
  }

  const cellStore = args.engine.workbook.cellStore
  if (literalCount > 0) {
    cellStore.ensureCapacity(cellStore.size + literalCount)
  }
  const formulaRefs = formulaCount === 0 ? [] : Array<EngineCellMutationRef>(formulaCount)
  let formulaIndex = 0
  const previousOnSetValue = cellStore.onSetValue
  cellStore.onSetValue = null
  try {
    for (let rowIndex = 0; rowIndex < args.content.length; rowIndex += 1) {
      const row = args.content[rowIndex]!
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        const raw = row[colIndex]!
        if (typeof raw === 'string') {
          const trimmed = raw.trim()
          if (trimmed.startsWith('=')) {
            formulaRefs[formulaIndex] = {
              sheetId: args.sheetId,
              mutation: {
                kind: 'setCellFormula',
                row: rowIndex,
                col: colIndex,
                formula: args.rewriteFormula(trimmed.slice(1), {
                  sheet: args.sheetId,
                  row: rowIndex,
                  col: colIndex,
                }),
              },
            }
            formulaIndex += 1
            continue
          }
        }
        if (raw === null) {
          continue
        }
        const cellIndex = cellStore.allocate(args.sheetId, rowIndex, colIndex)
        args.engine.workbook.cellKeyToIndex.set(makeCellKey(args.sheetId, rowIndex, colIndex), cellIndex)
        sheet.grid.set(rowIndex, colIndex, cellIndex)
        cellStore.flags[cellIndex] = CellFlags.Materialized
        cellStore.formulaIds[cellIndex] = 0
        cellStore.errors[cellIndex] = ErrorCode.None
        cellStore.versions[cellIndex] = 1
        cellStore.topoRanks[cellIndex] = 0
        cellStore.cycleGroupIds[cellIndex] = -1
        if (typeof raw === 'number') {
          cellStore.tags[cellIndex] = ValueTag.Number
          cellStore.numbers[cellIndex] = raw
          cellStore.stringIds[cellIndex] = 0
        } else if (typeof raw === 'boolean') {
          cellStore.tags[cellIndex] = ValueTag.Boolean
          cellStore.numbers[cellIndex] = raw ? 1 : 0
          cellStore.stringIds[cellIndex] = 0
        } else {
          cellStore.tags[cellIndex] = ValueTag.String
          cellStore.numbers[cellIndex] = 0
          cellStore.stringIds[cellIndex] = args.engine.strings.intern(raw)
        }
        args.engine.workbook.notifyCellValueWritten(cellIndex)
      }
    }
  } finally {
    cellStore.onSetValue = previousOnSetValue
  }

  return {
    formulaRefs,
    potentialNewCells: formulaRefs.length,
  }
}

export function loadInitialMixedSheet(args: {
  engine: SpreadsheetEngine
  sheetId: number
  content: WorkPaperSheet
  rewriteFormula: (formula: string, destination: WorkPaperCellAddress) => string
}): void {
  const prepared = prepareInitialMixedSheetLoad(args)
  if (prepared.formulaRefs.length === 0) {
    return
  }
  args.engine.initializeCellFormulasAt(prepared.formulaRefs, prepared.potentialNewCells)
}
