import type { SpreadsheetEngine } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, isDateLikeHeaderValue, isLikelyExcelDateSerialValue, type CellRangeRef } from '@bilig/protocol'
import {
  applyWorkbookAgentAnnotationCommand,
  isWorkbookAgentAnnotationCommand,
  applyWorkbookAgentConditionalFormatCommand,
  isWorkbookAgentConditionalFormatCommand,
  applyWorkbookAgentMediaCommand,
  isWorkbookAgentMediaCommand,
  applyWorkbookAgentObjectCommand,
  isWorkbookAgentObjectCommand,
  applyWorkbookAgentProtectionCommand,
  isWorkbookAgentProtectionCommand,
  applyWorkbookAgentStructuralCommand,
  isWorkbookAgentStructuralCommand,
  applyWorkbookAgentValidationCommand,
  isWorkbookAgentValidationCommand,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
} from '@bilig/agent-api'
import type { EngineOp } from '@bilig/workbook-domain'
import type { WorkbookChangeUndoBundle } from '@bilig/zero-sync'

function normalizeFormula(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula
}

function toEngineUndoBundle(undoOps: readonly EngineOp[] | null): WorkbookChangeUndoBundle | null {
  if (!undoOps || undoOps.length === 0) {
    return null
  }
  return {
    kind: 'engineOps',
    ops: structuredClone([...undoOps]),
  }
}

function buildWriteRangeTransaction(command: Extract<WorkbookAgentCommand, { kind: 'writeRange' }>): {
  readonly ops: EngineOp[]
  readonly potentialNewCells: number
} {
  const start = parseCellAddress(command.startAddress, command.sheetName)
  const ops: EngineOp[] = []
  let potentialNewCells = 0
  command.values.forEach((rowValues, rowOffset) => {
    rowValues.forEach((cellInput, colOffset) => {
      const address = formatAddress(start.row + rowOffset, start.col + colOffset)
      if (cellInput === null) {
        ops.push({ kind: 'clearCell', sheetName: command.sheetName, address })
        return
      }
      if (typeof cellInput === 'string' || typeof cellInput === 'number' || typeof cellInput === 'boolean') {
        ops.push({ kind: 'setCellValue', sheetName: command.sheetName, address, value: cellInput })
        potentialNewCells += 1
        return
      }
      if ('formula' in cellInput) {
        ops.push({
          kind: 'setCellFormula',
          sheetName: command.sheetName,
          address,
          formula: normalizeFormula(cellInput.formula),
        })
        potentialNewCells += 1
        return
      }
      ops.push({
        kind: 'setCellValue',
        sheetName: command.sheetName,
        address,
        value: cellInput.value,
      })
      potentialNewCells += 1
    })
  })
  return { ops, potentialNewCells }
}

function inferDateFormatRanges(engine: SpreadsheetEngine, range: CellRangeRef): readonly CellRangeRef[] {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  const inferredRanges: CellRangeRef[] = []

  for (let col = startCol; col <= endCol; col += 1) {
    let dataStartRow: number | null = null
    const rangeStartHeader = engine.getCell(range.sheetName, formatAddress(startRow, col)).value
    if (isDateLikeHeaderValue(rangeStartHeader) && startRow < endRow) {
      dataStartRow = startRow + 1
    } else if (startRow > 0) {
      const rowAboveHeader = engine.getCell(range.sheetName, formatAddress(startRow - 1, col)).value
      if (isDateLikeHeaderValue(rowAboveHeader)) {
        dataStartRow = startRow
      }
    }
    if (dataStartRow === null || dataStartRow > endRow) {
      continue
    }

    let sawDateSerial = false
    let allPopulatedValuesLookLikeDates = true
    for (let row = dataStartRow; row <= endRow; row += 1) {
      const cell = engine.getCell(range.sheetName, formatAddress(row, col))
      if (cell.value.tag === ValueTag.Empty) {
        continue
      }
      if (!isLikelyExcelDateSerialValue(cell.value)) {
        allPopulatedValuesLookLikeDates = false
        break
      }
      sawDateSerial = true
    }

    if (!sawDateSerial || !allPopulatedValuesLookLikeDates) {
      continue
    }

    inferredRanges.push({
      sheetName: range.sheetName,
      startAddress: formatAddress(dataStartRow, col),
      endAddress: formatAddress(endRow, col),
    })
  }

  return inferredRanges
}

function applyInferredDateFormats(engine: SpreadsheetEngine, range: CellRangeRef): readonly EngineOp[] {
  const aggregatedUndoOps: EngineOp[] = []
  for (const targetRange of inferDateFormatRanges(engine, range)) {
    const undoOps = engine.captureUndoOps(() => {
      engine.setRangeNumberFormat(targetRange, { kind: 'date', dateStyle: 'short' })
    }).undoOps
    if (undoOps?.length) {
      aggregatedUndoOps.unshift(...undoOps)
    }
  }
  return aggregatedUndoOps
}

function applyWorkbookAgentCommandWithUndoCapture(engine: SpreadsheetEngine, command: WorkbookAgentCommand): readonly EngineOp[] | null {
  if (isWorkbookAgentStructuralCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentStructuralCommand(engine, command)
    }).undoOps
  }
  if (isWorkbookAgentObjectCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentObjectCommand(engine, command)
    }).undoOps
  }
  if (isWorkbookAgentMediaCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentMediaCommand(engine, command)
    }).undoOps
  }
  if (isWorkbookAgentProtectionCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentProtectionCommand(engine, command)
    }).undoOps
  }
  if (isWorkbookAgentConditionalFormatCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentConditionalFormatCommand(engine, command)
    }).undoOps
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentAnnotationCommand(engine, command)
    }).undoOps
  }
  if (isWorkbookAgentValidationCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentValidationCommand(engine, command)
    }).undoOps
  }
  switch (command.kind) {
    case 'writeRange': {
      const transaction = buildWriteRangeTransaction(command)
      const undoOps =
        engine.applyOps(transaction.ops, {
          captureUndo: true,
          potentialNewCells: transaction.potentialNewCells,
        }) ?? []
      const inferredDateFormatUndoOps = applyInferredDateFormats(engine, {
        sheetName: command.sheetName,
        startAddress: command.startAddress,
        endAddress: formatAddress(
          parseCellAddress(command.startAddress, command.sheetName).row + command.values.length - 1,
          parseCellAddress(command.startAddress, command.sheetName).col +
            Math.max(0, ...command.values.map((row) => (row.length === 0 ? 0 : row.length - 1))),
        ),
      })
      return [...inferredDateFormatUndoOps, ...undoOps]
    }
    case 'setRangeFormulas': {
      const undoOps =
        engine.captureUndoOps(() => {
          engine.setRangeFormulas(
            command.range,
            command.formulas.map((row) => row.map((formula) => normalizeFormula(formula))),
          )
        }).undoOps ?? []
      const inferredDateFormatUndoOps = applyInferredDateFormats(engine, command.range)
      return [...inferredDateFormatUndoOps, ...undoOps]
    }
    case 'formatRange': {
      const aggregatedUndoOps: EngineOp[] = []
      if (command.patch !== undefined) {
        const undoOps = engine.captureUndoOps(() => {
          engine.setRangeStyle(command.range, command.patch!)
        }).undoOps
        if (undoOps?.length) {
          aggregatedUndoOps.unshift(...undoOps)
        }
      }
      if (command.numberFormat !== undefined) {
        const undoOps = engine.captureUndoOps(() => {
          engine.setRangeNumberFormat(command.range, command.numberFormat!)
        }).undoOps
        if (undoOps?.length) {
          aggregatedUndoOps.unshift(...undoOps)
        }
      }
      return aggregatedUndoOps
    }
    case 'clearRange':
      return engine.captureUndoOps(() => {
        engine.clearRange(command.range)
      }).undoOps
    case 'fillRange':
      return engine.captureUndoOps(() => {
        engine.fillRange(command.source, command.target)
      }).undoOps
    case 'copyRange':
      return engine.captureUndoOps(() => {
        engine.copyRange(command.source, command.target)
      }).undoOps
    case 'moveRange':
      return engine.captureUndoOps(() => {
        engine.moveRange(command.source, command.target)
      }).undoOps
    default: {
      const exhaustive: never = command
      throw new Error(`Unhandled workbook agent command: ${JSON.stringify(exhaustive)}`)
    }
  }
}

export function applyWorkbookAgentCommandBundleWithUndoCapture(
  engine: SpreadsheetEngine,
  bundle: WorkbookAgentCommandBundle,
): WorkbookChangeUndoBundle | null {
  const aggregatedUndoOps: EngineOp[] = []
  for (const command of bundle.commands) {
    const undoOps = applyWorkbookAgentCommandWithUndoCapture(engine, command)
    if (!undoOps || undoOps.length === 0) {
      continue
    }
    aggregatedUndoOps.unshift(...undoOps)
  }
  return toEngineUndoBundle(aggregatedUndoOps)
}
