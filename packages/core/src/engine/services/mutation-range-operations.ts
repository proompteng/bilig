import { Effect } from 'effect'
import { formatAddress } from '@bilig/formula'
import type { CellRangeRef, CellSnapshot, LiteralInput } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'
import { parseCsv, parseCsvCellInput, resolveCsvParseOptions, type CsvParseOptions } from '../../csv.js'
import { normalizeRange } from '../../engine-range-utils.js'
import type { WorkbookStore } from '../../workbook-store.js'
import { EngineMutationError } from '../errors.js'
import { getMutationMatrixCell } from './mutation-cell-content-helpers.js'
import {
  hasMutationCellContent,
  hasStoredMutationCellState,
  readDesiredMutationCellState,
  readStoredMutationCellState,
  shouldApplyMutationCellState,
  type ComparableMutationCellState,
} from './mutation-cell-state-helpers.js'

export interface MutationRangeOperations {
  readonly setRangeValues: (range: CellRangeRef, values: readonly (readonly LiteralInput[])[]) => Effect.Effect<void, EngineMutationError>
  readonly setRangeFormulas: (range: CellRangeRef, formulas: readonly (readonly string[])[]) => Effect.Effect<void, EngineMutationError>
  readonly clearRange: (range: CellRangeRef) => Effect.Effect<void, EngineMutationError>
  readonly fillRange: (source: CellRangeRef, target: CellRangeRef) => Effect.Effect<void, EngineMutationError>
  readonly copyRange: (source: CellRangeRef, target: CellRangeRef) => Effect.Effect<void, EngineMutationError>
  readonly moveRange: (source: CellRangeRef, target: CellRangeRef) => Effect.Effect<void, EngineMutationError>
  readonly importSheetCsv: (sheetName: string, csv: string, options?: CsvParseOptions) => Effect.Effect<void, EngineMutationError>
}

interface MutationRangeOperationsRuntime {
  readonly workbook: WorkbookStore
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly readRangeCells: (range: CellRangeRef) => CellSnapshot[][]
  readonly toCellStateOps: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[]
  readonly executeLocal: (
    ops: EngineOp[],
    potentialNewCells?: number,
    options?: { readonly returnUndoOps?: boolean },
  ) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>
}

function mutationRangeErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

export function createMutationRangeOperations(args: MutationRangeOperationsRuntime): MutationRangeOperations {
  const readStoredCellState = (sheetName: string, address: string): ComparableMutationCellState => {
    const cellIndex = args.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return { value: null, format: null, authoredBlank: false }
    }
    const snapshot = args.getCellByIndex(cellIndex)
    const format = args.workbook.getCellFormat(cellIndex) ?? null
    return readStoredMutationCellState(snapshot, format, args.workbook.cellStore.flags[cellIndex])
  }

  const readDesiredCellState = (
    targetSheetName: string,
    targetAddress: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
    formatOverride: string | null = snapshot.format ?? null,
  ): ComparableMutationCellState =>
    readDesiredMutationCellState({
      targetSheetName,
      targetAddress,
      snapshot,
      ...(sourceSheetName !== undefined ? { sourceSheetName } : {}),
      ...(sourceAddress !== undefined ? { sourceAddress } : {}),
      formatOverride,
    })

  const hasStoredCellContent = (sheetName: string, address: string): boolean => {
    const cellIndex = args.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return false
    }
    return hasMutationCellContent(args.getCellByIndex(cellIndex))
  }

  const hasStoredCellState = (sheetName: string, address: string): boolean => {
    const cellIndex = args.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return false
    }
    return hasStoredMutationCellState(
      args.getCellByIndex(cellIndex),
      args.workbook.getCellFormat(cellIndex),
      args.workbook.cellStore.flags[cellIndex],
    )
  }

  const shouldApplyCellState = (
    targetSheetName: string,
    targetAddress: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): boolean => {
    const stored = readStoredCellState(targetSheetName, targetAddress)
    const desired = readDesiredCellState(targetSheetName, targetAddress, snapshot, sourceSheetName, sourceAddress)
    return shouldApplyMutationCellState(stored, desired)
  }

  return {
    setRangeValues(range, values) {
      return Effect.try({
        try: () => {
          const bounds = normalizeRange(range)
          const expectedHeight = bounds.endRow - bounds.startRow + 1
          const expectedWidth = bounds.endCol - bounds.startCol + 1
          if (values.length !== expectedHeight || values.some((row) => row.length !== expectedWidth)) {
            throw new Error('setRangeValues requires a value matrix that exactly matches the target range')
          }

          const ops: EngineOp[] = []
          for (let rowOffset = 0; rowOffset < expectedHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < expectedWidth; colOffset += 1) {
              const address = formatAddress(bounds.startRow + rowOffset, bounds.startCol + colOffset)
              const current = readStoredCellState(range.sheetName, address)
              const nextValue = values[rowOffset]![colOffset] ?? null
              if (current.formula === undefined && current.value === nextValue) {
                continue
              }
              ops.push({
                kind: 'setCellValue',
                sheetName: range.sheetName,
                address,
                value: nextValue,
              })
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(args.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationRangeErrorMessage('Failed to set range values', cause),
            cause,
          }),
      })
    },
    setRangeFormulas(range, formulas) {
      return Effect.try({
        try: () => {
          const bounds = normalizeRange(range)
          const expectedHeight = bounds.endRow - bounds.startRow + 1
          const expectedWidth = bounds.endCol - bounds.startCol + 1
          if (formulas.length !== expectedHeight || formulas.some((row) => row.length !== expectedWidth)) {
            throw new Error('setRangeFormulas requires a formula matrix that exactly matches the target range')
          }

          const opCount = expectedHeight * expectedWidth
          const ops = Array.from<EngineOp>({ length: opCount })
          let opIndex = 0
          for (let rowOffset = 0; rowOffset < expectedHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < expectedWidth; colOffset += 1) {
              ops[opIndex] = {
                kind: 'setCellFormula',
                sheetName: range.sheetName,
                address: formatAddress(bounds.startRow + rowOffset, bounds.startCol + colOffset),
                formula: formulas[rowOffset]![colOffset] ?? '',
              }
              opIndex += 1
            }
          }
          Effect.runSync(args.executeLocal(ops, opCount))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationRangeErrorMessage('Failed to set range formulas', cause),
            cause,
          }),
      })
    },
    clearRange(range) {
      return Effect.try({
        try: () => {
          const bounds = normalizeRange(range)
          const ops: EngineOp[] = []
          for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
            for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
              const address = formatAddress(row, col)
              if (!hasStoredCellContent(range.sheetName, address)) {
                continue
              }
              ops.push({
                kind: 'clearCell',
                sheetName: range.sheetName,
                address,
              })
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(args.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationRangeErrorMessage('Failed to clear range', cause),
            cause,
          }),
      })
    },
    fillRange(source, target) {
      return Effect.try({
        try: () => {
          const sourceMatrix = args.readRangeCells(source)
          const targetBounds = normalizeRange(target)
          const sourceBounds = normalizeRange(source)
          const sourceHeight = sourceMatrix.length
          const sourceWidth = sourceMatrix[0]?.length ?? 0
          if (sourceHeight === 0 || sourceWidth === 0) {
            return
          }

          const ops: EngineOp[] = []
          for (let row = targetBounds.startRow; row <= targetBounds.endRow; row += 1) {
            for (let col = targetBounds.startCol; col <= targetBounds.endCol; col += 1) {
              const sourceRowOffset = (row - targetBounds.startRow) % sourceHeight
              const sourceColOffset = (col - targetBounds.startCol) % sourceWidth
              const sourceCell = getMutationMatrixCell(sourceMatrix, sourceRowOffset, sourceColOffset)
              const sourceAddress = formatAddress(sourceBounds.startRow + sourceRowOffset, sourceBounds.startCol + sourceColOffset)
              const nextAddress = formatAddress(row, col)
              if (!shouldApplyCellState(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress)) {
                continue
              }
              ops.push(...args.toCellStateOps(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress))
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(args.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationRangeErrorMessage('Failed to fill range', cause),
            cause,
          }),
      })
    },
    copyRange(source, target) {
      return Effect.try({
        try: () => {
          const sourceMatrix = args.readRangeCells(source)
          const targetBounds = normalizeRange(target)
          const sourceBounds = normalizeRange(source)
          const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1
          const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1
          const targetHeight = targetBounds.endRow - targetBounds.startRow + 1
          const targetWidth = targetBounds.endCol - targetBounds.startCol + 1
          if (sourceHeight !== targetHeight || sourceWidth !== targetWidth) {
            throw new Error('copyRange requires source and target dimensions to match exactly')
          }

          const ops: EngineOp[] = []
          for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
              const nextAddress = formatAddress(targetBounds.startRow + rowOffset, targetBounds.startCol + colOffset)
              const sourceAddress = formatAddress(sourceBounds.startRow + rowOffset, sourceBounds.startCol + colOffset)
              const sourceCell = getMutationMatrixCell(sourceMatrix, rowOffset, colOffset)
              if (!shouldApplyCellState(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress)) {
                continue
              }
              ops.push(...args.toCellStateOps(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress))
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(args.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationRangeErrorMessage('Failed to copy range', cause),
            cause,
          }),
      })
    },
    moveRange(source, target) {
      return Effect.try({
        try: () => {
          const sourceMatrix = args.readRangeCells(source)
          const targetBounds = normalizeRange(target)
          const sourceBounds = normalizeRange(source)
          const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1
          const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1
          const targetHeight = targetBounds.endRow - targetBounds.startRow + 1
          const targetWidth = targetBounds.endCol - targetBounds.startCol + 1
          if (sourceHeight !== targetHeight || sourceWidth !== targetWidth) {
            throw new Error('moveRange requires source and target dimensions to match exactly')
          }
          if (
            source.sheetName === target.sheetName &&
            sourceBounds.startRow === targetBounds.startRow &&
            sourceBounds.endRow === targetBounds.endRow &&
            sourceBounds.startCol === targetBounds.startCol &&
            sourceBounds.endCol === targetBounds.endCol
          ) {
            return
          }

          const ops: EngineOp[] = []
          for (let row = sourceBounds.startRow; row <= sourceBounds.endRow; row += 1) {
            for (let col = sourceBounds.startCol; col <= sourceBounds.endCol; col += 1) {
              const address = formatAddress(row, col)
              if (!hasStoredCellState(source.sheetName, address)) {
                continue
              }
              ops.push({
                kind: 'clearCell',
                sheetName: source.sheetName,
                address,
              })
            }
          }
          for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
              const nextAddress = formatAddress(targetBounds.startRow + rowOffset, targetBounds.startCol + colOffset)
              const sourceAddress = formatAddress(sourceBounds.startRow + rowOffset, sourceBounds.startCol + colOffset)
              const sourceCell = getMutationMatrixCell(sourceMatrix, rowOffset, colOffset)
              if (!shouldApplyCellState(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress)) {
                continue
              }
              ops.push(...args.toCellStateOps(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress))
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(args.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationRangeErrorMessage('Failed to move range', cause),
            cause,
          }),
      })
    },
    importSheetCsv(sheetName, csv, options) {
      return Effect.try({
        try: () => {
          const csvOptions = resolveCsvParseOptions(csv, options)
          const rows = parseCsv(csv, csvOptions)
          const existingSheet = args.workbook.getSheet(sheetName)
          const order = existingSheet?.order ?? args.workbook.sheetsByName.size
          const ops: EngineOp[] = []
          let potentialNewCells = 0

          if (existingSheet) {
            ops.push({ kind: 'deleteSheet', name: sheetName })
          }
          ops.push({ kind: 'upsertSheet', name: sheetName, order })

          rows.forEach((row, rowIndex) => {
            row.forEach((raw, colIndex) => {
              const parsed = parseCsvCellInput(raw, csvOptions)
              if (!parsed) {
                return
              }
              const address = formatAddress(rowIndex, colIndex)
              if (parsed.formula !== undefined) {
                ops.push({ kind: 'setCellFormula', sheetName, address, formula: parsed.formula })
                potentialNewCells += 1
                return
              }
              ops.push({ kind: 'setCellValue', sheetName, address, value: parsed.value ?? null })
              potentialNewCells += 1
            })
          })

          Effect.runSync(args.executeLocal(ops, potentialNewCells))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationRangeErrorMessage('Failed to import sheet CSV', cause),
            cause,
          }),
      })
    },
  }
}
