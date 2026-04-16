import { Effect } from 'effect'
import { ValueTag, type CellRangeRef, type CellSnapshot } from '@bilig/protocol'
import { formatAddress, parseCellAddress, translateFormulaReferences } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import { CellFlags } from '../../cell-store.js'
import { normalizeRange } from '../../engine-range-utils.js'
import type { EngineRuntimeState } from '../runtime-state.js'
import { EngineCellStateError } from '../errors.js'

export interface EngineCellStateService {
  readonly captureStoredCellOps: (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => Effect.Effect<EngineOp[], EngineCellStateError>
  readonly restoreCellOps: (sheetName: string, address: string) => Effect.Effect<EngineOp[], EngineCellStateError>
  readonly readRangeCells: (range: CellRangeRef) => Effect.Effect<CellSnapshot[][], EngineCellStateError>
  readonly toCellStateOps: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => Effect.Effect<EngineOp[], EngineCellStateError>
  readonly captureStoredCellOpsNow: (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[]
  readonly restoreCellOpsNow: (sheetName: string, address: string) => EngineOp[]
  readonly readRangeCellsNow: (range: CellRangeRef) => CellSnapshot[][]
  readonly toCellStateOpsNow: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[]
}

function cellStateErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function translateFormulaForTarget(
  formula: string,
  sourceSheetName: string,
  sourceAddress: string,
  targetSheetName: string,
  targetAddress: string,
): string {
  const source = parseCellAddress(sourceAddress, sourceSheetName)
  const target = parseCellAddress(targetAddress, targetSheetName)
  return translateFormulaReferences(formula, target.row - source.row, target.col - source.col)
}

export function createEngineCellStateService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook'>
  readonly getCell: (sheetName: string, address: string) => CellSnapshot
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
}): EngineCellStateService {
  const toCellStateOpsNow = (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
    formatOverride: string | null = snapshot.format ?? null,
  ): EngineOp[] => {
    const ops: EngineOp[] = []
    const shouldEmitFormat = formatOverride !== null && formatOverride !== undefined && formatOverride !== ''
    if (snapshot.formula !== undefined) {
      const translatedFormula =
        sourceSheetName && sourceAddress
          ? translateFormulaForTarget(snapshot.formula, sourceSheetName, sourceAddress, sheetName, address)
          : snapshot.formula
      ops.push({ kind: 'setCellFormula', sheetName, address, formula: translatedFormula })
    } else {
      switch (snapshot.value.tag) {
        case ValueTag.Empty:
          ops.push(
            (snapshot.flags & CellFlags.AuthoredBlank) !== 0
              ? { kind: 'setCellValue', sheetName, address, value: null, authoredBlank: true }
              : { kind: 'clearCell', sheetName, address },
          )
          break
        case ValueTag.Number:
        case ValueTag.Boolean:
        case ValueTag.String:
          ops.push({ kind: 'setCellValue', sheetName, address, value: snapshot.value.value })
          break
        case ValueTag.Error:
          ops.push({ kind: 'clearCell', sheetName, address })
          break
      }
    }
    if (shouldEmitFormat) {
      ops.push({
        kind: 'setCellFormat',
        sheetName,
        address,
        format: formatOverride,
      })
    }
    return ops
  }

  const captureStoredCellOpsNow = (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): EngineOp[] =>
    toCellStateOpsNow(
      sheetName,
      address,
      args.getCellByIndex(cellIndex),
      sourceSheetName,
      sourceAddress,
      args.state.workbook.getCellFormat(cellIndex) ?? null,
    )

  const restoreCellOpsNow = (sheetName: string, address: string): EngineOp[] => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return [{ kind: 'clearCell', sheetName, address }]
    }
    const snapshot = args.getCellByIndex(cellIndex)
    if (snapshot.formula !== undefined) {
      return [{ kind: 'setCellFormula', sheetName, address, formula: snapshot.formula }]
    }
    switch (snapshot.value.tag) {
      case ValueTag.Empty:
      case ValueTag.Error:
        return (snapshot.flags & CellFlags.AuthoredBlank) !== 0
          ? [{ kind: 'setCellValue', sheetName, address, value: null, authoredBlank: true }]
          : [{ kind: 'clearCell', sheetName, address }]
      case ValueTag.Number:
      case ValueTag.Boolean:
      case ValueTag.String:
        return [{ kind: 'setCellValue', sheetName, address, value: snapshot.value.value }]
    }
  }

  const readRangeCellsNow = (range: CellRangeRef): CellSnapshot[][] => {
    const bounds = normalizeRange(range)
    const rows: CellSnapshot[][] = []
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      const cells: CellSnapshot[] = []
      for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
        cells.push(args.getCell(range.sheetName, formatAddress(row, col)))
      }
      rows.push(cells)
    }
    return rows
  }

  return {
    captureStoredCellOps(cellIndex, sheetName, address, sourceSheetName, sourceAddress) {
      return Effect.try({
        try: () => captureStoredCellOpsNow(cellIndex, sheetName, address, sourceSheetName, sourceAddress),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(`Failed to capture stored cell ops for ${sheetName}!${address}`, cause),
            cause,
          }),
      })
    },
    restoreCellOps(sheetName, address) {
      return Effect.try({
        try: () => restoreCellOpsNow(sheetName, address),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(`Failed to restore cell ops for ${sheetName}!${address}`, cause),
            cause,
          }),
      })
    },
    readRangeCells(range) {
      return Effect.try({
        try: () => readRangeCellsNow(range),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(`Failed to read range ${range.sheetName}!${range.startAddress}:${range.endAddress}`, cause),
            cause,
          }),
      })
    },
    toCellStateOps(sheetName, address, snapshot, sourceSheetName, sourceAddress) {
      return Effect.try({
        try: () => toCellStateOpsNow(sheetName, address, snapshot, sourceSheetName, sourceAddress),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(`Failed to materialize cell state ops for ${sheetName}!${address}`, cause),
            cause,
          }),
      })
    },
    captureStoredCellOpsNow,
    restoreCellOpsNow,
    readRangeCellsNow,
    toCellStateOpsNow,
  }
}
