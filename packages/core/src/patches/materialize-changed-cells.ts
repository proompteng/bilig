import { formatAddress, indexToColumn } from '@bilig/formula'
import type { EngineChangedCell } from '@bilig/protocol'
import { addEngineCounter, type EngineCounters } from '../perf/engine-counters.js'
import type { StringPool } from '../string-pool.js'
import type { WorkbookStore } from '../workbook-store.js'
import type {
  EngineCellPatch,
  EngineColumnInvalidationPatch,
  EnginePatch,
  EngineRangeInvalidationPatch,
  EngineRowInvalidationPatch,
} from './patch-types.js'

export interface MaterializeChangedCellsArgs {
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly counters?: EngineCounters
}

export interface EnginePatchCaptureRequest {
  readonly changedCellIndices: readonly number[] | Uint32Array
  readonly invalidation?: 'cells' | 'full'
  readonly invalidatedRanges?: readonly EngineRangeInvalidationPatch['range'][]
  readonly invalidatedRows?: ReadonlyArray<{
    readonly sheetName: string
    readonly startIndex: number
    readonly endIndex: number
  }>
  readonly invalidatedColumns?: ReadonlyArray<{
    readonly sheetName: string
    readonly startIndex: number
    readonly endIndex: number
  }>
}

function readChangedCell(args: MaterializeChangedCellsArgs, cellIndex: number, fallbackSheetName?: string): EngineCellPatch | null {
  const sheetId = args.workbook.cellStore.sheetIds[cellIndex]
  const position = args.workbook.getCellPosition(cellIndex)
  if (sheetId === undefined || !position) {
    return null
  }
  return {
    kind: 'cell',
    cellIndex,
    address: { sheet: sheetId, row: position.row, col: position.col },
    sheetName: fallbackSheetName ?? args.workbook.getSheetNameById(sheetId),
    a1: formatAddress(position.row, position.col),
    newValue: args.workbook.cellStore.getValue(cellIndex, (id) => args.strings.get(id)),
  }
}

export function materializeChangedCellPatches(
  args: MaterializeChangedCellsArgs,
  changedCellIndices: readonly number[] | Uint32Array,
): readonly EngineCellPatch[] {
  if (changedCellIndices.length === 0) {
    return []
  }
  if (changedCellIndices.length <= 2) {
    const first = readChangedCell(args, changedCellIndices[0]!)
    if (!first) {
      return []
    }
    if (changedCellIndices.length === 1) {
      if (args.counters) {
        addEngineCounter(args.counters, 'changedCellPayloadsBuilt', 1)
      }
      return [first]
    }
    const secondCellIndex = changedCellIndices[1]!
    const secondSheetId = args.workbook.cellStore.sheetIds[secondCellIndex]
    const second =
      secondSheetId !== undefined && secondSheetId === first.address.sheet
        ? readChangedCell(args, secondCellIndex, first.sheetName)
        : readChangedCell(args, secondCellIndex)
    const patches = second ? [first, second] : [first]
    if (args.counters) {
      addEngineCounter(args.counters, 'changedCellPayloadsBuilt', patches.length)
    }
    return patches
  }
  const sheetNames = new Map<number, string>()
  const columnLabels: string[] = []
  const formatAddressCached = (row: number, col: number): string => {
    let label = columnLabels[col]
    if (label === undefined) {
      label = indexToColumn(col)
      columnLabels[col] = label
    }
    return `${label}${row + 1}`
  }
  const patches: EngineCellPatch[] = []
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    const sheetId = args.workbook.cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      continue
    }
    const position = args.workbook.getCellPosition(cellIndex)
    if (!position) {
      continue
    }
    let sheetName = sheetNames.get(sheetId)
    if (sheetName === undefined) {
      sheetName = args.workbook.getSheetNameById(sheetId)
      sheetNames.set(sheetId, sheetName)
    }
    patches.push({
      kind: 'cell',
      cellIndex,
      address: { sheet: sheetId, row: position.row, col: position.col },
      sheetName,
      a1: formatAddressCached(position.row, position.col),
      newValue: args.workbook.cellStore.getValue(cellIndex, (id) => args.strings.get(id)),
    })
  }
  if (args.counters) {
    addEngineCounter(args.counters, 'changedCellPayloadsBuilt', patches.length)
  }
  return patches
}

function materializeRangeInvalidationPatches(
  invalidatedRanges: readonly EngineRangeInvalidationPatch['range'][],
): readonly EngineRangeInvalidationPatch[] {
  return invalidatedRanges.map((range) => ({
    kind: 'range-invalidation',
    range,
  }))
}

function materializeRowInvalidationPatches(
  invalidatedRows: EnginePatchCaptureRequest['invalidatedRows'],
): readonly EngineRowInvalidationPatch[] {
  return (invalidatedRows ?? []).map((row) => ({
    kind: 'row-invalidation',
    sheetName: row.sheetName,
    startIndex: row.startIndex,
    endIndex: row.endIndex,
  }))
}

function materializeColumnInvalidationPatches(
  invalidatedColumns: EnginePatchCaptureRequest['invalidatedColumns'],
): readonly EngineColumnInvalidationPatch[] {
  return (invalidatedColumns ?? []).map((column) => ({
    kind: 'column-invalidation',
    sheetName: column.sheetName,
    startIndex: column.startIndex,
    endIndex: column.endIndex,
  }))
}

export function materializeEnginePatches(args: MaterializeChangedCellsArgs, request: EnginePatchCaptureRequest): readonly EnginePatch[] {
  const cellPatches = materializeChangedCellPatches(args, request.changedCellIndices)
  const rangePatches = materializeRangeInvalidationPatches(request.invalidatedRanges ?? [])
  const rowPatches = materializeRowInvalidationPatches(request.invalidatedRows)
  const columnPatches = materializeColumnInvalidationPatches(request.invalidatedColumns)
  if (rangePatches.length === 0 && rowPatches.length === 0 && columnPatches.length === 0) {
    return cellPatches
  }
  return [...cellPatches, ...rangePatches, ...rowPatches, ...columnPatches]
}

export function materializeChangedCells(
  args: MaterializeChangedCellsArgs,
  changedCellIndices: readonly number[] | Uint32Array,
): readonly EngineChangedCell[] {
  return materializeChangedCellPatches(args, changedCellIndices) as readonly EngineChangedCell[]
}
