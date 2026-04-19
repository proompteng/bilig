import { formatAddress } from '@bilig/formula'
import type { EngineChangedCell } from '@bilig/protocol'
import { addEngineCounter, type EngineCounters } from '../perf/engine-counters.js'
import type { StringPool } from '../string-pool.js'
import type { WorkbookStore } from '../workbook-store.js'
import type { EngineCellPatch } from './patch-types.js'

export interface MaterializeChangedCellsArgs {
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly counters?: EngineCounters
}

function readChangedCell(args: MaterializeChangedCellsArgs, cellIndex: number, fallbackSheetName?: string): EngineCellPatch | null {
  const sheetId = args.workbook.cellStore.sheetIds[cellIndex]
  const row = args.workbook.cellStore.rows[cellIndex]
  const col = args.workbook.cellStore.cols[cellIndex]
  if (sheetId === undefined || row === undefined || col === undefined) {
    return null
  }
  return {
    kind: 'cell',
    cellIndex,
    address: { sheet: sheetId, row, col },
    sheetName: fallbackSheetName ?? args.workbook.getSheetNameById(sheetId),
    a1: formatAddress(row, col),
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
  const patches: EngineCellPatch[] = []
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    const sheetId = args.workbook.cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      continue
    }
    let sheetName = sheetNames.get(sheetId)
    if (sheetName === undefined) {
      sheetName = args.workbook.getSheetNameById(sheetId)
      sheetNames.set(sheetId, sheetName)
    }
    const patch = readChangedCell(args, cellIndex, sheetName)
    if (patch) {
      patches.push(patch)
    }
  }
  if (args.counters) {
    addEngineCounter(args.counters, 'changedCellPayloadsBuilt', patches.length)
  }
  return patches
}

export function materializeChangedCells(
  args: MaterializeChangedCellsArgs,
  changedCellIndices: readonly number[] | Uint32Array,
): readonly EngineChangedCell[] {
  return materializeChangedCellPatches(args, changedCellIndices) as readonly EngineChangedCell[]
}
