import { parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, CellSnapshot } from '@bilig/protocol'
import { normalizeCellRange } from './workbook-optimistic-range.js'

export function optimisticCellTargetFromKey(key: string): { sheetName: string; address: string } | null {
  const separatorIndex = key.lastIndexOf(':')
  if (separatorIndex <= 0 || separatorIndex === key.length - 1) return null
  return { sheetName: key.slice(0, separatorIndex), address: key.slice(separatorIndex + 1) }
}

export function supersedeOptimisticSeedsInRange(
  range: CellRangeRef,
  seeds: Map<string, string>,
  resolvedValues: Map<string, string>,
  snapshots: Map<string, CellSnapshot>,
): (() => void) | null {
  const bounds = normalizeCellRange(range)
  const removed: Array<readonly [string, string, string | undefined, CellSnapshot | undefined]> = []
  for (const [key, seed] of seeds) {
    const target = optimisticCellTargetFromKey(key)
    if (target?.sheetName !== range.sheetName) continue
    const { row, col } = parseCellAddress(target.address, target.sheetName)
    if (row < bounds.startRow || row > bounds.endRow || col < bounds.startCol || col > bounds.endCol) continue
    removed.push([key, seed, resolvedValues.get(key), snapshots.get(key)])
    seeds.delete(key)
    resolvedValues.delete(key)
    snapshots.delete(key)
  }
  if (removed.length === 0) return null
  return () => {
    for (const [key, seed, resolvedValue, snapshot] of removed) {
      if (seeds.has(key)) continue
      seeds.set(key, seed)
      if (resolvedValue !== undefined) resolvedValues.set(key, resolvedValue)
      if (snapshot !== undefined) snapshots.set(key, snapshot)
    }
  }
}
