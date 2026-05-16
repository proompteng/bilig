import type { SpreadsheetEngine } from '@bilig/core'
import type { EngineEvent } from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { collectChangedCellsBySheet } from './worker-runtime-support.js'

export interface ProjectionOverlayScope {
  readonly fullScan: boolean
  readonly cellAddressesBySheet: Map<string, Set<string>>
  readonly rowAxisIndicesBySheet: Map<string, Set<number>>
  readonly columnAxisIndicesBySheet: Map<string, Set<number>>
}

function addAddressToScope(entries: Map<string, Set<string>>, sheetName: string, address: string): void {
  const addresses = entries.get(sheetName) ?? new Set<string>()
  addresses.add(address)
  entries.set(sheetName, addresses)
}

function addAxisIndexToScope(entries: Map<string, Set<number>>, sheetName: string, index: number): void {
  const indices = entries.get(sheetName) ?? new Set<number>()
  indices.add(index)
  entries.set(sheetName, indices)
}

function addRangeToScope(
  scope: ProjectionOverlayScope,
  input: {
    sheetName: string
    startAddress: string
    endAddress: string
  },
): void {
  const start = parseCellAddress(input.startAddress, input.sheetName)
  const end = parseCellAddress(input.endAddress, input.sheetName)
  const rowStart = Math.min(start.row, end.row)
  const rowEnd = Math.max(start.row, end.row)
  const colStart = Math.min(start.col, end.col)
  const colEnd = Math.max(start.col, end.col)
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      addAddressToScope(scope.cellAddressesBySheet, input.sheetName, formatAddress(row, col))
    }
  }
}

export function createEmptyProjectionOverlayScope(): ProjectionOverlayScope {
  return {
    fullScan: false,
    cellAddressesBySheet: new Map<string, Set<string>>(),
    rowAxisIndicesBySheet: new Map<string, Set<number>>(),
    columnAxisIndicesBySheet: new Map<string, Set<number>>(),
  }
}

function cloneProjectionOverlayScope(scope: ProjectionOverlayScope): ProjectionOverlayScope {
  return {
    fullScan: scope.fullScan,
    cellAddressesBySheet: new Map(
      [...scope.cellAddressesBySheet.entries()].map(([sheetName, addresses]) => [sheetName, new Set(addresses)]),
    ),
    rowAxisIndicesBySheet: new Map([...scope.rowAxisIndicesBySheet.entries()].map(([sheetName, indices]) => [sheetName, new Set(indices)])),
    columnAxisIndicesBySheet: new Map(
      [...scope.columnAxisIndicesBySheet.entries()].map(([sheetName, indices]) => [sheetName, new Set(indices)]),
    ),
  }
}

function hasProjectionOverlayScopeEntries(scope: ProjectionOverlayScope): boolean {
  return (
    scope.fullScan || scope.cellAddressesBySheet.size > 0 || scope.rowAxisIndicesBySheet.size > 0 || scope.columnAxisIndicesBySheet.size > 0
  )
}

function mergeScopeEntries<T>(target: Map<string, Set<T>>, source: Map<string, Set<T>>): void {
  source.forEach((sourceEntries, sheetName) => {
    const targetEntries = target.get(sheetName) ?? new Set<T>()
    sourceEntries.forEach((entry) => {
      targetEntries.add(entry)
    })
    target.set(sheetName, targetEntries)
  })
}

export function collectProjectionOverlayScopeFromEngineEvents(
  engine: SpreadsheetEngine,
  events: readonly EngineEvent[],
): ProjectionOverlayScope | null {
  if (events.length === 0) {
    return null
  }
  const scope = createEmptyProjectionOverlayScope()
  for (const event of events) {
    if (event.invalidation === 'full') {
      return {
        ...scope,
        fullScan: true,
      }
    }
    collectChangedCellsBySheet(engine, event.changedCellIndices).forEach((changedCells, sheetName) => {
      changedCells.addresses.forEach((address) => {
        addAddressToScope(scope.cellAddressesBySheet, sheetName, address)
      })
    })
    event.invalidatedRanges.forEach((range) => {
      addRangeToScope(scope, range)
    })
    event.invalidatedRows.forEach((entry) => {
      for (let index = entry.startIndex; index <= entry.endIndex; index += 1) {
        addAxisIndexToScope(scope.rowAxisIndicesBySheet, entry.sheetName, index)
      }
    })
    event.invalidatedColumns.forEach((entry) => {
      for (let index = entry.startIndex; index <= entry.endIndex; index += 1) {
        addAxisIndexToScope(scope.columnAxisIndicesBySheet, entry.sheetName, index)
      }
    })
  }
  return hasProjectionOverlayScopeEntries(scope) ? scope : null
}

export function mergeProjectionOverlayScopes(
  current: ProjectionOverlayScope | null,
  incoming: ProjectionOverlayScope | null,
): ProjectionOverlayScope | null {
  if (incoming === null) {
    return current ? cloneProjectionOverlayScope(current) : null
  }
  if (current === null) {
    return cloneProjectionOverlayScope(incoming)
  }
  if (current.fullScan || incoming.fullScan) {
    return {
      fullScan: true,
      cellAddressesBySheet: new Map<string, Set<string>>(),
      rowAxisIndicesBySheet: new Map<string, Set<number>>(),
      columnAxisIndicesBySheet: new Map<string, Set<number>>(),
    }
  }
  const merged = cloneProjectionOverlayScope(current)
  mergeScopeEntries(merged.cellAddressesBySheet, incoming.cellAddressesBySheet)
  mergeScopeEntries(merged.rowAxisIndicesBySheet, incoming.rowAxisIndicesBySheet)
  mergeScopeEntries(merged.columnAxisIndicesBySheet, incoming.columnAxisIndicesBySheet)
  return merged
}
