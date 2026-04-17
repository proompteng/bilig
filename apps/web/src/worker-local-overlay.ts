import type { SpreadsheetEngine } from '@bilig/core'
import type { WorkbookLocalProjectionOverlay } from '@bilig/storage-browser'
import {
  ValueTag,
  type CellSnapshot,
  type CellStyleRecord,
  type CellValue,
  type EngineEvent,
  type WorkbookAxisEntrySnapshot,
} from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { collectChangedCellsBySheet } from './worker-runtime-support.js'
import { collectMaterializedSheetAddresses } from './worker-local-materialization.js'

export interface ProjectionOverlayScope {
  readonly fullScan: boolean
  readonly cellAddressesBySheet: Map<string, Set<string>>
  readonly rowAxisIndicesBySheet: Map<string, Set<number>>
  readonly columnAxisIndicesBySheet: Map<string, Set<number>>
}

function valueEquals(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false
  }
  switch (left.tag) {
    case ValueTag.Empty:
      return true
    case ValueTag.Number:
      return right.tag === ValueTag.Number && left.value === right.value
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value && left.stringId === right.stringId
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code
  }
}

function snapshotEquals(left: CellSnapshot, right: CellSnapshot): boolean {
  return (
    valueEquals(left.value, right.value) &&
    left.flags === right.flags &&
    left.input === right.input &&
    left.formula === right.formula &&
    left.format === right.format &&
    left.styleId === right.styleId &&
    left.numberFormatId === right.numberFormatId
  )
}

function axisEntryEquals(left: WorkbookAxisEntrySnapshot | undefined, right: WorkbookAxisEntrySnapshot | undefined): boolean {
  return left?.size === right?.size && (left?.hidden ?? false) === (right?.hidden ?? false)
}

function styleEquals(left: CellStyleRecord | undefined, right: CellStyleRecord | undefined): boolean {
  return JSON.stringify(left ?? null) === JSON.stringify(right ?? null)
}

function listOrderedSheetNames(authoritativeEngine: SpreadsheetEngine, projectionEngine: SpreadsheetEngine): string[] {
  const sheetNames = new Set<string>()
  authoritativeEngine.workbook.sheetsByName.forEach((_sheet, sheetName) => {
    sheetNames.add(sheetName)
  })
  projectionEngine.workbook.sheetsByName.forEach((_sheet, sheetName) => {
    sheetNames.add(sheetName)
  })
  return [...sheetNames].toSorted((left, right) => {
    const leftSheet = authoritativeEngine.workbook.getSheet(left) ?? projectionEngine.workbook.getSheet(left)
    const rightSheet = authoritativeEngine.workbook.getSheet(right) ?? projectionEngine.workbook.getSheet(right)
    return (leftSheet?.order ?? 0) - (rightSheet?.order ?? 0)
  })
}

function listUnionMaterializedAddresses(
  authoritativeEngine: SpreadsheetEngine,
  projectionEngine: SpreadsheetEngine,
  sheetName: string,
): string[] {
  const addresses = new Set<string>(collectMaterializedSheetAddresses(authoritativeEngine, sheetName))
  collectMaterializedSheetAddresses(projectionEngine, sheetName).forEach((address) => {
    addresses.add(address)
  })
  return sortAddresses(addresses, sheetName)
}

function listAxisIndices(
  authoritativeEntries: readonly WorkbookAxisEntrySnapshot[],
  projectionEntries: readonly WorkbookAxisEntrySnapshot[],
): number[] {
  const indices = new Set<number>(authoritativeEntries.map((entry) => entry.index))
  projectionEntries.forEach((entry) => {
    indices.add(entry.index)
  })
  return [...indices].toSorted((left, right) => left - right)
}

function sortAddresses(addresses: Iterable<string>, sheetName: string): string[] {
  return [...addresses].toSorted((left, right) => {
    const leftParsed = parseCellAddress(left, sheetName)
    const rightParsed = parseCellAddress(right, sheetName)
    return leftParsed.row - rightParsed.row || leftParsed.col - rightParsed.col
  })
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

function listScopedSheetNames(scope: ProjectionOverlayScope): string[] {
  return [
    ...new Set([...scope.cellAddressesBySheet.keys(), ...scope.rowAxisIndicesBySheet.keys(), ...scope.columnAxisIndicesBySheet.keys()]),
  ]
}

function resolveCellAddresses(input: {
  authoritativeEngine: SpreadsheetEngine
  projectionEngine: SpreadsheetEngine
  sheetName: string
  scope: ProjectionOverlayScope | null
}): string[] {
  if (input.scope === null || input.scope.fullScan) {
    return listUnionMaterializedAddresses(input.authoritativeEngine, input.projectionEngine, input.sheetName)
  }
  return sortAddresses(input.scope.cellAddressesBySheet.get(input.sheetName) ?? [], input.sheetName)
}

function resolveAxisIndices(
  scopeEntries: Map<string, Set<number>>,
  authoritativeEntries: readonly WorkbookAxisEntrySnapshot[],
  projectionEntries: readonly WorkbookAxisEntrySnapshot[],
  sheetName: string,
  scope: ProjectionOverlayScope | null,
): number[] {
  if (scope === null || scope.fullScan) {
    return listAxisIndices(authoritativeEntries, projectionEntries)
  }
  return [...(scopeEntries.get(sheetName) ?? [])].toSorted((left, right) => left - right)
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

export function buildWorkbookLocalProjectionOverlay(input: {
  authoritativeEngine: SpreadsheetEngine
  projectionEngine: SpreadsheetEngine
  scope?: ProjectionOverlayScope | null
}): WorkbookLocalProjectionOverlay {
  const { authoritativeEngine, projectionEngine, scope = null } = input
  if (scope && !scope.fullScan && listScopedSheetNames(scope).length === 0) {
    return {
      cells: [],
      rowAxisEntries: [],
      columnAxisEntries: [],
      styles: [],
    }
  }

  const cells: Array<WorkbookLocalProjectionOverlay['cells'][number]> = []
  const rowAxisEntries: Array<WorkbookLocalProjectionOverlay['rowAxisEntries'][number]> = []
  const columnAxisEntries: Array<WorkbookLocalProjectionOverlay['columnAxisEntries'][number]> = []
  const overlayStyleIds = new Set<string>()
  const sheetNames = scope && !scope.fullScan ? listScopedSheetNames(scope) : listOrderedSheetNames(authoritativeEngine, projectionEngine)

  for (const sheetName of sheetNames) {
    const sheet = projectionEngine.workbook.getSheet(sheetName) ?? authoritativeEngine.workbook.getSheet(sheetName)
    if (!sheet) {
      continue
    }

    for (const address of resolveCellAddresses({
      authoritativeEngine,
      projectionEngine,
      sheetName,
      scope,
    })) {
      const authoritativeSnapshot = authoritativeEngine.getCell(sheetName, address)
      const projectionSnapshot = projectionEngine.getCell(sheetName, address)
      if (snapshotEquals(authoritativeSnapshot, projectionSnapshot)) {
        continue
      }
      const parsed = parseCellAddress(address, sheetName)
      cells.push({
        sheetId: sheet.id,
        sheetName,
        address,
        rowNum: parsed.row,
        colNum: parsed.col,
        value: projectionSnapshot.value,
        flags: projectionSnapshot.flags,
        version: projectionSnapshot.version,
        input: projectionSnapshot.input,
        formula: projectionSnapshot.formula,
        format: projectionSnapshot.format,
        styleId: projectionSnapshot.styleId,
        numberFormatId: projectionSnapshot.numberFormatId,
      })
      if (projectionSnapshot.styleId && projectionSnapshot.styleId !== 'style-0') {
        overlayStyleIds.add(projectionSnapshot.styleId)
      }
    }

    const authoritativeRowAxisEntries = authoritativeEngine.getRowAxisEntries(sheetName)
    const projectionRowAxisEntries = projectionEngine.getRowAxisEntries(sheetName)
    const authoritativeRowAxisByIndex = new Map(authoritativeRowAxisEntries.map((entry) => [entry.index, entry]))
    const projectionRowAxisByIndex = new Map(projectionRowAxisEntries.map((entry) => [entry.index, entry]))
    for (const index of resolveAxisIndices(
      scope?.rowAxisIndicesBySheet ?? new Map<string, Set<number>>(),
      authoritativeRowAxisEntries,
      projectionRowAxisEntries,
      sheetName,
      scope,
    )) {
      const authoritativeEntry = authoritativeRowAxisByIndex.get(index)
      const projectionEntry = projectionRowAxisByIndex.get(index)
      if (axisEntryEquals(authoritativeEntry, projectionEntry)) {
        continue
      }
      rowAxisEntries.push({
        sheetId: sheet.id,
        sheetName,
        entry: {
          id: projectionEntry?.id ?? authoritativeEntry?.id ?? `${sheetName}:row:${String(index)}`,
          index,
          ...(projectionEntry?.size !== undefined ? { size: projectionEntry.size } : {}),
          hidden: projectionEntry?.hidden ?? false,
        },
      })
    }

    const authoritativeColumnAxisEntries = authoritativeEngine.getColumnAxisEntries(sheetName)
    const projectionColumnAxisEntries = projectionEngine.getColumnAxisEntries(sheetName)
    const authoritativeColumnAxisByIndex = new Map(authoritativeColumnAxisEntries.map((entry) => [entry.index, entry]))
    const projectionColumnAxisByIndex = new Map(projectionColumnAxisEntries.map((entry) => [entry.index, entry]))
    for (const index of resolveAxisIndices(
      scope?.columnAxisIndicesBySheet ?? new Map<string, Set<number>>(),
      authoritativeColumnAxisEntries,
      projectionColumnAxisEntries,
      sheetName,
      scope,
    )) {
      const authoritativeEntry = authoritativeColumnAxisByIndex.get(index)
      const projectionEntry = projectionColumnAxisByIndex.get(index)
      if (axisEntryEquals(authoritativeEntry, projectionEntry)) {
        continue
      }
      columnAxisEntries.push({
        sheetId: sheet.id,
        sheetName,
        entry: {
          id: projectionEntry?.id ?? authoritativeEntry?.id ?? `${sheetName}:column:${String(index)}`,
          index,
          ...(projectionEntry?.size !== undefined ? { size: projectionEntry.size } : {}),
          hidden: projectionEntry?.hidden ?? false,
        },
      })
    }
  }

  const styles: CellStyleRecord[] = []
  overlayStyleIds.forEach((styleId) => {
    const projectionStyle = projectionEngine.getCellStyle(styleId)
    const authoritativeStyle = authoritativeEngine.getCellStyle(styleId)
    if (projectionStyle && !styleEquals(authoritativeStyle, projectionStyle)) {
      styles.push(projectionStyle)
    }
  })

  return {
    cells,
    rowAxisEntries,
    columnAxisEntries,
    styles,
  }
}
