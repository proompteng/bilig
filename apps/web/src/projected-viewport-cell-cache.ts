import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellRangeRef, type CellSnapshot, type CellStyleRecord, type Viewport } from '@bilig/protocol'
import { selectProjectedViewportKeysToEvict } from './projected-viewport-cache-pruning.js'
import {
  cellSnapshotSignature,
  isClearCellSnapshot,
  isOptimisticCellSnapshot,
  isOptimisticClearResurrection,
  isResetEmptySnapshot,
  prepareIncomingSnapshot,
  shouldKeepCurrentSnapshot,
} from './projected-viewport-cell-snapshot-policy.js'
import type { ProjectedViewportPatchApplicationResult } from './projected-viewport-patch-application.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from './workbook-optimistic-cell-flags.js'

const DEFAULT_STYLE_ID = 'style-0'
export const DEFAULT_MAX_CACHED_CELLS_PER_SHEET = 6000

interface CellSubscription {
  sheetName: string
  addresses: Set<string>
  listener: () => void
}

export interface ProjectedViewportCellSnapshotWriteResult {
  readonly acceptedSnapshot: CellSnapshot | null
  readonly changed: boolean
}

type CellSnapshotWriteOptions = { force?: boolean; forceOptimistic?: boolean; allowOptimisticClearResurrection?: boolean }

function normalizeMaxCachedCellsPerSheet(rawMaxCachedCellsPerSheet: number | undefined): number {
  if (typeof rawMaxCachedCellsPerSheet !== 'number' || !Number.isFinite(rawMaxCachedCellsPerSheet) || rawMaxCachedCellsPerSheet < 1) {
    return DEFAULT_MAX_CACHED_CELLS_PER_SHEET
  }
  return Math.floor(rawMaxCachedCellsPerSheet)
}

export class ProjectedViewportCellCache {
  private readonly cellSnapshots = new Map<string, CellSnapshot>()
  private readonly cellKeysBySheet = new Map<string, Set<string>>()
  private readonly cellStyles = new Map<string, CellStyleRecord>([[DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }]])
  private readonly cellSubscriptions = new Set<CellSubscription>()
  private readonly listeners = new Set<() => void>()
  private readonly knownSheets = new Set<string>()
  private readonly activeViewportKeysBySheet = new Map<string, Set<string>>()
  private readonly activeViewports = new Map<string, Viewport>()
  private readonly activeViewportRefCounts = new Map<string, number>()
  private readonly cellAccessTicks = new Map<string, number>()
  private nextCellAccessTick = 1

  private readonly options: {
    maxCachedCellsPerSheet: number
  }

  constructor(
    options: {
      maxCachedCellsPerSheet?: number
    } = {},
  ) {
    this.options = {
      maxCachedCellsPerSheet: normalizeMaxCachedCellsPerSheet(options.maxCachedCellsPerSheet),
    }
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  notifyListeners(): void {
    this.emitChange()
  }

  getSheet(sheetName: string):
    | {
        grid: {
          forEachCellEntry(listener: (cellIndex: number, row: number, col: number) => void): void
        }
      }
    | undefined {
    if (!this.knownSheets.has(sheetName)) {
      return undefined
    }
    const sheetCellKeys = this.cellKeysBySheet.get(sheetName)
    return {
      grid: {
        forEachCellEntry: (listener: (cellIndex: number, row: number, col: number) => void) => {
          let index = 0
          sheetCellKeys?.forEach((key) => {
            const snapshot = this.cellSnapshots.get(key)
            if (!snapshot) {
              return
            }
            const parsed = parseCellAddress(snapshot.address, snapshot.sheetName)
            listener(index++, parsed.row, parsed.col)
          })
        },
      },
    }
  }

  getPatchState(): {
    cellSnapshots: Map<string, CellSnapshot>
    cellKeysBySheet: Map<string, Set<string>>
    cellStyles: Map<string, CellStyleRecord>
    knownSheets: Set<string>
  } {
    return {
      cellSnapshots: this.cellSnapshots,
      cellKeysBySheet: this.cellKeysBySheet,
      cellStyles: this.cellStyles,
      knownSheets: this.knownSheets,
    }
  }

  peekCell(sheetName: string, address: string): CellSnapshot | undefined {
    const key = `${sheetName}!${address}`
    const snapshot = this.cellSnapshots.get(key)
    if (snapshot) {
      this.touchCellKey(key)
    }
    return snapshot
  }

  hasCellSnapshot(sheetName: string, address: string): boolean {
    return this.cellSnapshots.has(`${sheetName}!${address}`)
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    return this.peekCell(sheetName, address) ?? this.emptyCellSnapshot(sheetName, address)
  }

  forEachCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void): void {
    const start = parseCellAddress(range.startAddress, range.sheetName)
    const end = parseCellAddress(range.endAddress, range.sheetName)
    const startRow = Math.min(start.row, end.row)
    const endRow = Math.max(start.row, end.row)
    const startCol = Math.min(start.col, end.col)
    const endCol = Math.max(start.col, end.col)
    this.cellKeysBySheet.get(range.sheetName)?.forEach((key) => {
      const snapshot = this.cellSnapshots.get(key)
      if (!snapshot) {
        return
      }
      const parsed = parseCellAddress(snapshot.address, snapshot.sheetName)
      if (parsed.row < startRow || parsed.row > endRow || parsed.col < startCol || parsed.col > endCol) {
        return
      }
      listener(snapshot)
    })
  }

  forEachCachedOrVisibleCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void): void {
    const normalizedRange = normalizeRange(range)
    const visitedKeys = new Set<string>()
    const emitSnapshot = (snapshot: CellSnapshot) => {
      const key = `${snapshot.sheetName}!${snapshot.address}`
      if (visitedKeys.has(key)) {
        return
      }
      visitedKeys.add(key)
      listener(snapshot)
    }

    this.forEachCellSnapshotInRange(range, emitSnapshot)
    this.activeViewportKeysBySheet.get(range.sheetName)?.forEach((viewportKey) => {
      const viewport = this.activeViewports.get(viewportKey)
      if (!viewport) {
        return
      }
      const startRow = Math.max(normalizedRange.startRow, viewport.rowStart)
      const endRow = Math.min(normalizedRange.endRow, viewport.rowEnd)
      const startCol = Math.max(normalizedRange.startCol, viewport.colStart)
      const endCol = Math.min(normalizedRange.endCol, viewport.colEnd)
      if (startRow > endRow || startCol > endCol) {
        return
      }
      for (let row = startRow; row <= endRow; row += 1) {
        for (let col = startCol; col <= endCol; col += 1) {
          const address = formatAddress(row, col)
          emitSnapshot(this.peekCell(range.sheetName, address) ?? this.emptyCellSnapshot(range.sheetName, address))
        }
      }
    })
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    if (!styleId) {
      return this.cellStyles.get(DEFAULT_STYLE_ID)
    }
    return this.cellStyles.get(styleId) ?? this.cellStyles.get(DEFAULT_STYLE_ID)
  }

  upsertCellStyle(style: CellStyleRecord): void {
    this.cellStyles.set(style.id, style)
  }

  setCellSnapshot(
    snapshot: CellSnapshot,
    options: { force?: boolean; forceOptimistic?: boolean; allowOptimisticClearResurrection?: boolean } = {},
  ): boolean {
    return this.writeCellSnapshot(snapshot, options).changed
  }

  deleteCellSnapshot(sheetName: string, address: string): boolean {
    return this.deleteCellSnapshots([{ sheetName, address }])
  }

  deleteCellSnapshots(targets: readonly { readonly sheetName: string; readonly address: string }[]): boolean {
    const changedKeys = new Set<string>()
    for (const { sheetName, address } of targets) {
      const key = `${sheetName}!${address}`
      if (!this.cellSnapshots.delete(key)) continue
      changedKeys.add(key)
      this.cellAccessTicks.delete(key)
      const sheetCellKeys = this.cellKeysBySheet.get(sheetName)
      sheetCellKeys?.delete(key)
      if (sheetCellKeys?.size === 0) this.cellKeysBySheet.delete(sheetName)
    }
    if (changedKeys.size === 0) return false
    this.notifyCellSubscriptions(changedKeys)
    this.emitChange()
    return true
  }

  writeCellSnapshot(snapshot: CellSnapshot, options: CellSnapshotWriteOptions = {}): ProjectedViewportCellSnapshotWriteResult {
    const result = this.writeCellSnapshotWithoutNotifications(snapshot, options)
    if (result.changed) {
      this.notifyCellSubscriptions(new Set([`${snapshot.sheetName}!${snapshot.address}`]))
      this.emitChange()
    }
    return result
  }

  writeCellSnapshots(
    snapshots: readonly CellSnapshot[],
    options: CellSnapshotWriteOptions = {},
  ): ProjectedViewportCellSnapshotWriteResult[] {
    const changedKeys = new Set<string>()
    const results = snapshots.map((snapshot) => {
      const result = this.writeCellSnapshotWithoutNotifications(snapshot, options)
      if (result.changed) changedKeys.add(`${snapshot.sheetName}!${snapshot.address}`)
      return result
    })
    if (changedKeys.size > 0) {
      this.notifyCellSubscriptions(changedKeys)
      this.emitChange()
    }
    return results
  }

  private writeCellSnapshotWithoutNotifications(
    snapshot: CellSnapshot,
    options: CellSnapshotWriteOptions,
  ): ProjectedViewportCellSnapshotWriteResult {
    const key = `${snapshot.sheetName}!${snapshot.address}`
    const current = this.cellSnapshots.get(key)
    const incoming = current ? prepareIncomingSnapshot(current, snapshot) : snapshot
    if (!current && isResetEmptySnapshot(snapshot)) {
      this.knownSheets.add(snapshot.sheetName)
      return { acceptedSnapshot: null, changed: false }
    }
    if (current) {
      const forcedOptimisticClearHydration =
        options.force === true &&
        options.forceOptimistic === true &&
        options.allowOptimisticClearResurrection === true &&
        (current.flags & OPTIMISTIC_CELL_SNAPSHOT_FLAG) !== 0
      if (
        !forcedOptimisticClearHydration &&
        isClearCellSnapshot(current) &&
        !isClearCellSnapshot(incoming) &&
        current.version >= incoming.version
      ) {
        return { acceptedSnapshot: current, changed: false }
      }
      const shouldProtectCurrent =
        isOptimisticCellSnapshot(current) && isOptimisticClearResurrection(current, incoming) && !forcedOptimisticClearHydration
          ? true
          : options.force !== true || ((current.flags & OPTIMISTIC_CELL_SNAPSHOT_FLAG) !== 0 && options.forceOptimistic !== true)
      if (shouldProtectCurrent && shouldKeepCurrentSnapshot(current, incoming, { allowResetEmptyOverride: false })) {
        return { acceptedSnapshot: current, changed: false }
      }
      if (cellSnapshotSignature(current) === cellSnapshotSignature(incoming)) {
        return { acceptedSnapshot: current, changed: false }
      }
    }
    this.knownSheets.add(snapshot.sheetName)
    this.cellSnapshots.set(key, incoming)
    this.touchCellKey(key)
    this.sheetCellKeys(snapshot.sheetName).add(key)
    return { acceptedSnapshot: incoming, changed: true }
  }

  clearOptimisticCellFlagsForSheet(sheetName: string): CellSnapshot[] {
    const sheetCellKeys = this.cellKeysBySheet.get(sheetName)
    if (!sheetCellKeys) {
      return []
    }
    const changedKeys = new Set<string>()
    const changedSnapshots: CellSnapshot[] = []
    sheetCellKeys.forEach((key) => {
      const snapshot = this.cellSnapshots.get(key)
      if (!snapshot || (snapshot.flags & OPTIMISTIC_CELL_SNAPSHOT_FLAG) === 0) {
        return
      }
      const nextSnapshot = {
        ...snapshot,
        flags: snapshot.flags & ~OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      }
      this.cellSnapshots.set(key, nextSnapshot)
      changedSnapshots.push(nextSnapshot)
      changedKeys.add(key)
    })
    if (changedKeys.size === 0) {
      return []
    }
    this.notifyCellSubscriptions(changedKeys)
    this.emitChange()
    return changedSnapshots
  }

  markSheetKnown(sheetName: string): void {
    this.knownSheets.add(sheetName)
  }

  setKnownSheets(sheetNames: readonly string[]): string[] {
    if (sheetNames.length === this.knownSheets.size && sheetNames.every((sheetName) => this.knownSheets.has(sheetName))) {
      return []
    }
    const removedSheets = [...this.knownSheets].filter((sheetName) => !sheetNames.includes(sheetName))
    this.knownSheets.clear()
    sheetNames.forEach((sheetName) => this.knownSheets.add(sheetName))
    removedSheets.forEach((sheetName) => this.dropSheetCache(sheetName))
    this.emitChange()
    return removedSheets
  }

  getKnownSheetNames(): string[] {
    return [...this.knownSheets]
  }

  resetSheets(sheetNames: readonly string[]): void {
    sheetNames.forEach((sheetName) => {
      this.cellKeysBySheet.get(sheetName)?.forEach((key) => {
        this.cellSnapshots.delete(key)
        this.cellAccessTicks.delete(key)
      })
      this.cellKeysBySheet.delete(sheetName)
    })
    this.emitChange()
  }

  subscribeCells(sheetName: string, addresses: readonly string[], listener: () => void): () => void {
    const subscription: CellSubscription = {
      sheetName,
      addresses: new Set(addresses),
      listener,
    }
    this.cellSubscriptions.add(subscription)
    return () => {
      this.cellSubscriptions.delete(subscription)
    }
  }

  trackViewport(sheetName: string, viewport: Viewport): () => void {
    const viewportKey = `${sheetName}:${viewport.rowStart}:${viewport.rowEnd}:${viewport.colStart}:${viewport.colEnd}`
    this.activeViewports.set(viewportKey, viewport)
    this.activeViewportRefCounts.set(viewportKey, (this.activeViewportRefCounts.get(viewportKey) ?? 0) + 1)
    const sheetViewportKeys = this.activeViewportKeysBySheet.get(sheetName) ?? new Set<string>()
    sheetViewportKeys.add(viewportKey)
    this.activeViewportKeysBySheet.set(sheetName, sheetViewportKeys)
    let disposed = false
    return () => {
      if (disposed) {
        return
      }
      disposed = true
      const nextRefCount = (this.activeViewportRefCounts.get(viewportKey) ?? 0) - 1
      if (nextRefCount > 0) {
        this.activeViewportRefCounts.set(viewportKey, nextRefCount)
        return
      }
      this.activeViewportRefCounts.delete(viewportKey)
      this.activeViewports.delete(viewportKey)
      const nextSheetViewportKeys = this.activeViewportKeysBySheet.get(sheetName)
      nextSheetViewportKeys?.delete(viewportKey)
      if (nextSheetViewportKeys && nextSheetViewportKeys.size === 0) {
        this.activeViewportKeysBySheet.delete(sheetName)
      }
      this.pruneSheetCache(sheetName)
    }
  }

  applyPatchResult(
    sheetName: string,
    result: Pick<
      ProjectedViewportPatchApplicationResult,
      'changedKeys' | 'damage' | 'axisChanged' | 'columnsChanged' | 'rowsChanged' | 'freezeChanged' | 'mergesChanged'
    >,
  ): Pick<
    ProjectedViewportPatchApplicationResult,
    'damage' | 'axisChanged' | 'columnsChanged' | 'rowsChanged' | 'freezeChanged' | 'mergesChanged'
  > {
    this.pruneSheetCache(sheetName)
    this.notifyCellSubscriptions(result.changedKeys)
    if (result.damage.length > 0 || result.axisChanged || result.freezeChanged || result.mergesChanged) {
      this.emitChange()
    }
    return {
      damage: result.damage,
      axisChanged: result.axisChanged,
      columnsChanged: result.columnsChanged,
      rowsChanged: result.rowsChanged,
      freezeChanged: result.freezeChanged,
      mergesChanged: result.mergesChanged,
    }
  }

  touchCellKey(key: string): void {
    this.noteCellAccess(key)
  }

  private sheetCellKeys(sheetName: string): Set<string> {
    const existing = this.cellKeysBySheet.get(sheetName)
    if (existing) {
      return existing
    }
    const created = new Set<string>()
    this.cellKeysBySheet.set(sheetName, created)
    return created
  }

  private pruneSheetCache(sheetName: string): void {
    const sheetCellKeys = this.cellKeysBySheet.get(sheetName)
    if (!sheetCellKeys || sheetCellKeys.size <= this.options.maxCachedCellsPerSheet) {
      return
    }
    const activeViewportKeys = this.activeViewportKeysBySheet.get(sheetName)
    const activeViewports =
      activeViewportKeys && activeViewportKeys.size > 0
        ? [...activeViewportKeys]
            .map((key) => this.activeViewports.get(key))
            .filter((viewport): viewport is Viewport => viewport !== undefined)
        : []
    const pinnedKeys = new Set<string>()
    this.cellSubscriptions.forEach((subscription) => {
      if (subscription.sheetName !== sheetName) {
        return
      }
      subscription.addresses.forEach((address) => pinnedKeys.add(`${sheetName}!${address}`))
    })
    const keysToEvict = selectProjectedViewportKeysToEvict({
      sheetCellKeys: Array.from(sheetCellKeys),
      cellSnapshots: this.cellSnapshots,
      cellAccessTicks: this.cellAccessTicks,
      pinnedKeys,
      activeViewports,
      maxCachedCellsPerSheet: this.options.maxCachedCellsPerSheet,
    })
    keysToEvict.forEach((key) => {
      this.cellSnapshots.delete(key)
      this.cellAccessTicks.delete(key)
      sheetCellKeys.delete(key)
    })
  }

  private dropSheetCache(sheetName: string): void {
    this.cellKeysBySheet.get(sheetName)?.forEach((key) => {
      this.cellSnapshots.delete(key)
      this.cellAccessTicks.delete(key)
    })
    this.cellKeysBySheet.delete(sheetName)
    const viewportKeys = this.activeViewportKeysBySheet.get(sheetName)
    viewportKeys?.forEach((key) => {
      this.activeViewports.delete(key)
      this.activeViewportRefCounts.delete(key)
    })
    this.activeViewportKeysBySheet.delete(sheetName)
  }

  private noteCellAccess(key: string): void {
    this.cellAccessTicks.set(key, this.nextCellAccessTick++)
  }

  private notifyCellSubscriptions(changedKeys: ReadonlySet<string>): void {
    this.cellSubscriptions.forEach((subscription) => {
      for (const address of subscription.addresses) {
        if (changedKeys.has(`${subscription.sheetName}!${address}`)) {
          subscription.listener()
          return
        }
      }
    })
  }

  private emitChange(): void {
    this.listeners.forEach((listener) => listener())
  }

  private emptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
    return {
      sheetName,
      address,
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
    }
  }
}

function normalizeRange(range: CellRangeRef): {
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
} {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}
