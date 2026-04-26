import { parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellSnapshot, type CellStyleRecord, type Viewport } from '@bilig/protocol'
import { selectProjectedViewportKeysToEvict } from './projected-viewport-cache-pruning.js'
import {
  cellSnapshotSignature,
  shouldKeepCurrentSnapshot,
  type ProjectedViewportPatchApplicationResult,
} from './projected-viewport-patch-application.js'

const DEFAULT_STYLE_ID = 'style-0'

interface CellSubscription {
  sheetName: string
  addresses: Set<string>
  listener: () => void
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
  private readonly cellAccessTicks = new Map<string, number>()
  private nextCellAccessTick = 1

  constructor(
    private readonly options: {
      maxCachedCellsPerSheet: number
    } = {
      maxCachedCellsPerSheet: 6000,
    },
  ) {}

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

  getCell(sheetName: string, address: string): CellSnapshot {
    return this.peekCell(sheetName, address) ?? this.emptyCellSnapshot(sheetName, address)
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    if (!styleId) {
      return this.cellStyles.get(DEFAULT_STYLE_ID)
    }
    return this.cellStyles.get(styleId) ?? this.cellStyles.get(DEFAULT_STYLE_ID)
  }

  setCellSnapshot(snapshot: CellSnapshot): void {
    const key = `${snapshot.sheetName}!${snapshot.address}`
    const current = this.cellSnapshots.get(key)
    if (current) {
      if (shouldKeepCurrentSnapshot(current, snapshot, { allowResetEmptyOverride: false })) {
        return
      }
      if (cellSnapshotSignature(current) === cellSnapshotSignature(snapshot)) {
        return
      }
    }
    this.knownSheets.add(snapshot.sheetName)
    this.cellSnapshots.set(key, snapshot)
    this.touchCellKey(key)
    this.sheetCellKeys(snapshot.sheetName).add(key)
    this.notifyCellSubscriptions(new Set([key]))
    this.emitChange()
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
    const sheetViewportKeys = this.activeViewportKeysBySheet.get(sheetName) ?? new Set<string>()
    sheetViewportKeys.add(viewportKey)
    this.activeViewportKeysBySheet.set(sheetName, sheetViewportKeys)
    return () => {
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
      'changedKeys' | 'damage' | 'axisChanged' | 'columnsChanged' | 'rowsChanged' | 'freezeChanged'
    >,
  ): Pick<ProjectedViewportPatchApplicationResult, 'damage' | 'axisChanged' | 'columnsChanged' | 'rowsChanged' | 'freezeChanged'> {
    this.pruneSheetCache(sheetName)
    this.notifyCellSubscriptions(result.changedKeys)
    if (result.damage.length > 0 || result.axisChanged || result.freezeChanged) {
      this.emitChange()
    }
    return {
      damage: result.damage,
      axisChanged: result.axisChanged,
      columnsChanged: result.columnsChanged,
      rowsChanged: result.rowsChanged,
      freezeChanged: result.freezeChanged,
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
    viewportKeys?.forEach((key) => this.activeViewports.delete(key))
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
