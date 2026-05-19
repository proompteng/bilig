import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellRangeRef, type CellSnapshot, type Viewport } from '@bilig/protocol'
import { cellSnapshotSignature } from './projected-viewport-cell-snapshot-policy.js'

export interface NormalizedViewportRange {
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

export interface MaterializedOverlaySnapshot {
  readonly restoreSnapshot: boolean
  readonly snapshot: CellSnapshot
}

export interface PendingViewportRangeOverlay {
  readonly id: number
  readonly range: NormalizedViewportRange
  readonly apply: (snapshot: CellSnapshot) => CellSnapshot
  readonly materializedPreviousSnapshots: Map<string, MaterializedOverlaySnapshot>
}

export interface ProjectedViewportRangeOverlayStoreCallbacks {
  readonly deleteCellSnapshot: (sheetName: string, address: string) => void
  readonly forEachCachedOrVisibleCellSnapshotInRange: (range: CellRangeRef, listener: (snapshot: CellSnapshot) => void) => void
  readonly getCell: (sheetName: string, address: string) => CellSnapshot
  readonly hasCellSnapshot: (sheetName: string, address: string) => boolean
  readonly setCellSnapshot: (snapshot: CellSnapshot) => void
}

export class ProjectedViewportRangeOverlayStore {
  private readonly overlays: PendingViewportRangeOverlay[] = []
  private readonly suppressedOverlayMaxIdByCell = new Map<string, number>()
  private readonly resolvedSnapshotCache = new Map<
    string,
    { readonly baseSignature: string; readonly overlayRevision: number; readonly snapshot: CellSnapshot }
  >()
  private overlayRevision = 0
  private nextOverlayId = 1

  constructor(private readonly callbacks: ProjectedViewportRangeOverlayStoreCallbacks) {}

  hasOverlayForCell(sheetName: string, address: string): boolean {
    const cutoff = this.getSuppressedOverlayMaxId(sheetName, address)
    return this.overlays.some((overlay) => overlay.id > cutoff && overlayContainsAddress(overlay, sheetName, address))
  }

  apply(sheetName: string, address: string, snapshot: CellSnapshot): CellSnapshot {
    if (!this.hasOverlayForCell(sheetName, address)) {
      return snapshot
    }
    const key = snapshotOverlayKey(sheetName, address)
    const cutoff = this.getSuppressedOverlayMaxId(sheetName, address)
    const baseSignature = cellSnapshotSignature(snapshot)
    const cached = this.resolvedSnapshotCache.get(key)
    if (cached && cached.baseSignature === baseSignature && cached.overlayRevision === this.overlayRevision) {
      return cached.snapshot
    }
    let nextSnapshot = snapshot
    this.overlays.forEach((overlay) => {
      if (overlay.id <= cutoff || !overlayContainsAddress(overlay, sheetName, address)) {
        return
      }
      nextSnapshot = overlay.apply(nextSnapshot)
    })
    this.resolvedSnapshotCache.set(key, {
      baseSignature,
      overlayRevision: this.overlayRevision,
      snapshot: nextSnapshot,
    })
    return nextSnapshot
  }

  getSuppressedOverlayMaxId(sheetName: string, address: string): number {
    return this.suppressedOverlayMaxIdByCell.get(snapshotOverlayKey(sheetName, address)) ?? 0
  }

  restoreOverlaySuppression(sheetName: string, address: string, cutoff: number): void {
    const key = snapshotOverlayKey(sheetName, address)
    if (cutoff <= 0) {
      if (this.suppressedOverlayMaxIdByCell.delete(key)) {
        this.invalidateResolvedSnapshots()
      }
      return
    }
    if (this.suppressedOverlayMaxIdByCell.get(key) === cutoff) {
      return
    }
    this.suppressedOverlayMaxIdByCell.set(key, cutoff)
    this.invalidateResolvedSnapshots()
  }

  suppressExistingOverlaysForCell(sheetName: string, address: string): void {
    let cutoff = 0
    this.overlays.forEach((overlay) => {
      if (overlayContainsAddress(overlay, sheetName, address)) {
        cutoff = Math.max(cutoff, overlay.id)
      }
    })
    if (cutoff <= 0 || this.getSuppressedOverlayMaxId(sheetName, address) >= cutoff) {
      return
    }
    this.suppressedOverlayMaxIdByCell.set(snapshotOverlayKey(sheetName, address), cutoff)
    this.invalidateResolvedSnapshots()
  }

  register(range: CellRangeRef, apply: (snapshot: CellSnapshot) => CellSnapshot): (() => void) | null {
    const overlay = createPendingViewportRangeOverlay(this.nextOverlayId++, range, apply)
    this.overlays.push(overlay)
    this.invalidateResolvedSnapshots()
    this.materializeOverlayForCachedOrVisibleCells(overlay)
    let disposed = false
    return () => {
      if (disposed) {
        return
      }
      disposed = true
      this.remove(overlay)
    }
  }

  dropSheets(sheetNames: readonly string[]): void {
    if (sheetNames.length === 0) {
      return
    }
    const removedSheets = new Set(sheetNames)
    let changed = false
    if (this.overlays.length > 0) {
      for (let index = this.overlays.length - 1; index >= 0; index -= 1) {
        const overlay = this.overlays[index]
        if (overlay && removedSheets.has(overlay.range.sheetName)) {
          this.overlays.splice(index, 1)
          changed = true
        }
      }
    }
    for (const key of this.suppressedOverlayMaxIdByCell.keys()) {
      const separatorIndex = key.indexOf('!')
      const sheetName = separatorIndex >= 0 ? key.slice(0, separatorIndex) : key
      if (removedSheets.has(sheetName)) {
        this.suppressedOverlayMaxIdByCell.delete(key)
        changed = true
      }
    }
    if (changed || this.overlays.length > 0) {
      this.invalidateResolvedSnapshots()
    }
  }

  materializeViewport(sheetName: string, viewport: Viewport): void {
    this.overlays.forEach((overlay) => {
      const intersection = overlayViewportIntersection(overlay, sheetName, viewport)
      if (!intersection) {
        return
      }
      for (let row = intersection.startRow; row <= intersection.endRow; row += 1) {
        for (let col = intersection.startCol; col <= intersection.endCol; col += 1) {
          this.materializeAddress(sheetName, formatAddress(row, col))
        }
      }
    })
  }

  private remove(overlay: PendingViewportRangeOverlay): void {
    const overlayIndex = this.overlays.findIndex((candidate) => candidate.id === overlay.id)
    if (overlayIndex >= 0) {
      this.overlays.splice(overlayIndex, 1)
    }
    this.invalidateResolvedSnapshots()
    overlay.materializedPreviousSnapshots.forEach(({ restoreSnapshot, snapshot }) => {
      if (this.getSuppressedOverlayMaxId(snapshot.sheetName, snapshot.address) >= overlay.id) {
        return
      }
      const nextSnapshot = this.apply(snapshot.sheetName, snapshot.address, snapshot)
      if (restoreSnapshot || cellSnapshotSignature(nextSnapshot) !== cellSnapshotSignature(snapshot)) {
        this.callbacks.setCellSnapshot(nextSnapshot)
        return
      }
      this.callbacks.setCellSnapshot(snapshot)
      this.callbacks.deleteCellSnapshot(snapshot.sheetName, snapshot.address)
    })
    overlay.materializedPreviousSnapshots.clear()
  }

  private materializeOverlayForCachedOrVisibleCells(overlay: PendingViewportRangeOverlay): void {
    const range = {
      sheetName: overlay.range.sheetName,
      startAddress: formatAddress(overlay.range.startRow, overlay.range.startCol),
      endAddress: formatAddress(overlay.range.endRow, overlay.range.endCol),
    }
    this.callbacks.forEachCachedOrVisibleCellSnapshotInRange(range, (snapshot) => {
      this.materializeAddress(snapshot.sheetName, snapshot.address)
    })
  }

  private materializeAddress(sheetName: string, address: string): void {
    const key = snapshotOverlayKey(sheetName, address)
    let nextSnapshot = this.callbacks.getCell(sheetName, address)
    let changed = false
    this.overlays.forEach((overlay) => {
      if (!overlayContainsAddress(overlay, sheetName, address)) {
        return
      }
      const previousSnapshot = nextSnapshot
      nextSnapshot = overlay.apply(previousSnapshot)
      if (cellSnapshotSignature(previousSnapshot) === cellSnapshotSignature(nextSnapshot)) {
        return
      }
      changed = true
      if (!overlay.materializedPreviousSnapshots.has(key)) {
        overlay.materializedPreviousSnapshots.set(key, {
          restoreSnapshot: this.callbacks.hasCellSnapshot(sheetName, address) || shouldRestoreMaterializedOverlaySnapshot(previousSnapshot),
          snapshot: structuredClone(previousSnapshot),
        })
      }
    })
    if (changed && cellSnapshotSignature(this.callbacks.getCell(sheetName, address)) !== cellSnapshotSignature(nextSnapshot)) {
      this.callbacks.setCellSnapshot(nextSnapshot)
    }
  }

  private invalidateResolvedSnapshots(): void {
    this.overlayRevision += 1
    this.resolvedSnapshotCache.clear()
  }
}

export function normalizeViewportRange(range: CellRangeRef): NormalizedViewportRange {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    sheetName: range.sheetName,
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

export function viewportRangeCellCount(range: NormalizedViewportRange): number {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1)
}

export function createPendingViewportRangeOverlay(
  id: number,
  range: CellRangeRef,
  apply: (snapshot: CellSnapshot) => CellSnapshot,
): PendingViewportRangeOverlay {
  return {
    id,
    range: normalizeViewportRange(range),
    apply,
    materializedPreviousSnapshots: new Map(),
  }
}

export function overlayContainsAddress(overlay: PendingViewportRangeOverlay, sheetName: string, address: string): boolean {
  if (overlay.range.sheetName !== sheetName) {
    return false
  }
  const parsed = parseCellAddress(address, sheetName)
  return (
    parsed.row >= overlay.range.startRow &&
    parsed.row <= overlay.range.endRow &&
    parsed.col >= overlay.range.startCol &&
    parsed.col <= overlay.range.endCol
  )
}

export function overlayViewportIntersection(
  overlay: PendingViewportRangeOverlay,
  sheetName: string,
  viewport: Viewport,
): NormalizedViewportRange | null {
  if (overlay.range.sheetName !== sheetName) {
    return null
  }
  const startRow = Math.max(overlay.range.startRow, viewport.rowStart)
  const endRow = Math.min(overlay.range.endRow, viewport.rowEnd)
  const startCol = Math.max(overlay.range.startCol, viewport.colStart)
  const endCol = Math.min(overlay.range.endCol, viewport.colEnd)
  if (startRow > endRow || startCol > endCol) {
    return null
  }
  return {
    sheetName,
    startRow,
    endRow,
    startCol,
    endCol,
  }
}

export function snapshotOverlayKey(sheetName: string, address: string): string {
  return `${sheetName}!${address}`
}

function shouldRestoreMaterializedOverlaySnapshot(snapshot: CellSnapshot): boolean {
  return (
    snapshot.value.tag !== ValueTag.Empty ||
    snapshot.version !== 0 ||
    snapshot.flags !== 0 ||
    snapshot.formula !== undefined ||
    snapshot.input !== undefined ||
    snapshot.format !== undefined ||
    snapshot.styleId !== undefined ||
    snapshot.numberFormatId !== undefined
  )
}
