import type { GridSelectionSnapshot } from './gridTypes.js'

export interface GridSelectionPendingSyncInput {
  readonly sheetChanged: boolean
  readonly currentSnapshot: GridSelectionSnapshot
  readonly externalSnapshot: GridSelectionSnapshot
  readonly pendingLocalSnapshot: GridSelectionSnapshot | null
  readonly pendingBaseSnapshot: GridSelectionSnapshot | null
}

export interface GridSelectionPendingSyncResult {
  readonly keepCurrentSelection: boolean
  readonly pendingLocalSnapshot: GridSelectionSnapshot | null
  readonly pendingBaseSnapshot: GridSelectionSnapshot | null
}

export function gridSelectionSnapshotsEqual(left: GridSelectionSnapshot, right: GridSelectionSnapshot): boolean {
  return (
    left.sheetName === right.sheetName &&
    left.address === right.address &&
    left.kind === right.kind &&
    left.range.startAddress === right.range.startAddress &&
    left.range.endAddress === right.range.endAddress
  )
}

export function resolveGridSelectionPendingSync(input: GridSelectionPendingSyncInput): GridSelectionPendingSyncResult {
  if (input.sheetChanged) {
    return clearPendingAndSyncExternal()
  }

  if (input.pendingLocalSnapshot) {
    if (gridSelectionSnapshotsEqual(input.pendingLocalSnapshot, input.externalSnapshot)) {
      return gridSelectionSnapshotsEqual(input.currentSnapshot, input.externalSnapshot)
        ? {
            keepCurrentSelection: true,
            pendingBaseSnapshot: null,
            pendingLocalSnapshot: null,
          }
        : clearPendingAndSyncExternal()
    }

    if (gridSelectionSnapshotsEqual(input.currentSnapshot, input.pendingLocalSnapshot)) {
      if (input.pendingBaseSnapshot && gridSelectionSnapshotsEqual(input.externalSnapshot, input.pendingBaseSnapshot)) {
        return {
          keepCurrentSelection: true,
          pendingBaseSnapshot: input.pendingBaseSnapshot,
          pendingLocalSnapshot: input.pendingLocalSnapshot,
        }
      }
      if (input.pendingBaseSnapshot && gridSelectionSnapshotsEqual(input.pendingLocalSnapshot, input.pendingBaseSnapshot)) {
        return clearPendingAndSyncExternal()
      }
      return {
        keepCurrentSelection: true,
        pendingBaseSnapshot: input.pendingBaseSnapshot,
        pendingLocalSnapshot: input.pendingLocalSnapshot,
      }
    }
  }

  if (gridSelectionSnapshotsEqual(input.currentSnapshot, input.externalSnapshot)) {
    return {
      keepCurrentSelection: true,
      pendingBaseSnapshot: input.pendingBaseSnapshot,
      pendingLocalSnapshot: input.pendingLocalSnapshot,
    }
  }

  return clearPendingAndSyncExternal()
}

function clearPendingAndSyncExternal(): GridSelectionPendingSyncResult {
  return {
    keepCurrentSelection: false,
    pendingBaseSnapshot: null,
    pendingLocalSnapshot: null,
  }
}
