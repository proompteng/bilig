import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from './workbook-optimistic-cell-flags.js'

function snapshotValueKey(snapshot: CellSnapshot): string {
  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return `n:${snapshot.value.value}`
    case ValueTag.Boolean:
      return `b:${snapshot.value.value ? 1 : 0}`
    case ValueTag.String:
      return `s:${snapshot.value.stringId}:${snapshot.value.value}`
    case ValueTag.Error:
      return `e:${snapshot.value.code}`
    case ValueTag.Empty:
      return 'empty'
  }
  return 'empty'
}

export function isResetEmptySnapshot(snapshot: CellSnapshot): boolean {
  return (
    snapshot.value.tag === ValueTag.Empty &&
    snapshot.version === 0 &&
    snapshot.flags === 0 &&
    snapshot.formula === undefined &&
    snapshot.input === undefined &&
    snapshot.format === undefined &&
    snapshot.styleId === undefined &&
    snapshot.numberFormatId === undefined
  )
}

export function isClearCellSnapshot(snapshot: CellSnapshot): boolean {
  return snapshot.formula === undefined && snapshot.input === undefined && snapshot.value.tag === ValueTag.Empty
}

export function isOptimisticCellSnapshot(snapshot: CellSnapshot): boolean {
  return (snapshot.flags & OPTIMISTIC_CELL_SNAPSHOT_FLAG) !== 0
}

export function isOptimisticClearResurrection(current: CellSnapshot, incoming: CellSnapshot): boolean {
  return isClearCellSnapshot(current) && !isClearCellSnapshot(incoming)
}

export function cellSnapshotSignature(snapshot: CellSnapshot): string {
  return [
    snapshot.version,
    snapshot.flags,
    snapshot.formula ?? '',
    snapshot.format ?? '',
    snapshot.styleId ?? '',
    snapshot.numberFormatId ?? '',
    snapshot.input ?? '',
    snapshotValueKey(snapshot),
  ].join('|')
}

export function shouldKeepCurrentSnapshot(
  current: CellSnapshot,
  incoming: CellSnapshot,
  options: { readonly allowAuthoritativeClearOverride?: boolean; readonly allowResetEmptyOverride?: boolean } = {},
): boolean {
  if (isOptimisticCellSnapshot(current)) {
    if (incomingConfirmsOptimisticSnapshot(current, incoming)) {
      return false
    }
    if (options.allowAuthoritativeClearOverride === true && isClearCellSnapshot(incoming) && !isClearCellSnapshot(current)) {
      return false
    }
    if (
      current.formula !== undefined &&
      incoming.formula === undefined &&
      incoming.input === undefined &&
      incoming.value.tag !== ValueTag.Error
    ) {
      return true
    }
    if (isClearCellSnapshot(current) && !isClearCellSnapshot(incoming)) {
      return true
    }
    return current.version > incoming.version
  }

  const incomingIsResetEmptySnapshot = isResetEmptySnapshot(incoming)
  const allowResetEmptyOverride = options.allowResetEmptyOverride ?? true
  if (
    incomingIsResetEmptySnapshot &&
    !allowResetEmptyOverride &&
    current.version > incoming.version &&
    (current.formula !== undefined || current.input !== undefined)
  ) {
    return true
  }
  if (
    !incomingIsResetEmptySnapshot &&
    incoming.value.tag !== ValueTag.Error &&
    current.formula !== undefined &&
    incoming.formula === undefined &&
    incoming.input === undefined
  ) {
    return true
  }
  if (current.version > incoming.version) {
    if (incomingIsResetEmptySnapshot) {
      return !allowResetEmptyOverride
    }
    return true
  }
  if (current.version === incoming.version && isClearCellSnapshot(current) && !isClearCellSnapshot(incoming)) {
    return true
  }
  if (current.version < incoming.version) {
    return false
  }
  return (
    !incomingIsResetEmptySnapshot &&
    incoming.value.tag !== ValueTag.Error &&
    current.formula !== undefined &&
    incoming.formula === undefined
  )
}

function incomingConfirmsOptimisticSnapshot(current: CellSnapshot, incoming: CellSnapshot): boolean {
  if (isOptimisticCellSnapshot(incoming)) {
    return true
  }
  if (current.formula !== undefined) {
    return incoming.formula === current.formula
  }
  if (current.input !== undefined) {
    return incoming.input === current.input
  }
  if (current.value.tag === ValueTag.Empty) {
    return incoming.formula === undefined && incoming.input === undefined && incoming.value.tag === ValueTag.Empty
  }
  if (current.value.tag === ValueTag.Error && incoming.value.tag === ValueTag.Error) {
    return incoming.formula === undefined && incoming.input === undefined && incoming.value.code === current.value.code
  }
  return false
}

export function prepareIncomingSnapshot(
  current: CellSnapshot,
  incoming: CellSnapshot,
  options: { readonly releaseConfirmedOptimisticClear?: boolean } = {},
): CellSnapshot {
  if (!isOptimisticCellSnapshot(current) || !incomingConfirmsOptimisticSnapshot(current, incoming)) {
    return incoming
  }
  if (
    options.releaseConfirmedOptimisticClear === true &&
    isClearCellSnapshot(current) &&
    isClearCellSnapshot(incoming) &&
    !isOptimisticCellSnapshot(incoming)
  ) {
    return {
      ...incoming,
      flags: incoming.flags & ~OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: Math.max(current.version, incoming.version),
    }
  }
  return {
    ...incoming,
    flags: incoming.flags | OPTIMISTIC_CELL_SNAPSHOT_FLAG,
    version: Math.max(current.version, incoming.version),
  }
}

export function createClearTombstoneSnapshot(snapshot: CellSnapshot): CellSnapshot {
  return {
    sheetName: snapshot.sheetName,
    address: snapshot.address,
    value: { tag: ValueTag.Empty },
    flags: snapshot.flags & ~OPTIMISTIC_CELL_SNAPSHOT_FLAG,
    version: snapshot.version,
  }
}
