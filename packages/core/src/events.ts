import type { CellRangeRef, EngineEvent } from '@bilig/protocol'
import { parseCellAddress } from '@bilig/formula'
import type { EnginePatch } from './patches/patch-types.js'

interface NormalizedRange {
  sheetName: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
}

export interface EngineTrackedEvent {
  kind: EngineEvent['kind']
  invalidation: EngineEvent['invalidation']
  changedCellIndices: EngineEvent['changedCellIndices']
  patches?: readonly EnginePatch[]
  invalidatedRanges: EngineEvent['invalidatedRanges']
  invalidatedRows: EngineEvent['invalidatedRows']
  invalidatedColumns: EngineEvent['invalidatedColumns']
  metrics: EngineEvent['metrics']
  explicitChangedCount?: number
}

export class EngineEventBus {
  private readonly listeners = new Set<(event: EngineEvent) => void>()
  private readonly trackedListeners = new Set<(event: EngineTrackedEvent) => void>()
  private readonly cellIndexListeners = new Map<number, Set<() => void>>()
  private readonly addressListeners = new Map<string, Set<() => void>>()
  private readonly listenerIds = new WeakMap<() => void, number>()
  private listenerEpoch = 1
  private listenerEpochs = new Uint32Array(64)
  private nextListenerId = 1

  hasListeners(): boolean {
    return this.listeners.size > 0
  }

  hasTrackedListeners(): boolean {
    return this.trackedListeners.size > 0
  }

  hasCellListeners(): boolean {
    return this.cellIndexListeners.size > 0 || this.addressListeners.size > 0
  }

  hasAddressListeners(): boolean {
    return this.addressListeners.size > 0
  }

  subscribe(listener: (event: EngineEvent) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  subscribeTracked(listener: (event: EngineTrackedEvent) => void): () => void {
    this.trackedListeners.add(listener)
    return () => {
      this.trackedListeners.delete(listener)
    }
  }

  subscribeCellIndex(cellIndex: number, listener: () => void): () => void {
    let listeners = this.cellIndexListeners.get(cellIndex)
    if (!listeners) {
      listeners = new Set()
      this.cellIndexListeners.set(cellIndex, listeners)
    }
    listeners.add(listener)
    return () => {
      const current = this.cellIndexListeners.get(cellIndex)
      if (!current) {
        return
      }
      current.delete(listener)
      if (current.size === 0) {
        this.cellIndexListeners.delete(cellIndex)
      }
    }
  }

  subscribeCellAddress(qualifiedAddress: string, listener: () => void): () => void {
    let listeners = this.addressListeners.get(qualifiedAddress)
    if (!listeners) {
      listeners = new Set()
      this.addressListeners.set(qualifiedAddress, listeners)
    }
    listeners.add(listener)
    return () => {
      const current = this.addressListeners.get(qualifiedAddress)
      if (!current) {
        return
      }
      current.delete(listener)
      if (current.size === 0) {
        this.addressListeners.delete(qualifiedAddress)
      }
    }
  }

  subscribeCells(cellIndices: readonly number[], qualifiedAddresses: readonly string[], listener: () => void): () => void {
    const unsubscribers = [
      ...cellIndices.map((cellIndex) => this.subscribeCellIndex(cellIndex, listener)),
      ...qualifiedAddresses.map((address) => this.subscribeCellAddress(address, listener)),
    ]
    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe())
    }
  }

  emit(event: EngineEvent, changedCellIndices: readonly number[] | Uint32Array, resolveAddress?: (cellIndex: number) => string): void {
    for (const listener of this.listeners) {
      listener(event)
    }

    if (changedCellIndices.length === 0 && event.invalidatedRanges.length === 0) {
      return
    }

    this.beginListenerEpoch()
    for (let index = 0; index < changedCellIndices.length; index += 1) {
      const cellIndex = changedCellIndices[index]!
      this.cellIndexListeners.get(cellIndex)?.forEach((listener) => {
        this.notifyListener(listener)
      })
      if (this.addressListeners.size === 0 || !resolveAddress) {
        continue
      }
      const qualifiedAddress = resolveAddress(cellIndex)
      if (qualifiedAddress.length === 0 || qualifiedAddress.startsWith('!')) {
        continue
      }
      this.addressListeners.get(qualifiedAddress)?.forEach((listener) => {
        this.notifyListener(listener)
      })
    }

    if (event.invalidatedRanges.length === 0 || !resolveAddress) {
      return
    }

    const normalizedRanges = normalizeRanges(event.invalidatedRanges)
    if (normalizedRanges.length === 0) {
      return
    }

    this.cellIndexListeners.forEach((listeners, cellIndex) => {
      const qualifiedAddress = resolveAddress(cellIndex)
      if (!isQualifiedAddressInRanges(qualifiedAddress, normalizedRanges)) {
        return
      }
      listeners.forEach((listener) => {
        this.notifyListener(listener)
      })
    })

    this.addressListeners.forEach((listeners, qualifiedAddress) => {
      if (!isQualifiedAddressInRanges(qualifiedAddress, normalizedRanges)) {
        return
      }
      listeners.forEach((listener) => {
        this.notifyListener(listener)
      })
    })
  }

  emitAllWatched(event: EngineEvent): void {
    for (const listener of this.listeners) {
      listener(event)
    }

    if (this.cellIndexListeners.size === 0 && this.addressListeners.size === 0) {
      return
    }

    this.beginListenerEpoch()
    this.cellIndexListeners.forEach((listeners) => {
      listeners.forEach((listener) => {
        this.notifyListener(listener)
      })
    })
    this.addressListeners.forEach((listeners) => {
      listeners.forEach((listener) => {
        this.notifyListener(listener)
      })
    })
  }

  emitTracked(event: EngineTrackedEvent): void {
    for (const listener of this.trackedListeners) {
      listener(event)
    }
  }

  private beginListenerEpoch(): void {
    this.listenerEpoch += 1
    if (this.listenerEpoch === 0xffff_ffff) {
      this.listenerEpoch = 1
      this.listenerEpochs.fill(0)
    }
  }

  private notifyListener(listener: () => void): void {
    const listenerId = this.getListenerId(listener)
    if (this.listenerEpochs[listenerId] === this.listenerEpoch) {
      return
    }
    this.listenerEpochs[listenerId] = this.listenerEpoch
    listener()
  }

  private getListenerId(listener: () => void): number {
    const existing = this.listenerIds.get(listener)
    if (existing !== undefined) {
      return existing
    }
    const nextId = this.nextListenerId
    this.nextListenerId += 1
    if (nextId >= this.listenerEpochs.length) {
      const grown = new Uint32Array(this.listenerEpochs.length * 2)
      grown.set(this.listenerEpochs)
      this.listenerEpochs = grown
    }
    this.listenerIds.set(listener, nextId)
    return nextId
  }
}

function normalizeRanges(ranges: readonly CellRangeRef[]): NormalizedRange[] {
  return ranges.map((range) => {
    const start = parseCellAddress(range.startAddress, range.sheetName)
    const end = parseCellAddress(range.endAddress, range.sheetName)
    return {
      sheetName: range.sheetName,
      startRow: Math.min(start.row, end.row),
      endRow: Math.max(start.row, end.row),
      startCol: Math.min(start.col, end.col),
      endCol: Math.max(start.col, end.col),
    }
  })
}

function isQualifiedAddressInRanges(qualifiedAddress: string, ranges: readonly NormalizedRange[]): boolean {
  const separator = qualifiedAddress.indexOf('!')
  if (separator <= 0) {
    return false
  }
  const sheetName = qualifiedAddress.slice(0, separator)
  const address = qualifiedAddress.slice(separator + 1)
  const parsed = parseCellAddress(address, sheetName)
  for (let index = 0; index < ranges.length; index += 1) {
    const range = ranges[index]!
    if (range.sheetName !== sheetName) {
      continue
    }
    if (parsed.row >= range.startRow && parsed.row <= range.endRow && parsed.col >= range.startCol && parsed.col <= range.endCol) {
      return true
    }
  }
  return false
}
