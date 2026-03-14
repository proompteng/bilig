import type { EngineEvent } from "@bilig/protocol";

export class EngineEventBus {
  private readonly listeners = new Set<(event: EngineEvent) => void>();
  private readonly cellIndexListeners = new Map<number, Set<() => void>>();
  private readonly addressListeners = new Map<string, Set<() => void>>();
  private readonly listenerIds = new WeakMap<() => void, number>();
  private listenerEpoch = 1;
  private listenerEpochs = new Uint32Array(64);
  private nextListenerId = 1;

  hasListeners(): boolean {
    return this.listeners.size > 0;
  }

  hasCellListeners(): boolean {
    return this.cellIndexListeners.size > 0 || this.addressListeners.size > 0;
  }

  hasAddressListeners(): boolean {
    return this.addressListeners.size > 0;
  }

  subscribe(listener: (event: EngineEvent) => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  subscribeCellIndex(cellIndex: number, listener: () => void): () => void {
    let listeners = this.cellIndexListeners.get(cellIndex);
    if (!listeners) {
      listeners = new Set();
      this.cellIndexListeners.set(cellIndex, listeners);
    }
    listeners.add(listener);
    return () => {
      const current = this.cellIndexListeners.get(cellIndex);
      if (!current) {
        return;
      }
      current.delete(listener);
      if (current.size === 0) {
        this.cellIndexListeners.delete(cellIndex);
      }
    };
  }

  subscribeCellAddress(qualifiedAddress: string, listener: () => void): () => void {
    let listeners = this.addressListeners.get(qualifiedAddress);
    if (!listeners) {
      listeners = new Set();
      this.addressListeners.set(qualifiedAddress, listeners);
    }
    listeners.add(listener);
    return () => {
      const current = this.addressListeners.get(qualifiedAddress);
      if (!current) {
        return;
      }
      current.delete(listener);
      if (current.size === 0) {
        this.addressListeners.delete(qualifiedAddress);
      }
    };
  }

  subscribeCells(cellIndices: readonly number[], qualifiedAddresses: readonly string[], listener: () => void): () => void {
    const unsubscribers = [
      ...cellIndices.map((cellIndex) => this.subscribeCellIndex(cellIndex, listener)),
      ...qualifiedAddresses.map((address) => this.subscribeCellAddress(address, listener))
    ];
    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe());
    };
  }

  emit(event: EngineEvent, changedCellIndices: readonly number[] | Uint32Array, resolveAddress?: (cellIndex: number) => string): void {
    for (const listener of this.listeners) {
      listener(event);
    }

    if (changedCellIndices.length === 0) {
      return;
    }

    this.beginListenerEpoch();
    for (let index = 0; index < changedCellIndices.length; index += 1) {
      const cellIndex = changedCellIndices[index]!;
      this.cellIndexListeners.get(cellIndex)?.forEach((listener) => {
        this.notifyListener(listener);
      });
      if (this.addressListeners.size === 0 || !resolveAddress) {
        continue;
      }
      const qualifiedAddress = resolveAddress(cellIndex);
      if (qualifiedAddress.length === 0 || qualifiedAddress.startsWith("!")) {
        continue;
      }
      this.addressListeners.get(qualifiedAddress)?.forEach((listener) => {
        this.notifyListener(listener);
      });
    }
  }

  emitAllWatched(event: EngineEvent): void {
    for (const listener of this.listeners) {
      listener(event);
    }

    if (this.cellIndexListeners.size === 0 && this.addressListeners.size === 0) {
      return;
    }

    this.beginListenerEpoch();
    this.cellIndexListeners.forEach((listeners) => {
      listeners.forEach((listener) => {
        this.notifyListener(listener);
      });
    });
    this.addressListeners.forEach((listeners) => {
      listeners.forEach((listener) => {
        this.notifyListener(listener);
      });
    });
  }

  private beginListenerEpoch(): void {
    this.listenerEpoch += 1;
    if (this.listenerEpoch === 0xffff_ffff) {
      this.listenerEpoch = 1;
      this.listenerEpochs.fill(0);
    }
  }

  private notifyListener(listener: () => void): void {
    const listenerId = this.getListenerId(listener);
    if (this.listenerEpochs[listenerId] === this.listenerEpoch) {
      return;
    }
    this.listenerEpochs[listenerId] = this.listenerEpoch;
    listener();
  }

  private getListenerId(listener: () => void): number {
    const existing = this.listenerIds.get(listener);
    if (existing !== undefined) {
      return existing;
    }
    const nextId = this.nextListenerId;
    this.nextListenerId += 1;
    if (nextId >= this.listenerEpochs.length) {
      const grown = new Uint32Array(this.listenerEpochs.length * 2);
      grown.set(this.listenerEpochs);
      this.listenerEpochs = grown;
    }
    this.listenerIds.set(listener, nextId);
    return nextId;
  }
}
