import type { EngineEvent } from "@bilig/protocol";

export class EngineEventBus {
  private readonly listeners = new Set<(event: EngineEvent) => void>();
  private readonly cellListeners = new Map<string, Set<() => void>>();

  hasListeners(): boolean {
    return this.listeners.size > 0;
  }

  hasCellListeners(): boolean {
    return this.cellListeners.size > 0;
  }

  subscribe(listener: (event: EngineEvent) => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  subscribeCell(qualifiedAddress: string, listener: () => void): () => void {
    let listeners = this.cellListeners.get(qualifiedAddress);
    if (!listeners) {
      listeners = new Set();
      this.cellListeners.set(qualifiedAddress, listeners);
    }
    listeners.add(listener);
    return () => {
      const current = this.cellListeners.get(qualifiedAddress);
      if (!current) {
        return;
      }
      current.delete(listener);
      if (current.size === 0) {
        this.cellListeners.delete(qualifiedAddress);
      }
    };
  }

  subscribeCells(qualifiedAddresses: Iterable<string>, listener: () => void): () => void {
    const unsubscribers = [...qualifiedAddresses].map((address) => this.subscribeCell(address, listener));
    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe());
    };
  }

  emit(event: EngineEvent, changedQualifiedAddresses: readonly string[] = []): void {
    for (const listener of this.listeners) {
      listener(event);
    }

    if (changedQualifiedAddresses.length === 0) {
      return;
    }

    const notified = new Set<() => void>();
    changedQualifiedAddresses.forEach((address) => {
      this.cellListeners.get(address)?.forEach((listener) => {
        if (notified.has(listener)) {
          return;
        }
        notified.add(listener);
        listener();
      });
    });
  }

  emitAllWatched(event: EngineEvent): void {
    this.emit(event, this.cellListeners.size === 0 ? [] : [...this.cellListeners.keys()]);
  }
}
