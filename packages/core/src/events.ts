import type { EngineEvent } from "@bilig/protocol";

export class EngineEventBus {
  private readonly listeners = new Set<(event: EngineEvent) => void>();

  subscribe(listener: (event: EngineEvent) => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  emit(event: EngineEvent): void {
    for (const listener of this.listeners) {
      listener(event);
    }
  }
}
