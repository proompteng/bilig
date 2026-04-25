import type { GridGeometrySnapshot } from '../gridGeometry.js'

export class GridCameraStore {
  private snapshot: GridGeometrySnapshot | null = null
  private readonly listeners = new Set<() => void>()

  getSnapshot(): GridGeometrySnapshot | null {
    return this.snapshot
  }

  setSnapshot(next: GridGeometrySnapshot | null): void {
    if (this.snapshot === next || this.snapshot?.camera.seq === next?.camera.seq) {
      return
    }
    this.snapshot = next
    for (const listener of this.listeners) {
      listener()
    }
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }
}
