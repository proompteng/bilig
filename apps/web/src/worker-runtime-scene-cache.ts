import type { WorkbookPaneScenePacket } from '@bilig/grid'

interface CachedSceneEntry {
  readonly batchId: number
  readonly generation: number
  readonly scenes: readonly WorkbookPaneScenePacket[]
}

export class WorkerRuntimeSceneCache {
  private readonly entries = new Map<string, CachedSceneEntry>()

  read(key: string, batchId: number): CachedSceneEntry | null {
    const entry = this.entries.get(key)
    if (!entry || entry.batchId !== batchId) {
      return null
    }
    return entry
  }

  write(key: string, batchId: number, buildScenes: (generation: number) => readonly WorkbookPaneScenePacket[]): CachedSceneEntry {
    const previous = this.entries.get(key)
    const generation = (previous?.generation ?? 0) + 1
    const next: CachedSceneEntry = {
      batchId,
      generation,
      scenes: buildScenes(generation),
    }
    this.entries.set(key, next)
    return next
  }

  reset(): void {
    this.entries.clear()
  }
}
