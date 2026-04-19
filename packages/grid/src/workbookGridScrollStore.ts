export interface WorkbookGridScrollSnapshot {
  readonly tx: number
  readonly ty: number
}

const EMPTY_SCROLL_SNAPSHOT: WorkbookGridScrollSnapshot = Object.freeze({
  tx: 0,
  ty: 0,
})

export class WorkbookGridScrollStore {
  private snapshot: WorkbookGridScrollSnapshot = EMPTY_SCROLL_SNAPSHOT
  private readonly listeners = new Set<() => void>()

  getSnapshot(): WorkbookGridScrollSnapshot {
    return this.snapshot
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  setSnapshot(next: WorkbookGridScrollSnapshot): void {
    if (this.snapshot.tx === next.tx && this.snapshot.ty === next.ty) {
      return
    }
    this.snapshot = next
    for (const listener of this.listeners) {
      listener()
    }
  }
}
