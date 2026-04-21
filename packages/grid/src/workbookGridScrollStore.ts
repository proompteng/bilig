export interface WorkbookGridScrollSnapshot {
  readonly tx: number
  readonly ty: number
  readonly renderTx?: number | undefined
  readonly renderTy?: number | undefined
  readonly scrollLeft?: number | undefined
  readonly scrollTop?: number | undefined
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
    const currentRenderTx = this.snapshot.renderTx ?? this.snapshot.tx
    const currentRenderTy = this.snapshot.renderTy ?? this.snapshot.ty
    const nextRenderTx = next.renderTx ?? next.tx
    const nextRenderTy = next.renderTy ?? next.ty
    if (
      this.snapshot.tx === next.tx &&
      this.snapshot.ty === next.ty &&
      currentRenderTx === nextRenderTx &&
      currentRenderTy === nextRenderTy &&
      this.snapshot.scrollLeft === next.scrollLeft &&
      this.snapshot.scrollTop === next.scrollTop
    ) {
      return
    }
    this.snapshot = next
    for (const listener of this.listeners) {
      listener()
    }
  }
}
