import { describe, expect, it, vi } from 'vitest'
import { ZeroWorkbookRevisionSync } from '../runtime-zero-revision-sync.js'

class FakeLiveView {
  readonly listeners = new Set<(value: unknown) => void>()
  destroy = vi.fn()

  constructor(readonly data: unknown) {}

  addListener(listener: (value: unknown) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  emit(value: unknown): void {
    this.listeners.forEach((listener) => {
      listener(value)
    })
  }
}

function createRevisionSyncHarness(initialData: unknown) {
  const view = new FakeLiveView(initialData)
  const onRevisionState = vi.fn()
  const zero = {
    materialize: vi.fn(() => view),
  }
  const sync = new ZeroWorkbookRevisionSync({
    zero,
    documentId: 'doc-1',
    onRevisionState,
  })
  return { onRevisionState, sync, view, zero }
}

describe('ZeroWorkbookRevisionSync', () => {
  it('normalizes valid workbook revision state from the initial view and live updates', () => {
    const { onRevisionState, sync, view, zero } = createRevisionSyncHarness({
      headRevision: 4,
      calculatedRevision: 3,
    })

    expect(zero.materialize).toHaveBeenCalledOnce()
    expect(onRevisionState).toHaveBeenLastCalledWith({
      headRevision: 4,
      calculatedRevision: 3,
    })

    view.emit({
      headRevision: 5,
      calculatedRevision: 5,
    })

    expect(onRevisionState).toHaveBeenLastCalledWith({
      headRevision: 5,
      calculatedRevision: 5,
    })

    sync.dispose()
    expect(view.listeners.size).toBe(0)
    expect(view.destroy).toHaveBeenCalledOnce()
  })

  it('rejects unsafe or impossible workbook revision state from Zero', () => {
    const { onRevisionState, sync, view } = createRevisionSyncHarness({
      headRevision: Number.NaN,
      calculatedRevision: 0,
    })

    expect(onRevisionState).toHaveBeenLastCalledWith(null)

    ;[
      { headRevision: -1, calculatedRevision: 0 },
      { headRevision: 2, calculatedRevision: -1 },
      { headRevision: 1.5, calculatedRevision: 1 },
      { headRevision: Number.MAX_SAFE_INTEGER + 1, calculatedRevision: 1 },
      { headRevision: 2, calculatedRevision: 3 },
      { headRevision: '2', calculatedRevision: 2 },
      null,
    ].forEach((value) => {
      view.emit(value)
      expect(onRevisionState).toHaveBeenLastCalledWith(null)
    })

    sync.dispose()
  })
})
