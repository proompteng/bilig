// @vitest-environment jsdom
import { act, useEffect } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { useWorkbookAgentThreadSummaries, type ZeroWorkbookAgentSource } from '../use-workbook-agent-durable-state.js'

class FakeLiveView {
  private readonly listeners = new Set<(value: unknown) => void>()

  constructor(readonly data: unknown) {}

  addListener(listener: (value: unknown) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  destroy(): void {}

  emit(value: unknown): void {
    for (const listener of this.listeners) {
      listener(value)
    }
  }
}

function createThreadSummary(overrides: Record<string, unknown>) {
  return {
    threadId: 'thr-1',
    scope: 'private',
    ownerUserId: 'alex@example.com',
    updatedAtUnixMs: 100,
    entryCount: 1,
    reviewQueueItemCount: 0,
    latestEntryText: null,
    ...overrides,
  }
}

afterEach(() => {
  vi.restoreAllMocks()
})

describe('useWorkbookAgentThreadSummaries', () => {
  it('keeps private live Zero summaries scoped to the current user and shared threads', async () => {
    const view = new FakeLiveView([
      createThreadSummary({ threadId: 'owned-private', ownerUserId: 'alex@example.com', reviewQueueItemCount: 1 }),
      createThreadSummary({ threadId: 'other-private', ownerUserId: 'casey@example.com', reviewQueueItemCount: 7 }),
      createThreadSummary({ threadId: 'shared', scope: 'shared', ownerUserId: 'casey@example.com', reviewQueueItemCount: 2 }),
    ])
    const zero = {
      materialize: vi.fn(() => view),
    } satisfies ZeroWorkbookAgentSource
    const host = document.createElement('div')
    let root: Root | null = null
    const observed: unknown[] = []

    function Harness() {
      const summaries = useWorkbookAgentThreadSummaries({
        currentUserId: 'alex@example.com',
        documentId: 'doc-1',
        enabled: true,
        zero,
      })
      useEffect(() => {
        observed.push(summaries.map((summary) => [summary.threadId, summary.reviewQueueItemCount]))
      }, [summaries])
      return null
    }

    await act(async () => {
      root = createRoot(host)
      root.render(<Harness />)
    })

    expect(zero.materialize).toHaveBeenCalledTimes(1)
    expect(observed.at(-1)).toEqual([
      ['owned-private', 1],
      ['shared', 2],
    ])

    await act(async () => {
      root?.unmount()
    })
  })
})
