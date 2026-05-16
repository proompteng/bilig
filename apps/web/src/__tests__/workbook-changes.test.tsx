// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { useWorkbookChangesPane } from '../use-workbook-changes-pane.js'
import { normalizeWorkbookChangeRows } from '../workbook-changes-model.js'
import type { ZeroWorkbookChangeSource } from '../use-workbook-changes-pane.js'

interface MockZeroChangeHarness {
  readonly zero: ZeroWorkbookChangeSource
  readonly mutations: unknown[]
  readonly materializedView: {
    readonly data: unknown
    addListener(listener: (value: unknown) => void): () => void
    destroy(): void
  }
  emit(value: unknown): void
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return (typeof value === 'object' || typeof value === 'function') && value !== null
}

function getMutationName(mutation: unknown): string | null {
  if (!isRecord(mutation)) {
    return null
  }
  const mutator = mutation['mutator']
  if (!isRecord(mutator)) {
    return null
  }
  const mutatorName = mutator['mutatorName']
  return typeof mutatorName === 'string' ? mutatorName : null
}

function createMockZeroChangeHarness(initialValue: unknown): MockZeroChangeHarness {
  let currentValue = initialValue
  const listeners = new Set<(value: unknown) => void>()
  const mutations: unknown[] = []
  const materializedView = {
    get data() {
      return currentValue
    },
    addListener(listener: (value: unknown) => void) {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    destroy() {},
  }

  return {
    materializedView,
    zero: {
      materialize() {
        return materializedView
      },
      mutate(mutation: unknown) {
        mutations.push(mutation)
        return {}
      },
    },
    mutations,
    emit(value: unknown) {
      currentValue = value
      listeners.forEach((listener) => listener(value))
    },
  }
}

function ChangesHarness(props: {
  currentUserId: string
  documentId: string
  sheetNames: readonly string[]
  zero: MockZeroChangeHarness['zero']
  enabled: boolean
  pendingMutationSummary?: { readonly activeCount: number; readonly failedCount: number } | undefined
  localMutationEpoch?: number | undefined
  onHistoryMutationApplied?: (() => void | Promise<void>) | undefined
  onJump: (sheetName: string, address: string) => void
}) {
  const { canRedo, canUndo, changeCount, changesPanel, redoLatestChange, undoLatestChange } = useWorkbookChangesPane({
    documentId: props.documentId,
    currentUserId: props.currentUserId,
    sheetNames: props.sheetNames,
    zero: props.zero,
    enabled: props.enabled,
    pendingMutationSummary: props.pendingMutationSummary,
    localMutationEpoch: props.localMutationEpoch,
    onHistoryMutationApplied: props.onHistoryMutationApplied,
    onJump: props.onJump,
  })

  return (
    <div>
      <div data-testid="workbook-changes-count">{String(changeCount)}</div>
      <div data-testid="workbook-can-undo">{String(canUndo)}</div>
      <div data-testid="workbook-can-redo">{String(canRedo)}</div>
      <button data-testid="workbook-undo-latest" type="button" onClick={undoLatestChange} />
      <button data-testid="workbook-redo-latest" type="button" onClick={redoLatestChange} />
      {changesPanel}
    </div>
  )
}

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

describe('workbook changes', () => {
  it('drops materialized rows with event kinds outside the shared Zero event model', () => {
    expect(
      normalizeWorkbookChangeRows([
        {
          revision: 12,
          actorUserId: 'alex@example.com',
          clientMutationId: 'mutation-12',
          eventKind: 'legacyPatch',
          summary: 'Legacy patch',
          sheetId: 1,
          sheetName: 'Sheet1',
          anchorAddress: 'A1',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
          undoBundleJson: {
            kind: 'engineOps',
            ops: [],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAt: Date.parse('2026-04-06T13:44:00.000Z'),
        },
      ]),
    ).toEqual([])
  })

  it('drops materialized rows with invalid history revision metadata', () => {
    expect(
      normalizeWorkbookChangeRows([
        {
          revision: -1,
          actorUserId: 'alex@example.com',
          clientMutationId: 'mutation-negative',
          eventKind: 'setCellValue',
          summary: 'Updated Sheet1!A1',
          sheetId: 1,
          sheetName: 'Sheet1',
          anchorAddress: 'A1',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
          undoBundleJson: {
            kind: 'engineOps',
            ops: [],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAt: Date.parse('2026-04-06T13:45:00.000Z'),
        },
      ]),
    ).toEqual([])
  })

  it('renders authoritative change rows and jumps to available anchors', async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 12,
        actorUserId: 'amy.smith@example.com',
        clientMutationId: 'mutation-12',
        eventKind: 'fillRange',
        summary: 'Filled Sheet1!B1:B3',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'B1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B3' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T12:34:00.000Z'),
      },
      {
        revision: 11,
        actorUserId: 'guest:deadbeef',
        clientMutationId: null,
        eventKind: 'renderCommit',
        summary: 'Deleted sheet Archive',
        sheetId: null,
        sheetName: null,
        anchorAddress: null,
        rangeJson: null,
        undoBundleJson: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T12:30:00.000Z'),
      },
    ])
    const onJump = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={onJump}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    expect(host.querySelector("[data-testid='workbook-changes-count']")?.textContent).toBe('2')

    const rows = host.querySelectorAll<HTMLElement>("[data-testid='workbook-change-row']")
    expect(rows).toHaveLength(2)
    expect(host.textContent).toContain('Apr 6')
    expect(rows[0]?.textContent).toContain('Filled Sheet1!B1:B3')
    expect(rows[0]?.textContent).toContain('Amy Smith')
    expect(rows[0]?.textContent).not.toContain('r12')
    expect(rows[0]?.getAttribute('class')).toContain('border-b')
    expect(rows[0]?.getAttribute('class')).not.toContain('rounded-')
    expect(rows[1]?.textContent).toContain('Deleted sheet Archive')

    await act(async () => {
      rows[0]?.querySelector('button')?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onJump).toHaveBeenCalledWith('Sheet1', 'B1')

    await act(async () => {
      root.unmount()
    })
  })

  it('updates the visible change count when the Zero view publishes new rows', async () => {
    const changes = createMockZeroChangeHarness([])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={() => {}}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    expect(host.querySelector("[data-testid='workbook-changes-count']")?.textContent).toBe('0')
    const emptyState = host.querySelector("[data-testid='workbook-changes-empty-state']")
    expect(emptyState).not.toBeNull()
    expect(emptyState?.getAttribute('class')).toContain('min-h-0')
    expect(emptyState?.getAttribute('class')).not.toContain('min-h-[360px]')
    expect(host.textContent).toContain('No changes yet')
    expect(host.textContent).toContain('Workbook is up to date.')

    await act(async () => {
      changes.emit([
        {
          revision: 15,
          actorUserId: 'alex@example.com',
          clientMutationId: 'mutation-15',
          eventKind: 'setCellValue',
          summary: 'Updated Sheet1!C7',
          sheetId: 1,
          sheetName: 'Sheet1',
          anchorAddress: 'C7',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'C7', endAddress: 'C7' },
          undoBundleJson: {
            kind: 'engineOps',
            ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'C7' }],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAt: Date.now(),
        },
      ])
    })

    expect(host.querySelector("[data-testid='workbook-changes-count']")?.textContent).toBe('1')
    expect(host.querySelector("[data-testid='workbook-changes-empty-state']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('renders per-row revert controls for revertible authoritative changes', async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 21,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-21',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:12:00.000Z'),
      },
    ])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const onJump = vi.fn()

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={onJump}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    const revertButton = host.querySelector<HTMLButtonElement>("[data-testid='workbook-change-revert']")
    expect(revertButton).not.toBeNull()
    expect(revertButton?.getAttribute('aria-label')).toBe('Revert Updated Sheet1!A1')
    expect(changes.mutations).toHaveLength(0)

    await act(async () => {
      revertButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onJump).not.toHaveBeenCalled()
    expect(changes.mutations).toHaveLength(1)
    expect(changes.mutations[0]).toMatchObject({
      '~': 'MutateRequest',
      args: {
        documentId: 'doc-1',
        revision: 21,
      },
    })
    expect(getMutationName(changes.mutations[0])).toBe('workbook.revertChange')

    await act(async () => {
      root.unmount()
    })
  })

  it("exposes undo and redo availability from the current user's authoritative history", async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 31,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-31',
        eventKind: 'revertChange',
        summary: 'Reverted r30: Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }],
        },
        revertedByRevision: null,
        revertsRevision: 30,
        createdAt: Date.parse('2026-04-06T13:30:00.000Z'),
      },
      {
        revision: 30,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-30',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
        },
        revertedByRevision: 31,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:25:00.000Z'),
      },
      {
        revision: 29,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-29',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!B1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'B1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:20:00.000Z'),
      },
    ])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const onHistoryMutationApplied = vi.fn()

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onHistoryMutationApplied={onHistoryMutationApplied}
          onJump={() => {}}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    expect(host.querySelector("[data-testid='workbook-can-undo']")?.textContent).toBe('true')
    expect(host.querySelector("[data-testid='workbook-can-redo']")?.textContent).toBe('true')

    await act(async () => {
      host.querySelector("[data-testid='workbook-undo-latest']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      host.querySelector("[data-testid='workbook-redo-latest']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(changes.mutations).toHaveLength(2)
    expect(onHistoryMutationApplied).toHaveBeenCalledTimes(2)

    await act(async () => {
      root.unmount()
    })
  })

  it('hides undo when a collaborator has changed the same cell after the current user edit', async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 41,
        actorUserId: 'morgan@example.com',
        clientMutationId: 'mutation-41',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:30:00.000Z'),
      },
      {
        revision: 40,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-40',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:25:00.000Z'),
      },
    ])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={() => {}}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    expect(host.querySelector("[data-testid='workbook-can-undo']")?.textContent).toBe('false')

    await act(async () => {
      host.querySelector("[data-testid='workbook-undo-latest']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(changes.mutations).toHaveLength(0)

    await act(async () => {
      root.unmount()
    })
  })

  it('hides per-row revert when a later collaborator change overlaps the target', async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 51,
        actorUserId: 'morgan@example.com',
        clientMutationId: 'mutation-51',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:31:00.000Z'),
      },
      {
        revision: 50,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-50',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:30:00.000Z'),
      },
    ])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={() => {}}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    expect(host.querySelector("[data-testid='workbook-change-revert']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('uses structural Zero range scope when hiding stale undo and revert actions', async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 61,
        actorUserId: 'morgan@example.com',
        clientMutationId: 'mutation-61',
        eventKind: 'insertRows',
        summary: 'Inserted rows 3:4 on Sheet1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A3',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'rows' },
        undoBundleJson: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:41:00.000Z'),
      },
      {
        revision: 60,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-60',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!B3',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'B3',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'B3', endAddress: 'B3' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B3' }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:40:00.000Z'),
      },
    ])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={() => {}}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    expect(host.querySelector("[data-testid='workbook-can-undo']")?.textContent).toBe('false')
    expect(host.querySelector("[data-testid='workbook-change-revert']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('hides stale undo and revert actions when a malformed Zero range cannot be trusted', async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 63,
        actorUserId: 'morgan@example.com',
        clientMutationId: 'mutation-63',
        eventKind: 'insertRows',
        summary: 'Inserted rows 3:4 on Sheet1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A3',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'row-band' },
        undoBundleJson: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:43:00.000Z'),
      },
      {
        revision: 62,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-62',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!Z99',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'Z99',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'Z99', endAddress: 'Z99' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'Z99' }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:42:00.000Z'),
      },
    ])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={() => {}}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    expect(host.querySelector("[data-testid='workbook-can-undo']")?.textContent).toBe('false')
    expect(host.querySelector("[data-testid='workbook-change-revert']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('queues an undo shortcut while a workbook mutation is waiting for authoritative history', async () => {
    const changes = createMockZeroChangeHarness([])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={() => {}}
          pendingMutationSummary={{ activeCount: 1, failedCount: 0 }}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    await act(async () => {
      host.querySelector("[data-testid='workbook-undo-latest']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(changes.mutations).toHaveLength(0)

    await act(async () => {
      changes.emit([
        {
          revision: 30,
          actorUserId: 'alex@example.com',
          clientMutationId: 'mutation-30',
          eventKind: 'clearRange',
          summary: 'Cleared Sheet1!D12',
          sheetId: 1,
          sheetName: 'Sheet1',
          anchorAddress: 'D12',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'D12', endAddress: 'D12' },
          undoBundleJson: {
            kind: 'engineOps',
            ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'D12', value: 'delete-undo-redo' }],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAt: Date.parse('2026-04-06T13:25:00.000Z'),
        },
      ])
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={() => {}}
          pendingMutationSummary={{ activeCount: 0, failedCount: 0 }}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    expect(changes.mutations).toHaveLength(1)
    expect(getMutationName(changes.mutations[0])).toContain('undoLatestChange')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not undo stale history while a newer local mutation is not materialized yet', async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 29,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-29',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!D12',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'D12',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'D12', endAddress: 'D12' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'D12' }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-06T13:20:00.000Z'),
      },
    ])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          localMutationEpoch={1}
          onJump={() => {}}
          pendingMutationSummary={{ activeCount: 0, failedCount: 0 }}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          localMutationEpoch={2}
          onJump={() => {}}
          pendingMutationSummary={{ activeCount: 0, failedCount: 0 }}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    await act(async () => {
      host.querySelector("[data-testid='workbook-undo-latest']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(changes.mutations).toHaveLength(0)

    await act(async () => {
      changes.emit([
        {
          revision: 30,
          actorUserId: 'alex@example.com',
          clientMutationId: 'mutation-30',
          eventKind: 'clearRange',
          summary: 'Cleared Sheet1!D12',
          sheetId: 1,
          sheetName: 'Sheet1',
          anchorAddress: 'D12',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'D12', endAddress: 'D12' },
          undoBundleJson: {
            kind: 'engineOps',
            ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'D12', value: 'delete-undo-redo' }],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAt: Date.parse('2026-04-06T13:25:00.000Z'),
        },
        {
          revision: 29,
          actorUserId: 'alex@example.com',
          clientMutationId: 'mutation-29',
          eventKind: 'setCellValue',
          summary: 'Updated Sheet1!D12',
          sheetId: 1,
          sheetName: 'Sheet1',
          anchorAddress: 'D12',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'D12', endAddress: 'D12' },
          undoBundleJson: {
            kind: 'engineOps',
            ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'D12' }],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAt: Date.parse('2026-04-06T13:20:00.000Z'),
        },
      ])
    })

    expect(changes.mutations).toHaveLength(1)
    expect(getMutationName(changes.mutations[0])).toContain('undoLatestChange')

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps redo available when a longer undo chain still has older reverted entries after a newer redo', async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 25,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-25',
        eventKind: 'redoChange',
        summary: 'Redid r24: Reverted r21: Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
        },
        revertedByRevision: null,
        revertsRevision: 24,
        createdAt: Date.parse('2026-04-18T09:15:00.000Z'),
      },
      {
        revision: 24,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-24',
        eventKind: 'revertChange',
        summary: 'Reverted r21: Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'a1' }],
        },
        revertedByRevision: 25,
        revertsRevision: 21,
        createdAt: Date.parse('2026-04-18T09:14:00.000Z'),
      },
      {
        revision: 23,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-23',
        eventKind: 'revertChange',
        summary: 'Reverted r22: Updated Sheet1!B1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'B1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 'b1' }],
        },
        revertedByRevision: null,
        revertsRevision: 22,
        createdAt: Date.parse('2026-04-18T09:13:00.000Z'),
      },
      {
        revision: 22,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-22',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!B1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'B1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
        },
        revertedByRevision: 23,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-18T09:12:00.000Z'),
      },
      {
        revision: 21,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-21',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
        },
        revertedByRevision: 24,
        revertsRevision: null,
        createdAt: Date.parse('2026-04-18T09:11:00.000Z'),
      },
    ])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ChangesHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onJump={() => {}}
          sheetNames={['Sheet1']}
          zero={changes.zero}
        />,
      )
    })

    expect(host.querySelector("[data-testid='workbook-can-undo']")?.textContent).toBe('true')
    expect(host.querySelector("[data-testid='workbook-can-redo']")?.textContent).toBe('true')

    await act(async () => {
      root.unmount()
    })
  })
})
