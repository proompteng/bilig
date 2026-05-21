// @vitest-environment jsdom
import { act, useCallback, useEffect, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { toast } from 'sonner'
import { isWorkbookAgentCommandBundle, toWorkbookAgentReviewQueueItem, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import { ValueTag } from '@bilig/protocol'
import { WorkbookToastRegion } from '../WorkbookToastRegion.js'
import { resetWorkbookAgentClientTransportStateForTests } from '../workbook-agent-client.js'
import { clearWorkbookAgentPreviewCache } from '../workbook-agent-preview-cache.js'
import { useWorkbookAgentPane } from '../use-workbook-agent-pane.js'

function agentStorageKey(userId = 'alex@example.com'): string {
  return `bilig:workbook-agent:doc-1:${encodeURIComponent(userId)}`
}

async function flushToasts(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

class MockEventSource {
  static latest: MockEventSource | null = null
  readonly url: string
  private readonly listeners = new Map<string, Set<(event: Event) => void>>()

  constructor(url: string) {
    this.url = url
    MockEventSource.latest = this
  }

  close() {}

  addEventListener(type: string, listener: (event: Event) => void): void {
    const entries = this.listeners.get(type) ?? new Set()
    entries.add(listener)
    this.listeners.set(type, entries)
  }

  removeEventListener(type: string, listener: (event: Event) => void): void {
    const entries = this.listeners.get(type)
    if (!entries) {
      return
    }
    entries.delete(listener)
    if (entries.size === 0) {
      this.listeners.delete(type)
    }
  }

  emit(data: unknown): void {
    this.listeners.get('message')?.forEach((listener) => {
      listener(
        new MessageEvent('message', {
          data: JSON.stringify(data),
        }),
      )
    })
  }

  emitRaw(data: string): void {
    this.listeners.get('message')?.forEach((listener) => {
      listener(
        new MessageEvent('message', {
          data,
        }),
      )
    })
  }

  emitError(): void {
    this.listeners.get('error')?.forEach((listener) => {
      listener(new Event('error'))
    })
  }
}

function requestUrl(input: RequestInfo | URL): string {
  if (typeof input === 'string') {
    return input
  }
  if (input instanceof URL) {
    return input.href
  }
  return input.url
}

function requestBody(init: RequestInit | undefined): unknown {
  if (!init || typeof init.body !== 'string') {
    return null
  }
  return JSON.parse(init.body) as unknown
}

function requestMethod(init: RequestInit | undefined): string {
  return init?.method ?? 'GET'
}

function createDefaultWorkflowContext() {
  return {
    selection: {
      sheetName: 'Sheet1',
      address: 'A1',
    },
    viewport: {
      rowStart: 0,
      rowEnd: 10,
      colStart: 0,
      colEnd: 5,
    },
  }
}

function createReviewQueueItem(bundle: WorkbookAgentCommandBundle) {
  return toWorkbookAgentReviewQueueItem({
    bundle,
    reviewMode: bundle.sharedReview ? 'ownerReview' : 'manual',
    ...(bundle.sharedReview ? { sharedReview: bundle.sharedReview } : {}),
  })
}

function createSnapshot(overrides: Record<string, unknown> = {}) {
  const reviewBundleOverride = overrides['reviewBundle']
  const reviewQueueItemsOverride = overrides['reviewQueueItems']
  const { reviewBundle: _reviewBundle, ...restOverrides } = overrides
  const overrideEntries = Array.isArray(overrides['entries'])
    ? overrides['entries'].map((entry) =>
        typeof entry === 'object' && entry !== null && !('citations' in entry)
          ? {
              ...entry,
              citations: [],
            }
          : entry,
      )
    : undefined
  return {
    documentId: 'doc-1',
    threadId: 'thr-1',
    scope: 'private',
    executionPolicy: 'autoApplyAll',
    status: 'idle',
    activeTurnId: null,
    lastError: null,
    context: createDefaultWorkflowContext(),
    entries: [
      {
        id: 'assistant-1',
        kind: 'assistant',
        turnId: 'turn-1',
        text: '',
        phase: null,
        toolName: null,
        toolStatus: null,
        argumentsText: null,
        outputText: null,
        success: null,
        citations: [],
      },
    ],
    reviewQueueItems: Array.isArray(reviewQueueItemsOverride)
      ? reviewQueueItemsOverride
      : isWorkbookAgentCommandBundle(reviewBundleOverride)
        ? [createReviewQueueItem(reviewBundleOverride)]
        : [],
    executionRecords: [],
    workflowRuns: [],
    ...restOverrides,
    ...(overrideEntries ? { entries: overrideEntries } : {}),
  }
}

function createPreviewSummary(overrides: Record<string, unknown> = {}) {
  return {
    ranges: [],
    structuralChanges: [],
    cellDiffs: [],
    effectSummary: {
      displayedCellDiffCount: 0,
      truncatedCellDiffs: false,
      inputChangeCount: 0,
      formulaChangeCount: 0,
      styleChangeCount: 0,
      numberFormatChangeCount: 0,
      structuralChangeCount: 0,
    },
    ...overrides,
  }
}

function createThreadSummary(overrides: Record<string, unknown> = {}) {
  return {
    threadId: 'thr-1',
    scope: 'private',
    ownerUserId: 'alex@example.com',
    updatedAtUnixMs: 100,
    entryCount: 1,
    reviewQueueItemCount: typeof overrides['reviewQueueItemCount'] === 'number' ? overrides['reviewQueueItemCount'] : 0,
    latestEntryText: null,
    ...overrides,
  }
}

interface MockZeroAgentHarness {
  readonly zero: {
    materialize(query: unknown): {
      readonly data: unknown
      addListener(listener: (value: unknown) => void): () => void
      destroy(): void
    }
  }
}

function createMockZeroAgentHarness(input: {
  readonly initialThreadSummaries: unknown
  readonly initialWorkflowRuns: unknown
}): MockZeroAgentHarness {
  let threadSummaryValue = input.initialThreadSummaries
  let workflowRunValue = input.initialWorkflowRuns
  const threadSummaryListeners = new Set<(value: unknown) => void>()
  const workflowRunListeners = new Set<(value: unknown) => void>()
  let materializeCallCount = 0

  return {
    zero: {
      materialize(_query: unknown) {
        const isThreadSummaryQuery = materializeCallCount === 0
        materializeCallCount += 1
        return {
          get data() {
            return isThreadSummaryQuery ? threadSummaryValue : workflowRunValue
          },
          addListener(listener: (value: unknown) => void) {
            const listeners = isThreadSummaryQuery ? threadSummaryListeners : workflowRunListeners
            listeners.add(listener)
            return () => {
              listeners.delete(listener)
            }
          },
          destroy() {},
        }
      },
    },
  }
}

function AgentHarness(props: {
  readonly currentUserId?: string
  readonly previewCommandBundle?: Parameters<typeof useWorkbookAgentPane>[0]['previewCommandBundle']
  readonly syncAuthoritativeRevision?: Parameters<typeof useWorkbookAgentPane>[0]['syncAuthoritativeRevision']
  readonly zero?: Parameters<typeof useWorkbookAgentPane>[0]['zero']
  readonly zeroEnabled?: boolean
  readonly apiEnabled?: boolean
}) {
  const { agentError, agentPanel, clearAgentError } = useWorkbookAgentPane({
    currentUserId: props.currentUserId ?? 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext: () => ({
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    }),
    previewCommandBundle: props.previewCommandBundle ?? vi.fn(async () => createPreviewSummary()),
    ...(props.syncAuthoritativeRevision ? { syncAuthoritativeRevision: props.syncAuthoritativeRevision } : {}),
    ...(props.apiEnabled !== undefined ? { apiEnabled: props.apiEnabled } : {}),
    ...(props.zero ? { zero: props.zero } : {}),
    ...(props.zeroEnabled !== undefined ? { zeroEnabled: props.zeroEnabled } : {}),
  })

  return (
    <div>
      <WorkbookToastRegion
        toasts={
          agentError
            ? [
                {
                  id: 'agent-error',
                  tone: 'error',
                  message: agentError,
                  onDismiss: clearAgentError,
                },
              ]
            : []
        }
      />
      {agentPanel}
    </div>
  )
}

function UnstableLiveThreadSummaryHarness(props: { readonly zero: Parameters<typeof useWorkbookAgentPane>[0]['zero'] }) {
  const getContext = useCallback(() => createDefaultWorkflowContext(), [])
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    applyContext: () => undefined,
    previewCommandBundle,
    syncAuthoritativeRevision: () => undefined,
    zero: props.zero,
    zeroEnabled: true,
  })

  return <div>{agentPanel}</div>
}

function LaggyContextHarness() {
  const [selection, setSelection] = useState({
    sheetName: 'Sheet1',
    address: 'A1',
  })
  const selectionRef = useRef(selection)

  useEffect(() => {
    selectionRef.current = selection
  }, [selection])

  const getContext = useCallback(() => {
    const currentSelection = selectionRef.current
    return {
      selection: {
        sheetName: currentSelection.sheetName,
        address: currentSelection.address,
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    }
  }, [])

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    previewCommandBundle: vi.fn(async () => createPreviewSummary()),
  })

  return (
    <div>
      <button
        data-testid="switch-context"
        type="button"
        onClick={() => {
          setSelection({
            sheetName: 'sheet3',
            address: 'A1',
          })
        }}
      >
        Switch
      </button>
      {agentPanel}
    </div>
  )
}

function RapidSelectionContextHarness() {
  const [row, setRow] = useState(1)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      selection: {
        sheetName: 'Sheet1',
        address: `A${row}`,
      },
      viewport: {
        rowStart: row - 1,
        rowEnd: row + 9,
        colStart: 0,
        colEnd: 5,
      },
    }),
    [row],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: `Sheet1!A${row}`,
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-selection-context"
        type="button"
        onClick={() => {
          setRow((current) => current + 1)
        }}
      >
        {row}
      </button>
      {agentPanel}
    </div>
  )
}

function ToggleableContextSyncHarness() {
  const [enabled, setEnabled] = useState(true)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(() => createDefaultWorkflowContext(), [])

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled,
    getContext,
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="disable-agent-context"
        type="button"
        onClick={() => {
          setEnabled(false)
        }}
      >
        Disable
      </button>
      {agentPanel}
    </div>
  )
}

function VolatileRenderedContextHarness() {
  const [renderCount, setRenderCount] = useState(0)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: 7,
        batchId: 11,
        selection: null,
        visibleRange: null,
      },
    }),
    [],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="force-render"
        type="button"
        onClick={() => {
          setRenderCount((current) => current + 1)
        }}
      >
        {renderCount}
      </button>
      {agentPanel}
    </div>
  )
}

function VolatileRenderedBatchContextHarness() {
  const [batchId, setBatchId] = useState(11)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: 7,
        batchId,
        selection: null,
        visibleRange: null,
      },
    }),
    [batchId],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-render-batch"
        type="button"
        onClick={() => {
          setBatchId((current) => current + 1)
        }}
      >
        {batchId}
      </button>
      {agentPanel}
    </div>
  )
}

function RapidRenderedRevisionContextHarness() {
  const [capturedRevision, setCapturedRevision] = useState(7)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision,
        batchId: capturedRevision,
        selection: null,
        visibleRange: null,
      },
    }),
    [capturedRevision],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-render-revision"
        type="button"
        onClick={() => {
          setCapturedRevision((current) => current + 1)
        }}
      >
        {capturedRevision}
      </button>
      {agentPanel}
    </div>
  )
}

function RapidRenderedRangeContextHarness() {
  const [renderedVersion, setRenderedVersion] = useState(0)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: 20 + renderedVersion,
        batchId: 20 + renderedVersion,
        selection: null,
        visibleRange: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'A1',
          },
          rowCount: 1,
          columnCount: 1,
          cellCount: 1,
          truncated: false,
          rows: [
            [
              {
                address: 'A1',
                input: `rendered-${renderedVersion}`,
                value: {
                  tag: ValueTag.String,
                  value: `rendered-${renderedVersion}`,
                  stringId: renderedVersion,
                },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
          ],
        },
      },
    }),
    [renderedVersion],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-rendered-range"
        type="button"
        onClick={() => {
          setRenderedVersion((current) => current + 1)
        }}
      >
        {renderedVersion}
      </button>
      {agentPanel}
    </div>
  )
}

function VolatileRenderedStringIdContextHarness() {
  const [stringId, setStringId] = useState(1)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: 7,
        batchId: 11,
        selection: null,
        visibleRange: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'A1',
          },
          rowCount: 1,
          columnCount: 1,
          cellCount: 1,
          truncated: false,
          rows: [
            [
              {
                address: 'A1',
                input: 'same visible value',
                value: {
                  tag: ValueTag.String,
                  value: 'same visible value',
                  stringId,
                },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
          ],
        },
      },
    }),
    [stringId],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-string-id"
        type="button"
        onClick={() => {
          setStringId((current) => current + 1)
        }}
      >
        {stringId}
      </button>
      {agentPanel}
    </div>
  )
}

function VersionedContextRenderHarness(props: { readonly onBuildContext: (address: string) => void }) {
  const { onBuildContext } = props
  const [renderCount, setRenderCount] = useState(0)
  const [contextVersion, setContextVersion] = useState(0)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const address = `A${String(contextVersion + 1)}`
  const getContext = useCallback(() => {
    onBuildContext(address)
    return {
      selection: {
        sheetName: 'Sheet1',
        address,
      },
      viewport: {
        rowStart: contextVersion,
        rowEnd: contextVersion + 10,
        colStart: 0,
        colEnd: 5,
      },
    }
  }, [address, contextVersion, onBuildContext])

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: `Sheet1!${address}`,
    contextVersion,
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="force-versioned-context-render"
        type="button"
        onClick={() => {
          setRenderCount((current) => current + 1)
        }}
      >
        {renderCount}
      </button>
      <button
        data-testid="advance-versioned-context"
        type="button"
        onClick={() => {
          setContextVersion((current) => current + 1)
        }}
      >
        {contextVersion}
      </button>
      {agentPanel}
    </div>
  )
}

beforeEach(() => {
  vi.stubGlobal('EventSource', MockEventSource)
  window.sessionStorage.clear()
  clearWorkbookAgentPreviewCache()
})

afterEach(() => {
  toast.dismiss()
  vi.restoreAllMocks()
  resetWorkbookAgentClientTransportStateForTests()
  window.sessionStorage.clear()
  clearWorkbookAgentPreviewCache()
  document.body.innerHTML = ''
})

describe('workbook agent pane', () => {
  it('renders the assistant panel without the skill-card strip', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(JSON.stringify(createSnapshot()), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          }),
      ),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const input = host.querySelector("[data-testid='workbook-agent-input']")
    expect(input instanceof HTMLTextAreaElement).toBe(true)
    expect(input instanceof HTMLTextAreaElement ? input.value : '').toBe('')
    expect(host.textContent).not.toContain('Local Skills')
    expect(host.textContent).not.toContain('Inspect Selection')
    expect(host.textContent).not.toContain('Ask the assistant to inspect, edit, or restructure this workbook.')
    expect(host.querySelector("[data-testid='workbook-agent-empty-state']")).toBeNull()
    expect(host.textContent).not.toContain('No messages yet')
    expect(host.textContent).not.toContain('Active context: Sheet1!A1')
    expect(input instanceof HTMLTextAreaElement ? input.getAttribute('placeholder') : null).toBe('Ask the workbook assistant')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders durable thread summaries and workflow runs from Zero projections', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const zero = createMockZeroAgentHarness({
      initialThreadSummaries: [
        createThreadSummary({
          threadId: 'thr-1',
          scope: 'shared',
          ownerUserId: 'casey@example.com',
          latestEntryText: 'Completed workflow: Summarize Workbook',
        }),
      ],
      initialWorkflowRuns: [
        {
          runId: 'wf-zero-1',
          threadId: 'thr-1',
          startedByUserId: 'casey@example.com',
          workflowTemplate: 'summarizeWorkbook',
          title: 'Summarize Workbook',
          summary: 'Summarized workbook structure across 2 sheets.',
          status: 'completed',
          createdAtUnixMs: 1,
          updatedAtUnixMs: 2,
          completedAtUnixMs: 2,
          errorMessage: null,
          steps: [
            {
              stepId: 'inspect-workbook',
              label: 'Inspect workbook structure',
              status: 'completed',
              summary: 'Read durable workbook structure across 2 sheets.',
              updatedAtUnixMs: 1,
            },
          ],
          artifact: {
            kind: 'markdown',
            title: 'Workbook Summary',
            text: '## Workbook Summary',
          },
        },
      ],
    })
    sessionStorage.setItem(agentStorageKey(), JSON.stringify({ threadId: 'thr-1' }))
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1')) {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1', workflowRuns: [] })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness zero={zero.zero} zeroEnabled />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-scope-private']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-scope-shared']")).toBeNull()
    expect(host.textContent).toContain('Workflows')
    expect(host.textContent).toContain('Summarize Workbook')
    expect(host.textContent).toContain('Workbook Summary')
    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads') && requestMethod(init) === 'GET'),
    ).toHaveLength(0)

    await act(async () => {
      root.unmount()
    })
  })

  it('hides applied preview system timeline entries from the assistant panel', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(agentStorageKey(), JSON.stringify({ threadId: 'thr-1' }))
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads')) {
          return new Response(JSON.stringify([]), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        if (url.endsWith('/chat/threads/thr-1')) {
          return new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: 'system-apply:run-1',
                    kind: 'system',
                    turnId: 'turn-1',
                    text: 'Applied workbook change set at revision r7: Write cells in Sheet1!B2',
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                    citations: [
                      {
                        kind: 'range',
                        sheetName: 'Sheet1',
                        startAddress: 'B2',
                        endAddress: 'B2',
                        role: 'target',
                      },
                      {
                        kind: 'revision',
                        revision: 7,
                      },
                    ],
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          )
        }
        throw new Error(`Unexpected fetch to ${url}`)
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.textContent).not.toContain('Applied workbook change set at revision r7')
    expect(host.textContent).not.toContain('Sheet1!B2')
    expect(host.textContent).not.toContain('r7')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders durable workflow runs in the assistant panel', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(agentStorageKey(), JSON.stringify({ threadId: 'thr-1' }))
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                workflowRuns: [
                  {
                    runId: 'wf-1',
                    threadId: 'thr-1',
                    startedByUserId: 'alex@example.com',
                    workflowTemplate: 'summarizeWorkbook',
                    title: 'Summarize Workbook',
                    summary: 'Summarized workbook structure across 2 sheets.',
                    status: 'completed',
                    createdAtUnixMs: 1,
                    updatedAtUnixMs: 2,
                    completedAtUnixMs: 2,
                    errorMessage: null,
                    steps: [
                      {
                        stepId: 'inspect-workbook',
                        label: 'Inspect workbook structure',
                        status: 'completed',
                        summary: 'Read durable workbook structure across 2 sheets.',
                        updatedAtUnixMs: 1,
                      },
                      {
                        stepId: 'draft-summary',
                        label: 'Draft summary artifact',
                        status: 'completed',
                        summary: 'Prepared the durable workbook summary artifact for the thread.',
                        updatedAtUnixMs: 2,
                      },
                    ],
                    artifact: {
                      kind: 'markdown',
                      title: 'Workbook Summary',
                      text: '## Workbook Summary\n\nSheets: 2\n### Sheets\n- Sheet1',
                    },
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          ),
      ),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.textContent).toContain('Workflows')
    expect(host.textContent).toContain('Summarize Workbook')
    expect(host.textContent).toContain('Inspect workbook structure')
    expect(host.textContent).toContain('Workbook Summary')
    expect(host.textContent).toContain('Sheets: 2')
    expect(host.textContent).toContain('Done')

    await act(async () => {
      root.unmount()
    })
  })

  it('loads durable thread summaries into the assistant panel', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads')) {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-shared',
              scope: 'shared',
              entryCount: 4,
              reviewQueueItemCount: 1,
              latestEntryText: 'Applied workbook change set at revision r7',
            }),
            createThreadSummary({
              threadId: 'thr-private',
              scope: 'private',
              entryCount: 2,
              latestEntryText: 'Review item queued',
            }),
          ]),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-shared']")).not.toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-private']")).not.toBeNull()
    expect(host.textContent).toContain('Shared')
    expect(host.textContent).toContain('Review')
    expect(host.textContent).toContain('4 items')
    expect(host.textContent).toContain('Applied workbook change set at revision r7')

    await act(async () => {
      root.unmount()
    })
  })

  it('switches to a durable thread from the summary strip', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-2',
              scope: 'shared',
              entryCount: 3,
            }),
          ]),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/chat/threads/thr-2')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-2',
              scope: 'shared',
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const threadButton = host.querySelector("[data-testid='workbook-agent-thread-thr-2']")
    expect(threadButton instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(threadButton instanceof HTMLButtonElement)) {
        throw new Error('Thread button not found')
      }
      threadButton.click()
    })

    expect(MockEventSource.latest?.url).toBe('/v2/documents/doc-1/chat/threads/thr-2/events')
    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-2']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-scope-private']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-scope-shared']")).toBeNull()
    expect(fetchSpy).toHaveBeenCalledWith('/v2/documents/doc-1/chat/threads/thr-2')

    await act(async () => {
      root.unmount()
    })
  })

  it('hides the summary strip when it would only repeat the active thread', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-1',
              scope: 'private',
              entryCount: 64,
              latestEntryText: 'Done - operating plan now exists as a sheet.',
            }),
          ]),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/chat/threads/thr-1')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-1',
              scope: 'private',
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-1']")).toBeNull()
    expect(host.textContent).not.toContain('64 items')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not render thread scope controls', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
          return new Response(JSON.stringify([]), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        throw new Error(`Unexpected fetch to ${url}`)
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-scope-private']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-scope-shared']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('restores a new-thread draft after remount', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads')) {
          return new Response(JSON.stringify([]), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        throw new Error(`Unexpected fetch to ${url}`)
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const input = host.querySelector("[data-testid='workbook-agent-input']")
    expect(input instanceof HTMLTextAreaElement).toBe(true)

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error('Agent input not found')
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, 'set') : null
      if (typeof valueSetter !== 'function') {
        throw new Error('Textarea value setter not found')
      }
      Reflect.apply(valueSetter, input, ['Persisted draft'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })

    await act(async () => {
      root.unmount()
    })

    const remountRoot = createRoot(host)
    await act(async () => {
      remountRoot.render(<AgentHarness />)
    })

    const restoredInput = host.querySelector("[data-testid='workbook-agent-input']")
    expect(restoredInput instanceof HTMLTextAreaElement ? restoredInput.value : null).toBe('Persisted draft')

    await act(async () => {
      remountRoot.unmount()
    })
  })

  it('submits the draft on Enter from the chat composer', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify(createSnapshot({ entries: [] })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/turns')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: 'inProgress',
              activeTurnId: 'turn-1',
              entries: [
                {
                  id: 'optimistic-user:turn-1',
                  kind: 'user',
                  turnId: 'turn-1',
                  text: 'Summarize this sheet',
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const input = host.querySelector("[data-testid='workbook-agent-input']")
    expect(input instanceof HTMLTextAreaElement).toBe(true)

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error('Agent input not found')
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, 'set') : null
      if (typeof valueSetter !== 'function') {
        throw new Error('Textarea value setter not found')
      }
      Reflect.apply(valueSetter, input, ['Summarize this sheet'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          key: 'Enter',
        }),
      )
    })

    const turnCall = fetchSpy.mock.calls.find(([requestInput]) => requestUrl(requestInput).endsWith('/chat/threads/thr-1/turns'))
    expect(turnCall?.[0]).toBe('/v2/documents/doc-1/chat/threads/thr-1/turns')
    expect(host.textContent).not.toContain('Reviewing workbook context and drafting a response.')
    const nextInput = host.querySelector("[data-testid='workbook-agent-input']")
    expect(nextInput instanceof HTMLTextAreaElement ? nextInput.value : null).toBe('')

    await act(async () => {
      root.unmount()
    })
  })

  it('coalesces rapid first prompt submits into one assistant session and one turn', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    let resolveSession: (() => void) | null = null
    const fetchSpy = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return Promise.resolve(
          new Response(JSON.stringify([]), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          }),
        )
      }
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'POST') {
        return new Promise<Response>((resolve) => {
          resolveSession = () => {
            resolve(
              new Response(JSON.stringify(createSnapshot({ entries: [] })), {
                status: 200,
                headers: { 'content-type': 'application/json' },
              }),
            )
          }
        })
      }
      if (url.endsWith('/chat/threads/thr-1/turns')) {
        return Promise.resolve(
          new Response(
            JSON.stringify(
              createSnapshot({
                status: 'inProgress',
                activeTurnId: 'turn-1',
                entries: [
                  {
                    id: 'user:turn-1',
                    kind: 'user',
                    turnId: 'turn-1',
                    text: 'Summarize this sheet',
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          ),
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const callsTo = (suffix: string, method: string) =>
      fetchSpy.mock.calls.filter(([requestInput, init]) => requestUrl(requestInput).endsWith(suffix) && requestMethod(init) === method)

    try {
      await act(async () => {
        root.render(<AgentHarness />)
      })

      const input = host.querySelector("[data-testid='workbook-agent-input']")
      const submit = host.querySelector("[data-testid='workbook-agent-send']")
      expect(input instanceof HTMLTextAreaElement).toBe(true)
      expect(submit instanceof HTMLButtonElement).toBe(true)

      await act(async () => {
        if (!(input instanceof HTMLTextAreaElement) || !(submit instanceof HTMLButtonElement)) {
          throw new Error('Assistant composer not found')
        }
        const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')
        const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, 'set') : null
        if (typeof valueSetter !== 'function') {
          throw new Error('Textarea value setter not found')
        }
        Reflect.apply(valueSetter, input, ['Summarize this sheet'])
        input.dispatchEvent(new Event('input', { bubbles: true }))
        submit.click()
        submit.click()
      })

      expect(callsTo('/chat/threads', 'POST')).toHaveLength(1)
      expect(callsTo('/chat/threads/thr-1/turns', 'POST')).toHaveLength(0)

      await act(async () => {
        resolveSession?.()
        await Promise.resolve()
        await Promise.resolve()
      })

      expect(callsTo('/chat/threads', 'POST')).toHaveLength(1)
      expect(callsTo('/chat/threads/thr-1/turns', 'POST')).toHaveLength(1)
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('submits follow-up prompts through the durable thread route when a thread is already active', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify(createSnapshot({ entries: [] })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/turns')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: 'inProgress',
              activeTurnId: 'turn-2',
              entries: [
                {
                  id: 'optimistic-user:turn-2',
                  kind: 'user',
                  turnId: 'turn-2',
                  text: 'Continue working',
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const input = host.querySelector("[data-testid='workbook-agent-input']")
    expect(input instanceof HTMLTextAreaElement).toBe(true)

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error('Agent input not found')
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, 'set') : null
      if (typeof valueSetter !== 'function') {
        throw new Error('Textarea value setter not found')
      }
      Reflect.apply(valueSetter, input, ['Continue working'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          key: 'Enter',
        }),
      )
    })

    const turnCall = fetchSpy.mock.calls.find(([requestInput]) => requestUrl(requestInput).endsWith('/chat/threads/thr-1/turns'))
    expect(turnCall?.[0]).toBe('/v2/documents/doc-1/chat/threads/thr-1/turns')
    expect(host.textContent).not.toContain('Reviewing workbook context and drafting a response.')

    await act(async () => {
      root.unmount()
    })
  })

  it('restores the draft and shows the server message when a turn request fails', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify([createThreadSummary()]), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/turns')) {
        return new Response(
          JSON.stringify({
            message: 'Prompt rejected by server',
          }),
          {
            status: 422,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const input = host.querySelector("[data-testid='workbook-agent-input']")
    expect(input instanceof HTMLTextAreaElement).toBe(true)

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error('Agent input not found')
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, 'set') : null
      if (typeof valueSetter !== 'function') {
        throw new Error('Textarea value setter not found')
      }
      Reflect.apply(valueSetter, input, ['Continue working'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          key: 'Enter',
        }),
      )
    })
    await flushToasts()

    const turnCall = fetchSpy.mock.calls.find(([requestInput]) => requestUrl(requestInput).endsWith('/chat/threads/thr-1/turns'))
    expect(requestBody(turnCall?.[1])).toEqual({
      prompt: 'Continue working',
      context: createDefaultWorkflowContext(),
    })
    expect(host.textContent).toContain('Prompt rejected by server')
    expect(host.textContent).not.toContain('Workbook agent request failed with status 422')
    const restoredInput = host.querySelector("[data-testid='workbook-agent-input']")
    expect(restoredInput instanceof HTMLTextAreaElement ? restoredInput.value : null).toBe('Continue working')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not inject a synthetic progress row before the turn request resolves', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    let resolveTurnResponse: ((response: Response) => void) | null = null
    const turnResponse = new Promise<Response>((resolve) => {
      resolveTurnResponse = resolve
    })
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot()), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/turns')) {
        return await turnResponse
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const input = host.querySelector("[data-testid='workbook-agent-input']")
    expect(input instanceof HTMLTextAreaElement).toBe(true)

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error('Agent input not found')
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, 'set') : null
      if (typeof valueSetter !== 'function') {
        throw new Error('Textarea value setter not found')
      }
      Reflect.apply(valueSetter, input, ['yo'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          key: 'Enter',
        }),
      )
      await Promise.resolve()
    })

    expect(host.textContent).toContain('yo')
    expect(host.textContent).not.toContain('Reviewing workbook context and drafting a response.')
    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).toBeNull()

    await act(async () => {
      resolveTurnResponse?.(
        new Response(
          JSON.stringify(
            createSnapshot({
              status: 'inProgress',
              activeTurnId: 'turn-3',
              entries: [
                {
                  id: 'optimistic-user:turn-3',
                  kind: 'user',
                  turnId: 'turn-3',
                  text: 'yo',
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        ),
      )
      await Promise.resolve()
    })

    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).not.toBeNull()
    expect(host.textContent).toContain('Thinking')

    await act(async () => {
      root.unmount()
    })
  })

  it('uses the composer button to interrupt an active turn', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: 'inProgress',
              activeTurnId: 'turn-1',
              entries: [
                {
                  id: 'assistant-1',
                  kind: 'assistant',
                  turnId: 'turn-1',
                  text: 'Working',
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/interrupt')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: 'idle',
              activeTurnId: null,
              entries: [
                {
                  id: 'assistant-1',
                  kind: 'assistant',
                  turnId: 'turn-1',
                  text: 'Working',
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const button = host.querySelector("[data-testid='workbook-agent-send']")
    expect(button instanceof HTMLButtonElement).toBe(true)
    expect(button instanceof HTMLButtonElement ? button.getAttribute('aria-label') : null).toBe('Stop')

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error('Agent button not found')
      }
      button.click()
    })

    const interruptCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/chat/threads/thr-1/interrupt'))
    expect(interruptCall?.[0]).toBe('/v2/documents/doc-1/chat/threads/thr-1/interrupt')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders structured workbook comprehension tool results in the rail', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: 'tool-search',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: null,
                    phase: null,
                    toolName: 'search_workbook',
                    toolStatus: 'completed',
                    argumentsText: '{"query":"gross margin"}',
                    outputText: JSON.stringify({
                      query: 'gross margin',
                      summary: { matchCount: 1, truncated: false },
                      matches: [
                        {
                          kind: 'cell',
                          sheetName: 'Sheet1',
                          address: 'A2',
                          snippet: 'Gross Margin',
                          reasons: ['value'],
                          score: 65,
                        },
                      ],
                    }),
                    success: true,
                  },
                  {
                    id: 'tool-issues',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: null,
                    phase: null,
                    toolName: 'find_formula_issues',
                    toolStatus: 'completed',
                    argumentsText: '{}',
                    outputText: JSON.stringify({
                      summary: {
                        issueCount: 1,
                        scannedFormulaCells: 3,
                        errorCount: 1,
                        cycleCount: 0,
                        unsupportedCount: 0,
                      },
                      issues: [
                        {
                          sheetName: 'Sheet1',
                          address: 'C1',
                          formula: '=1/0',
                          valueText: '#DIV/0!',
                          issueKinds: ['error'],
                        },
                      ],
                    }),
                    success: true,
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          ),
      ),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel-scroll-viewport']")).not.toBeNull()
    expect(host.textContent).toContain('Search Workbook')
    expect(host.textContent).toContain('Find Formula Issues')
    expect(host.textContent).not.toContain('Gross Margin')
    expect(host.textContent).not.toContain('gross margin')
    expect(host.textContent).not.toContain('C1')

    const searchToggle = host.querySelector("[data-testid='workbook-agent-tool-toggle-tool-search']")
    const issuesToggle = host.querySelector("[data-testid='workbook-agent-tool-toggle-tool-issues']")
    expect(searchToggle instanceof HTMLButtonElement).toBe(true)
    expect(issuesToggle instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(searchToggle instanceof HTMLButtonElement) || !(issuesToggle instanceof HTMLButtonElement)) {
        throw new Error('Tool toggles not found')
      }
      searchToggle.click()
      issuesToggle.click()
    })

    expect(host.textContent).toContain('Gross Margin')
    expect(host.textContent).toContain('gross margin')
    expect(host.textContent).toContain('C1')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders workbook inspection tool payloads as structured result cards instead of raw JSON blobs', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: 'tool-tables',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: null,
                    phase: null,
                    toolName: 'list_tables',
                    toolStatus: 'completed',
                    argumentsText: '{}',
                    outputText: JSON.stringify({
                      documentId: 'bilig-demo',
                      tableCount: 1,
                      tables: [
                        {
                          name: 'OperatingPlan',
                          sheetName: 'sheet3',
                          startAddress: 'A6',
                          endAddress: 'K10',
                          headerRowCount: 1,
                          rowCount: 4,
                          columnCount: 11,
                          columnNames: ['Item', 'Vendor', 'Category'],
                        },
                      ],
                    }),
                    success: true,
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          ),
      ),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.textContent).toContain('List Tables')
    expect(host.textContent).not.toContain('"documentId": "bilig-demo"')
    expect(host.textContent).not.toContain('"tableCount": 1')

    const readToggle = host.querySelector("[data-testid='workbook-agent-tool-toggle-tool-tables']")
    expect(readToggle instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(readToggle instanceof HTMLButtonElement)) {
        throw new Error('List tables tool toggle not found')
      }
      readToggle.click()
    })

    const readPanelViewport = host.querySelector("[data-testid='workbook-agent-tool-panel-tool-tables-viewport']")
    expect(readPanelViewport instanceof HTMLDivElement).toBe(true)
    expect(readPanelViewport?.className).toContain('h-44')
    expect(host.textContent).toContain('1 table')
    expect(host.textContent).toContain('OperatingPlan')
    expect(host.textContent).toContain('sheet3!A6:K10')
    expect(host.textContent).toContain('4 rows')
    expect(host.textContent).toContain('11 columns')
    expect(host.textContent).not.toContain('"tables": [')

    await act(async () => {
      root.unmount()
    })
  })

  it('summarizes attached selection ranges in tool rows', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: 'tool-context',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: null,
                    phase: null,
                    toolName: 'get_context',
                    toolStatus: 'completed',
                    argumentsText: '{}',
                    outputText: JSON.stringify({
                      selection: {
                        sheetName: 'Sheet1',
                        address: 'E20',
                        range: {
                          startAddress: 'C11',
                          endAddress: 'F20',
                        },
                      },
                      visibleRange: {
                        sheetName: 'Sheet1',
                        startAddress: 'A1',
                        endAddress: 'J20',
                      },
                    }),
                    success: true,
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          ),
      ),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.textContent).toContain('Get Context')
    expect(host.textContent).toContain('Sheet1!C11:F20')
    expect(host.textContent).not.toContain('Sheet1!E20')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not poll assistant thread APIs when the runtime reports the assistant service disabled', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async () => new Response(JSON.stringify([]), { status: 200 }))
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness apiEnabled={false} />)
    })

    expect(fetchSpy).not.toHaveBeenCalled()
    expect(host.textContent).not.toContain('No messages yet')

    await act(async () => {
      root.unmount()
    })
  })

  it('hides raw app-server protocol errors behind user-facing copy', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(
            JSON.stringify({
              error: 'WORKBOOK_AGENT_RUNTIME_UNAVAILABLE',
              message: 'thread/start.dynamicTools requires experimentalApi capability',
              retryable: true,
            }),
            {
              status: 503,
              headers: { 'content-type': 'application/json' },
            },
          ),
      ),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const input = host.querySelector("[data-testid='workbook-agent-input']")
    expect(input instanceof HTMLTextAreaElement).toBe(true)

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error('Agent input not found')
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, 'set') : null
      if (typeof valueSetter !== 'function') {
        throw new Error('Textarea value setter not found')
      }
      Reflect.apply(valueSetter, input, ['Summarize this sheet'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })

    const submit = host.querySelector("[data-testid='workbook-agent-send']")
    await act(async () => {
      if (!(submit instanceof HTMLButtonElement)) {
        throw new Error('Send button not found')
      }
      submit.click()
    })
    await flushToasts()

    expect(host.textContent).toContain('Retry in a moment.')
    expect(host.textContent).not.toContain('thread/start.dynamicTools requires experimentalApi capability')

    await act(async () => {
      root.unmount()
    })
  })

  it('bootstraps the assistant session and streams assistant deltas into the rail', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(JSON.stringify(createSnapshot()), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain('Thinking')
    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'entryTextDelta',
        itemId: 'assistant-1',
        turnId: 'turn-1',
        entryKind: 'assistant',
        delta: 'Updated Sheet1',
      })
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('Updated Sheet1')

    await act(async () => {
      root.unmount()
    })
  })

  it('surfaces malformed assistant stream payloads with stable copy', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(JSON.stringify(createSnapshot()), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        return new Response(JSON.stringify([]), { status: 200, headers: { 'content-type': 'application/json' } })
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    await act(async () => {
      MockEventSource.latest?.emitRaw('{')
    })
    await flushToasts()

    expect(host.textContent).toContain('Assistant stream returned malformed event data.')
    expect(host.textContent).not.toContain('SyntaxError')

    await act(async () => {
      root.unmount()
    })
  })

  it('streams command execution output deltas into command tool rows', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: 'cmd-1',
                    kind: 'system',
                    turnId: 'turn-1',
                    text: 'Codex emitted commandExecution.',
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          )
        }
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain('Codex emitted commandExecution.')
    expect(host.querySelector("[data-testid='workbook-agent-empty-state']")).toBeNull()

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'entryToolOutputDelta',
        itemId: 'cmd-1',
        turnId: 'turn-1',
        delta: 'hi\n',
      })
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('Command')
    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain('Codex emitted commandExecution.')

    const toggle = host.querySelector("[data-testid='workbook-agent-tool-toggle-cmd-1']")
    expect(toggle instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(toggle instanceof HTMLButtonElement)) {
        throw new Error('Command execution toggle not found')
      }
      toggle.click()
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('hi')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders reasoning text immediately from streamed deltas without waiting for a snapshot refresh', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(
            JSON.stringify(
              createSnapshot({
                status: 'inProgress',
                activeTurnId: 'turn-1',
                entries: [
                  {
                    id: 'optimistic-user:turn-1',
                    kind: 'user',
                    turnId: 'turn-1',
                    text: 'Check version issues',
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                    citations: [],
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          )
        }
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain('Thought')

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'entryTextDelta',
        itemId: 'reasoning-1',
        turnId: 'turn-1',
        entryKind: 'reasoning',
        delta: 'Examining version issues',
      })
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('Thought')
    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('Examining version issues')

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'entryTextDelta',
        itemId: 'reasoning-1',
        turnId: 'turn-1',
        entryKind: 'reasoning',
        delta: ' before deciding whether staged changes must be cleared.',
      })
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain(
      'Examining version issues before deciding whether staged changes must be cleared.',
    )

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps the thinking row visible while tool activity is still streaming', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(
            JSON.stringify(
              createSnapshot({
                status: 'inProgress',
                activeTurnId: 'turn-1',
                entries: [
                  {
                    id: 'optimistic-user:turn-1',
                    kind: 'user',
                    turnId: 'turn-1',
                    text: 'Build the operating plan',
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                    citations: [],
                  },
                  {
                    id: 'tool-1',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: '',
                    phase: null,
                    toolName: 'bilig_read_workbook',
                    toolStatus: 'completed',
                    argumentsText: null,
                    outputText: null,
                    success: true,
                    citations: [],
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          )
        }
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.textContent).toContain('Read Workbook')
    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).not.toBeNull()
    expect(host.textContent).toContain('Thinking')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not refetch thread summaries when stream snapshots arrive', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify([createThreadSummary({ threadId: 'thr-1' })]), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads') && requestMethod(init) === 'GET'),
    ).toHaveLength(1)

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'snapshot',
        snapshot: createSnapshot({
          threadId: 'thr-1',
          status: 'inProgress',
          activeTurnId: 'turn-2',
        }),
      })
    })

    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads') && requestMethod(init) === 'GET'),
    ).toHaveLength(1)

    await act(async () => {
      root.unmount()
    })
  })

  it('does not restart live thread bootstrap when callback props churn across internal rerenders', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const zero = createMockZeroAgentHarness({
      initialThreadSummaries: [],
      initialWorkflowRuns: [],
    })
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<UnstableLiveThreadSummaryHarness zero={zero.zero} />)
      await Promise.resolve()
      await Promise.resolve()
    })

    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET'),
    ).toHaveLength(1)
    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads') && requestMethod(init) === 'GET'),
    ).toHaveLength(0)
    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')

    await act(async () => {
      root.unmount()
    })
  })

  it('syncs the latest workbook context after a sheet change even when the context getter reads laggy refs', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(<LaggyContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      await act(async () => {
        host.querySelector("[data-testid='switch-context']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      const contextCalls = fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )
      expect(contextCalls.length).toBeGreaterThan(0)
      expect(requestBody(contextCalls.at(-1)?.[1])).toMatchObject({
        context: {
          selection: {
            sheetName: 'sheet3',
            address: 'A1',
          },
        },
      })
    } finally {
      // no-op
    }

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps workbook context sync single-flight when selection changes faster than the backend responds', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const contextResponses: Array<() => void> = []
    const fetchSpy = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return Promise.resolve(
          new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          }),
        )
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Promise<Response>((resolve) => {
          contextResponses.push(() => {
            resolve(
              new Response(JSON.stringify({ ok: true }), {
                status: 200,
                headers: { 'content-type': 'application/json' },
              }),
            )
          })
        })
      }
      return Promise.reject(new Error(`Unexpected fetch to ${url}`))
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    const advanceSelection = async () => {
      await act(async () => {
        host.querySelector("[data-testid='advance-selection-context']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      })
      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })
    }

    try {
      await act(async () => {
        root.render(<RapidSelectionContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(1)

      await advanceSelection()
      await advanceSelection()
      await advanceSelection()

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        contextResponses[0]?.()
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(2)
      expect(requestBody(contextCalls()[1]?.[1])).toMatchObject({
        context: {
          selection: {
            sheetName: 'Sheet1',
            address: 'A4',
          },
        },
      })

      await act(async () => {
        contextResponses[1]?.()
        await Promise.resolve()
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('retries workbook context sync after a failed server response without marking stale context as synced', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    let contextAttempts = 0
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        contextAttempts += 1
        if (contextAttempts === 1) {
          return new Response(JSON.stringify({ message: 'temporary context failure' }), {
            status: 503,
            headers: { 'content-type': 'application/json' },
          })
        }
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<RapidSelectionContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 900))
      })

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 1_300))
      })

      expect(contextCalls()).toHaveLength(2)
      expect(requestBody(contextCalls()[1]?.[1])).toMatchObject({
        context: {
          selection: {
            sheetName: 'Sheet1',
            address: 'A1',
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not retry a failed in-flight context sync after the assistant is disabled', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    let failContextSync: ((error: Error) => void) | null = null
    const fetchSpy = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return Promise.resolve(
          new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          }),
        )
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Promise<Response>((_resolve, reject) => {
          failContextSync = reject
        })
      }
      return Promise.reject(new Error(`Unexpected fetch to ${url}`))
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<ToggleableContextSyncHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        host.querySelector("[data-testid='disable-agent-context']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        await Promise.resolve()
      })

      await act(async () => {
        failContextSync?.(new Error('context transport down'))
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 900))
      })

      expect(contextCalls()).toHaveLength(1)
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not resync workbook context only because the rendered capture timestamp changes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<VolatileRenderedContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(1)

      const forceInertRender = async () => {
        await act(async () => {
          host.querySelector("[data-testid='force-render']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 200))
        })
      }

      await forceInertRender()
      await forceInertRender()
      await forceInertRender()

      expect(contextCalls()).toHaveLength(1)
      expect(requestBody(contextCalls()[0]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 7,
            batchId: 11,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not rebuild workbook context on unrelated assistant pane renders when a context version is provided', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    const buildContext = vi.fn()
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<VersionedContextRenderHarness onBuildContext={buildContext} />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(buildContext).toHaveBeenCalledTimes(1)
      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        const button = host.querySelector("[data-testid='force-versioned-context-render']")
        for (let index = 0; index < 20; index += 1) {
          button?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        }
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(buildContext).toHaveBeenCalledTimes(1)
      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        host.querySelector("[data-testid='advance-versioned-context']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(buildContext).toHaveBeenCalledTimes(2)
      expect(buildContext).toHaveBeenLastCalledWith('A2')
      expect(contextCalls()).toHaveLength(2)
      expect(requestBody(contextCalls()[1]?.[1])).toMatchObject({
        context: {
          selection: {
            address: 'A2',
            sheetName: 'Sheet1',
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not resync workbook context only because the rendered batch id changes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<VolatileRenderedBatchContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceBatch = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-render-batch']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 200))
        })
      }

      await advanceBatch()
      await advanceBatch()
      await advanceBatch()

      expect(contextCalls()).toHaveLength(1)
      expect(requestBody(contextCalls()[0]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 7,
            batchId: 11,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not resync workbook context only because rendered proof revisions change', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<RapidRenderedRevisionContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceRevision = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-render-revision']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 100))
        })
      }
      await advanceRevision()
      await advanceRevision()
      await advanceRevision()
      await advanceRevision()
      await advanceRevision()

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 900))
      })

      expect(contextCalls()).toHaveLength(1)
      expect(requestBody(contextCalls()[0]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 7,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not poll authoritative revisions only because an assistant turn is in progress', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.useFakeTimers()
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const syncAuthoritativeRevision = vi.fn(async () => undefined)
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-1',
              status: 'inProgress',
              activeTurnId: 'turn-1',
              executionRecords: [],
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(<AgentHarness syncAuthoritativeRevision={syncAuthoritativeRevision} />)
      })
      await act(async () => {
        await Promise.resolve()
        await Promise.resolve()
      })

      expect(syncAuthoritativeRevision).not.toHaveBeenCalled()

      await act(async () => {
        await vi.advanceTimersByTimeAsync(6_000)
      })

      expect(syncAuthoritativeRevision).not.toHaveBeenCalled()
    } finally {
      await act(async () => {
        root.unmount()
      })
      vi.useRealTimers()
    }
  })

  it('requests the exact authoritative revision from applied assistant execution records', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const syncAuthoritativeRevision = vi.fn(async () => undefined)
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-1',
              executionRecords: [
                {
                  id: 'run-1',
                  appliedRevision: 12,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(<AgentHarness syncAuthoritativeRevision={syncAuthoritativeRevision} />)
      })
      await act(async () => {
        await Promise.resolve()
        await Promise.resolve()
      })

      expect(syncAuthoritativeRevision).toHaveBeenCalledTimes(1)
      expect(syncAuthoritativeRevision).toHaveBeenCalledWith(12)
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not treat rendered range churn as immediate context and flood sync requests', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<RapidRenderedRangeContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceRenderedRange = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-rendered-range']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 200))
        })
      }

      await advanceRenderedRange()
      await advanceRenderedRange()
      await advanceRenderedRange()

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 1_000))
      })

      expect(contextCalls()).toHaveLength(1)
      expect(requestBody(contextCalls()[0]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 20,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('throttles rendered range context churn while the assistant turn is active', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              activeTurnId: 'turn-1',
              status: 'inProgress',
              threadId: 'thr-1',
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<RapidRenderedRangeContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceRenderedRange = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-rendered-range']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 800))
        })
      }

      await advanceRenderedRange()
      await advanceRenderedRange()
      await advanceRenderedRange()

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 3_000))
      })

      expect(contextCalls()).toHaveLength(2)
      expect(requestBody(contextCalls()[1]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 23,
            batchId: 23,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  }, 10_000)

  it('does not resync workbook context only because rendered string intern ids change', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<VolatileRenderedStringIdContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceStringId = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-string-id']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 220))
        })
      }

      await advanceStringId()
      await advanceStringId()
      await advanceStringId()

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 900))
      })

      expect(contextCalls()).toHaveLength(1)
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('recreates the assistant session and reconnects the stream after a stale session error', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )

    let resumeCount = 0
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        resumeCount += 1
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-1',
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')

    await act(async () => {
      MockEventSource.latest?.emitError()
      await Promise.resolve()
      await Promise.resolve()
    })

    const sessionCalls = fetchSpy.mock.calls.filter(
      ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET',
    )
    expect(sessionCalls).toHaveLength(2)
    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')
    expect(window.sessionStorage.getItem(agentStorageKey())).toBe(
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )

    await act(async () => {
      root.unmount()
    })
  })

  it('bootstraps from a stored durable thread id without requiring a stored session id', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-1',
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const bootstrapSessionCall = fetchSpy.mock.calls.find(([input, init]) => {
      return requestUrl(input).endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET'
    })
    expect(bootstrapSessionCall).toBeDefined()
    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')
    expect(window.sessionStorage.getItem(agentStorageKey())).toContain('"threadId":"thr-1"')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not render private review controls for restored private review items', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const preview = createPreviewSummary({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
          role: 'target' as const,
        },
      ],
      cellDiffs: [
        {
          sheetName: 'Sheet1',
          address: 'A1',
          beforeInput: 1,
          beforeFormula: null,
          afterInput: 1,
          afterFormula: null,
          changeKinds: ['style'],
        },
      ],
      effectSummary: {
        displayedCellDiffCount: 1,
        truncatedCellDiffs: false,
        inputChangeCount: 0,
        formulaChangeCount: 0,
        styleChangeCount: 1,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    })
    const previewBundle = vi.fn(async () => preview)
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              reviewBundle: {
                id: 'bundle-1',
                documentId: 'doc-1',
                threadId: 'thr-1',
                turnId: 'turn-1',
                goalText: 'Bold the selected cell',
                summary: 'Format Sheet1!A1',
                scope: 'selection',
                riskClass: 'low',
                baseRevision: 3,
                createdAtUnixMs: 10,
                context: {
                  selection: {
                    sheetName: 'Sheet1',
                    address: 'A1',
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A1',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'A1',
                    endAddress: 'A1',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 1,
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/review-items/bundle-1/apply')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              reviewQueueItems: [],
              executionRecords: [
                {
                  id: 'run-1',
                  bundleId: 'bundle-1',
                  documentId: 'doc-1',
                  threadId: 'thr-1',
                  turnId: 'turn-1',
                  actorUserId: 'user@example.com',
                  goalText: 'Bold the selected cell',
                  planText: 'Apply bold formatting',
                  summary: 'Format Sheet1!A1',
                  scope: 'selection',
                  riskClass: 'low',
                  acceptedScope: 'full',
                  appliedBy: 'auto',
                  baseRevision: 3,
                  appliedRevision: 4,
                  createdAtUnixMs: 10,
                  appliedAtUnixMs: 20,
                  context: {
                    selection: {
                      sheetName: 'Sheet1',
                      address: 'A1',
                    },
                    viewport: {
                      rowStart: 0,
                      rowEnd: 10,
                      colStart: 0,
                      colEnd: 5,
                    },
                  },
                  commands: [
                    {
                      kind: 'formatRange',
                      range: {
                        sheetName: 'Sheet1',
                        startAddress: 'A1',
                        endAddress: 'A1',
                      },
                      patch: {
                        font: {
                          bold: true,
                        },
                      },
                    },
                  ],
                  preview,
                },
              ],
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      return new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness previewCommandBundle={previewBundle} />)
    })

    await act(async () => {
      await Promise.resolve()
    })

    expect(previewBundle).not.toHaveBeenCalled()
    const applyCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-1/apply'))
    expect(applyCall).toBeUndefined()
    expect(host.textContent).not.toContain('Apply')
    expect(host.textContent).not.toContain('Executions')
    expect(host.textContent).not.toContain('Replay')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not auto-apply low-risk review items on shared threads', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-shared',
      }),
    )
    const preview = createPreviewSummary({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
          role: 'target' as const,
        },
      ],
    })
    const previewBundle = vi.fn(async () => preview)
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-shared') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-1',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-1',
                goalText: 'Bold the selected cell',
                summary: 'Format Sheet1!A1',
                scope: 'selection',
                riskClass: 'low',
                baseRevision: 3,
                createdAtUnixMs: 10,
                context: {
                  selection: {
                    sheetName: 'Sheet1',
                    address: 'A1',
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A1',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'A1',
                    endAddress: 'A1',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 1,
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness previewCommandBundle={previewBundle} />)
    })

    await act(async () => {
      await Promise.resolve()
    })

    expect(previewBundle).toHaveBeenCalled()
    expect(previewBundle.mock.calls[0]?.[0]).toMatchObject({
      id: 'bundle-shared-1',
    })
    const applyCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-shared-1/apply'))
    expect(applyCall).toBeUndefined()

    await act(async () => {
      root.unmount()
    })
  })

  it('blocks collaborator approval of shared medium-risk bundles in the panel', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey('casey@example.com'),
      JSON.stringify({
        threadId: 'thr-shared',
      }),
    )
    const preview = createPreviewSummary({
      structuralChanges: ['Format selected range'],
    })
    const previewBundle = vi.fn(async () => preview)
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && (init?.method ?? 'GET') === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-shared',
              scope: 'shared',
              ownerUserId: 'alex@example.com',
              entryCount: 3,
              reviewQueueItemCount: 1,
              latestEntryText: 'Review item queued',
            }),
          ]),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/chat/threads/thr-shared') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-2',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: {
                  selection: {
                    sheetName: 'Sheet1',
                    address: 'A1',
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'A1',
                    endAddress: 'A20',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness currentUserId="casey@example.com" previewCommandBundle={previewBundle} />)
    })

    await act(async () => {
      await Promise.resolve()
    })

    const applyButton = host.querySelector("[data-testid='workbook-agent-apply-review-item']")
    if (!(applyButton instanceof HTMLButtonElement)) {
      throw new Error('Expected apply button to render')
    }
    expect(applyButton.disabled).toBe(true)
    expect(host.textContent).toContain('Owner review routes medium/high-risk changes to Alex on this shared thread.')
    expect(host.textContent).toContain('Owner review is in progress with Alex.')

    await act(async () => {
      root.unmount()
    })
  })

  it('lets the shared thread owner approve a medium-risk bundle before apply', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-shared',
      }),
    )
    const preview = createPreviewSummary({
      structuralChanges: ['Normalize selected range'],
    })
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && (init?.method ?? 'GET') === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-shared',
              scope: 'shared',
              ownerUserId: 'alex@example.com',
              entryCount: 3,
              reviewQueueItemCount: 1,
              latestEntryText: 'Review item queued',
            }),
          ]),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/chat/threads/thr-shared') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-owner',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: {
                  selection: {
                    sheetName: 'Sheet1',
                    address: 'A1',
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'A1',
                    endAddress: 'A20',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/review-items/bundle-shared-owner/review')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-owner',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: null,
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'approved',
                  decidedByUserId: 'alex@example.com',
                  decidedAtUnixMs: 25,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness currentUserId="alex@example.com" previewCommandBundle={vi.fn(async () => preview)} />)
    })

    await act(async () => {
      await Promise.resolve()
    })

    const applyButton = host.querySelector("[data-testid='workbook-agent-apply-review-item']")
    const approveButton = host.querySelector("[data-testid='workbook-agent-review-item-approve']")
    if (!(applyButton instanceof HTMLButtonElement)) {
      throw new Error('Expected apply button')
    }
    if (!(approveButton instanceof HTMLButtonElement)) {
      throw new Error('Expected approve button')
    }
    expect(applyButton.disabled).toBe(true)

    await act(async () => {
      approveButton.click()
    })

    const reviewCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-shared-owner/review'))
    expect(requestBody(reviewCall?.[1])).toEqual({
      decision: 'approved',
    })
    expect(host.textContent).toContain('Approved by Alex.')
    const refreshedApplyButton = host.querySelector("[data-testid='workbook-agent-apply-review-item']")
    if (!(refreshedApplyButton instanceof HTMLButtonElement)) {
      throw new Error('Expected refreshed apply button')
    }
    expect(refreshedApplyButton.disabled).toBe(false)

    await act(async () => {
      root.unmount()
    })
  })

  it('lets collaborators recommend approval on shared medium-risk bundles', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey('casey@example.com'),
      JSON.stringify({
        threadId: 'thr-shared',
      }),
    )
    const preview = createPreviewSummary({
      structuralChanges: ['Normalize selected range'],
    })
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && (init?.method ?? 'GET') === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-shared',
              scope: 'shared',
              ownerUserId: 'alex@example.com',
              entryCount: 3,
              reviewQueueItemCount: 1,
              latestEntryText: 'Review item queued',
            }),
          ]),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/chat/threads/thr-shared') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-collab',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: null,
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/review-items/bundle-shared-collab/review')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-collab',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: null,
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [
                    {
                      userId: 'casey@example.com',
                      decision: 'approved',
                      decidedAtUnixMs: 30,
                    },
                  ],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness currentUserId="casey@example.com" previewCommandBundle={async () => preview} />)
    })

    const approveButton = host.querySelector("[data-testid='workbook-agent-review-item-approve']")
    const clearButton = host.querySelector("[data-testid='workbook-agent-dismiss-review-item']")
    expect(approveButton instanceof HTMLButtonElement).toBe(true)
    expect(clearButton instanceof HTMLButtonElement).toBe(true)
    expect(clearButton instanceof HTMLButtonElement ? clearButton.disabled : false).toBe(true)
    expect(host.textContent).toContain('Owner review is in progress with Alex.')

    await act(async () => {
      if (!(approveButton instanceof HTMLButtonElement)) {
        throw new Error('Expected recommend approve button')
      }
      approveButton.click()
    })

    const reviewCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-shared-collab/review'))
    expect(requestBody(reviewCall?.[1])).toEqual({
      decision: 'approved',
    })
    expect(host.textContent).toContain('1 approval recommendation')
    expect(host.textContent).toContain('You recommended approval.')

    await act(async () => {
      root.unmount()
    })
  })

  it('re-previews and applies only the selected command subset', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fullPreview = createPreviewSummary({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C3',
          role: 'target' as const,
        },
      ],
      effectSummary: {
        displayedCellDiffCount: 2,
        truncatedCellDiffs: false,
        inputChangeCount: 2,
        formulaChangeCount: 0,
        styleChangeCount: 0,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    })
    const subsetPreview = createPreviewSummary({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          endAddress: 'C3',
          role: 'target' as const,
        },
      ],
      cellDiffs: [
        {
          sheetName: 'Sheet1',
          address: 'C3',
          beforeInput: null,
          beforeFormula: null,
          afterInput: 2,
          afterFormula: null,
          changeKinds: ['input'],
        },
      ],
      effectSummary: {
        displayedCellDiffCount: 1,
        truncatedCellDiffs: false,
        inputChangeCount: 1,
        formulaChangeCount: 0,
        styleChangeCount: 0,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    })
    const previewBundle = vi.fn(async (bundle) => (bundle.commands.length === 1 ? subsetPreview : fullPreview))
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-1',
                documentId: 'doc-1',
                threadId: 'thr-1',
                turnId: 'turn-1',
                goalText: 'Update two cells',
                summary: 'Write cells in Sheet1!B2 and 1 more change',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 3,
                createdAtUnixMs: 10,
                context: {
                  selection: {
                    sheetName: 'Sheet1',
                    address: 'A1',
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: 'writeRange',
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    values: [[1]],
                  },
                  {
                    kind: 'writeRange',
                    sheetName: 'Sheet1',
                    startAddress: 'C3',
                    values: [[2]],
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    endAddress: 'B2',
                    role: 'target',
                  },
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'C3',
                    endAddress: 'C3',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 2,
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/review-items/bundle-1/apply')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-2',
                documentId: 'doc-1',
                threadId: 'thr-1',
                turnId: 'turn-1',
                goalText: 'Update two cells',
                summary: 'Write cells in Sheet1!B2',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: {
                  selection: {
                    sheetName: 'Sheet1',
                    address: 'A1',
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: 'writeRange',
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    values: [[1]],
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    endAddress: 'B2',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 1,
              },
              executionRecords: [
                {
                  id: 'run-1',
                  bundleId: 'bundle-1',
                  documentId: 'doc-1',
                  threadId: 'thr-1',
                  turnId: 'turn-1',
                  actorUserId: 'user@example.com',
                  goalText: 'Update two cells',
                  planText: 'Apply only the second cell',
                  summary: 'Write cells in Sheet1!C3',
                  scope: 'sheet',
                  riskClass: 'medium',
                  acceptedScope: 'partial',
                  appliedBy: 'user',
                  baseRevision: 3,
                  appliedRevision: 4,
                  createdAtUnixMs: 10,
                  appliedAtUnixMs: 20,
                  context: {
                    selection: {
                      sheetName: 'Sheet1',
                      address: 'A1',
                    },
                    viewport: {
                      rowStart: 0,
                      rowEnd: 10,
                      colStart: 0,
                      colEnd: 5,
                    },
                  },
                  commands: [
                    {
                      kind: 'writeRange',
                      sheetName: 'Sheet1',
                      startAddress: 'C3',
                      values: [[2]],
                    },
                  ],
                  preview: subsetPreview,
                },
              ],
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      return new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness previewCommandBundle={previewBundle} />)
    })

    expect(previewBundle).toHaveBeenCalledTimes(1)
    expect(host.textContent).toContain('2/2')

    const firstToggle = host.querySelector("[data-testid='workbook-agent-review-command-toggle-0']")
    expect(firstToggle instanceof HTMLInputElement).toBe(true)

    await act(async () => {
      firstToggle?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    await act(async () => {
      await Promise.resolve()
    })

    expect(previewBundle).toHaveBeenCalledTimes(2)
    expect(previewBundle.mock.calls[1]?.[0]).toEqual(
      expect.objectContaining({
        commands: [
          {
            kind: 'writeRange',
            sheetName: 'Sheet1',
            startAddress: 'C3',
            values: [[2]],
          },
        ],
      }),
    )
    expect(host.textContent).toContain('1/2')

    const applyButton = host.querySelector("[data-testid='workbook-agent-apply-review-item']")
    expect(applyButton).toBeTruthy()

    await act(async () => {
      applyButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const applyCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-1/apply'))
    expect(applyCall?.[1]?.body).toBe(
      JSON.stringify({
        appliedBy: 'user',
        commandIndexes: [1],
      }),
    )
    expect(host.textContent).toContain('Write cells in Sheet1!B2')
    expect(host.textContent).toContain('Sheet1!C3')
    expect(host.textContent).not.toContain('Recent changes')
    expect(host.textContent).not.toContain('Run again')

    await act(async () => {
      root.unmount()
    })
  })
})
