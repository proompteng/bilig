// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { useWorkbookAppPanels } from '../use-workbook-app-panels.js'

const { useWorkbookAgentPane } = vi.hoisted(() => ({
  useWorkbookAgentPane: vi.fn(),
}))

vi.mock('../use-workbook-agent-pane.js', () => ({
  useWorkbookAgentPane,
}))

vi.mock('../use-workbook-presence.js', () => ({
  useWorkbookPresence: () => [],
}))

function renderHarness(host: HTMLElement) {
  function Harness() {
    const panels = useWorkbookAppPanels({
      documentId: 'doc-1',
      currentUserId: 'alex@example.com',
      presenceClientId: 'presence:self',
      replicaId: 'replica-1',
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      sheetNames: ['Sheet1'],
      zero: undefined,
      runtimeReady: true,
      zeroConfigured: true,
      remoteSyncAvailable: true,
      changeCount: 1,
      changesPanel: <div data-testid="changes-panel">Changes panel</div>,
      selectAddress: vi.fn(),
      getAgentContext: () => ({
        selection: { sheetName: 'Sheet1', address: 'A1' },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 8,
        },
      }),
      previewAgentCommandBundle: vi.fn(),
    })

    return (
      <>
        <div data-testid="toolbar-trailing-content">{panels.toolbarTrailingContent}</div>
        <div data-testid="side-rail">{panels.sidePanel}</div>
      </>
    )
  }

  const root = createRoot(host)
  return {
    root,
    async render() {
      await act(async () => {
        root.render(<Harness />)
      })
    },
    async unmount() {
      await act(async () => {
        root.unmount()
      })
    },
  }
}

function createThreadSummary(overrides: Record<string, unknown> = {}) {
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

function mockAgentPane(pendingCommandCount: number) {
  useWorkbookAgentPane.mockReturnValue({
    agentPanel: <div data-testid="assistant-panel">Assistant panel</div>,
    activeThreadId: 'thr-1',
    agentError: null,
    clearAgentError: vi.fn(),
    pendingCommandCount,
    previewRanges: [],
    selectThread: vi.fn(),
    startNewThread: vi.fn(),
    threadSummaries: [],
  })
}

describe('useWorkbookAppPanels', () => {
  beforeEach(() => {
    const backingStore = new Map<string, string>()
    const storage = {
      clear() {
        backingStore.clear()
      },
      getItem(key: string) {
        return backingStore.get(key) ?? null
      },
      key(index: number) {
        return [...backingStore.keys()][index] ?? null
      },
      removeItem(key: string) {
        backingStore.delete(key)
      },
      setItem(key: string, value: string) {
        backingStore.set(key, value)
      },
      get length() {
        return backingStore.size
      },
    }
    Object.defineProperty(window, 'localStorage', {
      configurable: true,
      value: storage,
    })
  })

  afterEach(() => {
    vi.clearAllMocks()
    window.localStorage.clear()
    document.body.innerHTML = ''
  })

  it('opens the assistant panel by default on first render', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    mockAgentPane(0)
    const host = document.createElement('div')
    document.body.appendChild(host)
    const harness = renderHarness(host)

    await harness.render()
    expect(host.querySelector("[data-testid='workbook-side-panel-toggle-group']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-side-panel-panel-assistant']")).not.toBeNull()

    mockAgentPane(2)
    await harness.render()

    expect(host.querySelector("[data-testid='workbook-side-panel-panel-assistant']")).not.toBeNull()
    expect(host.textContent).toContain('Assistant panel')

    await harness.unmount()
  })

  it('switches from the changes panel to the assistant panel when a review item is staged', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    mockAgentPane(0)
    const host = document.createElement('div')
    document.body.appendChild(host)
    const harness = renderHarness(host)

    await harness.render()

    await act(async () => {
      host.querySelector("[data-testid='workbook-side-panel-tab-changes']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(host.querySelector("[data-testid='workbook-side-panel-panel-changes']")).not.toBeNull()
    expect(host.textContent).toContain('Changes panel')

    mockAgentPane(1)
    await harness.render()

    expect(host.querySelector("[data-testid='workbook-side-panel-panel-assistant']")).not.toBeNull()
    expect(host.textContent).toContain('Assistant panel')

    await harness.unmount()
  })

  it('renders the new thread action in the rail tab row', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const startNewThread = vi.fn()
    useWorkbookAgentPane.mockReturnValue({
      agentPanel: <div data-testid="assistant-panel">Assistant panel</div>,
      activeThreadId: 'thr-1',
      agentError: null,
      clearAgentError: vi.fn(),
      pendingCommandCount: 0,
      previewRanges: [],
      selectThread: vi.fn(),
      startNewThread,
      threadSummaries: [],
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const harness = renderHarness(host)

    await harness.render()

    const newThreadButton = host.querySelector<HTMLButtonElement>("[data-testid='workbook-agent-new-thread']")
    expect(newThreadButton).not.toBeNull()
    expect(newThreadButton?.getAttribute('aria-label')).toBe('New thread')
    expect(newThreadButton?.textContent?.trim()).toBe('')
    expect(newThreadButton?.className).toContain('bg-[var(--color-mauve-50)]')
    expect(newThreadButton?.className).toContain('border-transparent')

    await act(async () => {
      newThreadButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(startNewThread).toHaveBeenCalledTimes(1)

    await harness.unmount()
  })

  it('renders a previous conversations menu in the assistant rail header', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectThread = vi.fn()
    useWorkbookAgentPane.mockReturnValue({
      agentPanel: <div data-testid="assistant-panel">Assistant panel</div>,
      activeThreadId: 'thr-1',
      agentError: null,
      clearAgentError: vi.fn(),
      pendingCommandCount: 0,
      previewRanges: [],
      selectThread,
      startNewThread: vi.fn(),
      threadSummaries: [
        createThreadSummary(),
        createThreadSummary({
          threadId: 'thr-2',
          scope: 'shared',
          ownerUserId: 'sam@example.com',
          updatedAtUnixMs: 200,
          entryCount: 4,
          reviewQueueItemCount: 1,
          latestEntryText: 'Applied workbook change set at revision r7',
        }),
      ],
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const harness = renderHarness(host)

    await harness.render()

    const historyButton = host.querySelector<HTMLButtonElement>("[data-testid='workbook-agent-history-trigger']")
    expect(historyButton).not.toBeNull()
    expect(historyButton?.getAttribute('aria-label')).toBe('Previous conversations')

    await act(async () => {
      historyButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const threadButton = document.querySelector<HTMLButtonElement>("[data-testid='workbook-agent-history-thread-thr-2']")
    expect(threadButton).not.toBeNull()
    expect(threadButton?.textContent).toContain('Sam')
    expect(threadButton?.textContent).toContain('4 items')

    await act(async () => {
      threadButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(selectThread).toHaveBeenCalledWith('thr-2')

    await harness.unmount()
  })
})
