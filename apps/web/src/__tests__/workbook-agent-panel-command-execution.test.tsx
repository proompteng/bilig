// @vitest-environment jsdom
import type { ComponentProps } from 'react'
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { WorkbookAgentThreadSnapshot, WorkbookAgentTimelineEntry } from '@bilig/contracts'
import { WorkbookAgentPanel } from '../WorkbookAgentPanel.js'

beforeEach(() => {
  ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

function createSnapshot(entries: readonly WorkbookAgentTimelineEntry[]): WorkbookAgentThreadSnapshot {
  return {
    documentId: 'doc-1',
    threadId: 'thr-1',
    scope: 'private',
    executionPolicy: 'autoApplyAll',
    status: 'idle',
    activeTurnId: null,
    lastError: null,
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
    entries: [...entries],
    reviewQueueItems: [],
    executionRecords: [],
    workflowRuns: [],
  }
}

function createPanelProps(snapshot: WorkbookAgentThreadSnapshot): ComponentProps<typeof WorkbookAgentPanel> {
  return {
    activeThreadId: 'thr-1',
    activeContextLabel: 'Sheet1!A1',
    activeResponseTurnId: null,
    optimisticEntries: [],
    snapshot,
    showAssistantProgress: false,
    activeReviewBundle: null,
    preview: null,
    sharedApprovalOwnerUserId: null,
    sharedReviewOwnerUserId: null,
    sharedReviewStatus: null,
    sharedReviewDecidedByUserId: null,
    sharedReviewRecommendations: [],
    currentUserSharedRecommendation: null,
    canFinalizeSharedBundle: false,
    canRecommendSharedBundle: false,
    selectedCommandIndexes: [],
    workflowRuns: [],
    cancellingWorkflowRunId: null,
    threadSummaries: [],
    draft: '',
    isLoading: false,
    isApplyingReviewItem: false,
    onApplyReviewItem: vi.fn(),
    onDraftChange: vi.fn(),
    onDismissReviewItem: vi.fn(),
    onReviewReviewItem: vi.fn(),
    onInterrupt: vi.fn(),
    onSelectAllReviewCommands: vi.fn(),
    onSelectThread: vi.fn(),
    onToggleReviewCommand: vi.fn(),
    onCancelWorkflowRun: vi.fn(),
    onSubmit: vi.fn(),
  }
}

function renderPanel(snapshot: WorkbookAgentThreadSnapshot) {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)

  return {
    host,
    root,
    render: async () => {
      await act(async () => {
        root.render(<WorkbookAgentPanel {...createPanelProps(snapshot)} />)
      })
    },
    unmount: async () => {
      await act(async () => {
        root.unmount()
      })
    },
  }
}

describe('WorkbookAgentPanel command execution rendering', () => {
  it('renders command execution entries as command tool rows with raw terminal output', async () => {
    const commandEntry = {
      id: 'cmd-1',
      kind: 'tool',
      turnId: 'turn-1',
      text: null,
      phase: null,
      toolName: 'command_execution',
      toolStatus: 'completed',
      argumentsText: JSON.stringify(
        {
          command: 'pnpm test',
          cwd: '/Users/gregkonush/github.com/bilig',
          status: 'completed',
          processId: null,
          exitCode: 0,
          durationMs: 42,
        },
        null,
        2,
      ),
      outputText: '{"json": true}\n',
      success: true,
      citations: [],
    } satisfies WorkbookAgentTimelineEntry
    const panel = renderPanel(createSnapshot([commandEntry]))

    await panel.render()

    expect(panel.host.textContent).toContain('Command')
    expect(panel.host.textContent).toContain('$ pnpm test')
    expect(panel.host.textContent).toContain('exit 0')
    expect(panel.host.textContent).not.toContain('Codex emitted commandExecution.')
    expect(panel.host.textContent).not.toContain('{"json": true}')

    const toggle = panel.host.querySelector("[data-testid='workbook-agent-tool-toggle-cmd-1']")
    expect(toggle instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(toggle instanceof HTMLButtonElement)) {
        throw new Error('Command execution toggle not found')
      }
      toggle.click()
    })

    const preTexts = [...panel.host.querySelectorAll('pre')].map((node) => node.textContent)
    expect(preTexts).toEqual(expect.arrayContaining([expect.stringContaining('"command": "pnpm test"')]))
    expect(preTexts).toEqual(expect.arrayContaining([expect.stringContaining('{"json": true}')]))

    await panel.unmount()
  })

  it('hides legacy generic command execution placeholder rows from persisted threads', async () => {
    const legacyEntry = {
      id: 'cmd-legacy-1',
      kind: 'system',
      turnId: 'turn-1',
      text: 'Codex emitted commandExecution.',
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
      citations: [],
    } satisfies WorkbookAgentTimelineEntry
    const panel = renderPanel(createSnapshot([legacyEntry]))

    await panel.render()

    expect(panel.host.textContent).not.toContain('Codex emitted commandExecution.')
    expect(panel.host.querySelector("[data-testid='workbook-agent-empty-state']")).toBeNull()

    await panel.unmount()
  })
})
