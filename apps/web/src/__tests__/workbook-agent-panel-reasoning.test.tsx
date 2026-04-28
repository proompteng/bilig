// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { WorkbookAgentPanel } from '../WorkbookAgentPanel.js'

beforeEach(() => {
  ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(() => {
  document.body.innerHTML = ''
})

function renderPanel(entry: {
  id: string
  kind: 'reasoning' | 'plan' | 'system'
  text: string | null
  citations?: Array<
    | { kind: 'revision'; revision: number }
    | {
        kind: 'range'
        sheetName: string
        startAddress: string
        endAddress: string
        role: 'source' | 'target'
      }
  >
  executionRecords?: Array<Record<string, unknown>>
}) {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)

  return {
    host,
    root,
    render: async () => {
      await act(async () => {
        root.render(
          <WorkbookAgentPanel
            activeThreadId="thr-1"
            activeContextLabel="Sheet1!A1"
            canFinalizeSharedBundle={false}
            canRecommendSharedBundle={false}
            cancellingWorkflowRunId={null}
            currentUserSharedRecommendation={null}
            draft=""
            executionRecords={entry.executionRecords ?? []}
            isApplyingReviewItem={false}
            isLoading={false}
            optimisticEntries={[]}
            activeResponseTurnId={null}
            onApplyReviewItem={() => {}}
            onCancelWorkflowRun={() => {}}
            onDismissReviewItem={() => {}}
            onDraftChange={() => {}}
            onInterrupt={() => {}}
            onReplayExecutionRecord={() => {}}
            onReviewReviewItem={() => {}}
            onSelectAllReviewCommands={() => {}}
            onSelectThread={() => {}}
            onSubmit={() => {}}
            onToggleReviewCommand={() => {}}
            activeReviewBundle={null}
            preview={null}
            selectedCommandIndexes={[]}
            showAssistantProgress={false}
            sharedApprovalOwnerUserId={null}
            sharedReviewDecidedByUserId={null}
            sharedReviewOwnerUserId={null}
            sharedReviewRecommendations={[]}
            sharedReviewStatus={null}
            snapshot={{
              sessionId: 'agent-session-1',
              documentId: 'doc-1',
              threadId: 'thr-1',
              executionPolicy: 'autoApplyAll',
              scope: 'private',
              status: 'idle',
              activeTurnId: null,
              lastError: null,
              context: {
                selection: { sheetName: 'Sheet1', address: 'A1' },
                viewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 5 },
              },
              entries: [
                {
                  id: entry.id,
                  kind: entry.kind,
                  turnId: 'turn-1',
                  text: entry.text,
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                  citations: entry.citations ?? [],
                },
              ],
              reviewQueueItems: [],
              executionRecords: entry.executionRecords ?? [],
              workflowRuns: [],
            }}
            threadSummaries={[]}
            workflowRuns={[]}
          />,
        )
      })
    },
  }
}

describe('WorkbookAgentPanel reasoning', () => {
  it('renders reasoning entries as collapsible thought rows', async () => {
    const panel = renderPanel({
      id: 'reasoning-1',
      kind: 'reasoning',
      text: '**Check** the visible formulas before applying edits.',
    })

    await panel.render()

    expect(panel.host.textContent).toContain('Thought')
    expect(panel.host.textContent).toContain('Check the visible formulas before applying edits.')
    expect(panel.host.textContent).not.toContain('**Check**')

    const toggle = panel.host.querySelector("[data-testid='workbook-agent-reasoning-toggle-reasoning-1']")
    expect(toggle instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(toggle instanceof HTMLButtonElement)) {
        throw new Error('Reasoning toggle not found')
      }
      toggle.click()
    })

    expect(panel.host.querySelector("[data-testid='workbook-agent-reasoning-panel-reasoning-1']")).not.toBeNull()
    expect(panel.host.textContent).not.toContain('**Check**')

    const expandedBody = panel.host.querySelector("[data-testid='workbook-agent-reasoning-panel-reasoning-1-viewport'] > div")
    expect(expandedBody?.className).toContain('px-2')
    expect(expandedBody?.className).toContain('pb-2')
    expect(expandedBody?.className).toContain('w-full')
    expect(expandedBody?.className).toContain('max-w-full')

    const expandedCard = expandedBody?.firstElementChild
    expect(expandedCard?.className).toContain('bg-[var(--wb-surface-muted)]')
    expect(expandedCard?.className).toContain('rounded-[var(--wb-radius-control)]')
    expect(expandedCard?.className).toContain('w-full')
    expect(expandedCard?.className).toContain('max-w-full')
    expect(expandedCard?.className).toContain('overflow-x-hidden')
    expect(expandedCard?.className).toContain('px-3')
    expect(expandedCard?.className).toContain('py-3')

    const trigger = panel.host.querySelector("[data-testid='workbook-agent-reasoning-toggle-reasoning-1']")
    expect(trigger?.parentElement?.className).not.toContain('bg-[var(--wb-surface-muted)]')

    await act(async () => {
      panel.root.unmount()
    })
  })

  it('renders plan entries as plan rows instead of overloading them as reasoning', async () => {
    const panel = renderPanel({
      id: 'plan-1',
      kind: 'plan',
      text: 'Inspect the named ranges and verify the import columns.',
    })

    await panel.render()

    expect(panel.host.textContent).toContain('Plan')
    expect(panel.host.textContent).toContain('Inspect the named ranges and verify the import columns.')

    await act(async () => {
      panel.root.unmount()
    })
  })

  it('renders the composer inside a Base UI scroll viewport', async () => {
    const panel = renderPanel({
      id: 'system-input-1',
      kind: 'system',
      text: 'Ready',
    })

    await panel.render()

    const input = panel.host.querySelector("[data-testid='workbook-agent-input']")
    const viewport = panel.host.querySelector("[data-testid='workbook-agent-input-viewport']")

    expect(input instanceof HTMLTextAreaElement).toBe(true)
    expect(viewport instanceof HTMLDivElement).toBe(true)
    expect(input?.className).toContain('overflow-hidden')

    await act(async () => {
      panel.root.unmount()
    })
  })

  it('renders citations as inline metadata instead of pills', async () => {
    const panel = renderPanel({
      id: 'system-apply-1',
      kind: 'system',
      text: 'Prepared workbook review item: Write cells in Sheet1!B2',
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
    })

    await panel.render()

    expect(panel.host.textContent).toContain('Prepared workbook review item: Write cells in Sheet1!B2')
    expect(panel.host.textContent).toContain('Target Sheet1!B2')
    expect(panel.host.textContent).not.toContain('Target Sheet1!B2 · r7')

    await act(async () => {
      panel.root.unmount()
    })
  })

  it('does not render execution history cards above the composer', async () => {
    const panel = renderPanel({
      id: 'system-1',
      kind: 'system',
      text: 'Applied workbook change set at revision r7: Write cells in Sheet1!B2',
      executionRecords: [
        {
          id: 'run-1',
          bundleId: 'bundle-1',
          documentId: 'doc-1',
          threadId: 'thr-1',
          turnId: 'turn-1',
          actorUserId: 'alex@example.com',
          goalText: 'Write 42',
          planText: null,
          summary: 'Write cells in Sheet1!B2',
          scope: 'selection',
          riskClass: 'low',
          acceptedScope: 'full',
          appliedBy: 'auto',
          baseRevision: 6,
          appliedRevision: 7,
          createdAtUnixMs: 100,
          appliedAtUnixMs: 200,
          context: null,
          commands: [
            {
              kind: 'writeRange',
              sheetName: 'Sheet1',
              startAddress: 'B2',
              values: [[42]],
            },
          ],
          preview: null,
        },
      ],
    })

    await panel.render()

    expect(panel.host.textContent).not.toContain('Recent changes')
    expect(panel.host.textContent).not.toContain('Applied automatically')
    expect(panel.host.textContent).not.toContain('Revision r7')
    expect(panel.host.textContent).not.toContain('Run again')

    await act(async () => {
      panel.root.unmount()
    })
  })
})
