import { describe, expect, it } from 'vitest'
import type { WorkbookAgentCommand } from '@bilig/agent-api'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'
import {
  applyWorkbookAgentStructuralContextHints,
  stripRenderedWorkbookAgentContext,
  updateWorkbookAgentDurableUiContextFromUser,
} from './workbook-agent-service-context.js'

function createContext(overrides: Partial<WorkbookAgentUiContext> = {}): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: 'Revenue',
      address: 'B2',
      range: {
        startAddress: 'B2',
        endAddress: 'C3',
      },
    },
    viewport: {
      rowStart: 0,
      rowEnd: 20,
      colStart: 0,
      colEnd: 10,
    },
    rendered: {
      capturedAtUnixMs: 100,
      capturedRevision: 3,
      batchId: 1,
      selection: null,
      visibleRange: null,
    },
    ...overrides,
  }
}

function createThreadState(context: WorkbookAgentUiContext | null): WorkbookAgentThreadState {
  return {
    documentId: 'doc-1',
    userId: 'alex@example.com',
    storageActorUserId: 'alex@example.com',
    scope: 'private',
    executionPolicy: 'autoApplyAll',
    threadId: 'thr-1',
    durable: {
      context,
      entries: [],
      reviewQueueItems: [],
      executionRecords: [],
      workflowRuns: [],
    },
    live: {
      activeTurnId: 'turn-1',
      status: 'idle',
      lastError: null,
      stagedPrivateBundleByTurn: new Map(),
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      turnActorUserIdByTurn: new Map([['turn-1', 'alex@example.com']]),
      turnContextByTurn: new Map(),
      lastAccessedAt: 0,
    },
  }
}

describe('workbook agent service context helpers', () => {
  it('strips rendered context while preserving selection and viewport', () => {
    expect(stripRenderedWorkbookAgentContext(createContext())).toEqual({
      selection: {
        sheetName: 'Revenue',
        address: 'B2',
        range: {
          startAddress: 'B2',
          endAddress: 'C3',
        },
      },
      viewport: {
        rowStart: 0,
        rowEnd: 20,
        colStart: 0,
        colEnd: 10,
      },
    })
  })

  it('applies structural command hints to workbook context', () => {
    const renamed = applyWorkbookAgentStructuralContextHints(createContext(), [
      {
        kind: 'renameSheet',
        currentName: 'Revenue',
        nextName: 'Forecast',
      } satisfies Extract<WorkbookAgentCommand, { kind: 'renameSheet' }>,
    ])
    expect(renamed?.selection.sheetName).toBe('Forecast')
    expect(renamed).not.toHaveProperty('rendered')

    const created = applyWorkbookAgentStructuralContextHints(createContext(), [
      {
        kind: 'createSheet',
        name: 'Ops',
      } satisfies Extract<WorkbookAgentCommand, { kind: 'createSheet' }>,
    ])
    expect(created).toEqual({
      selection: {
        sheetName: 'Ops',
        address: 'A1',
        range: {
          startAddress: 'A1',
          endAddress: 'A1',
        },
      },
      viewport: {
        rowStart: 0,
        rowEnd: 20,
        colStart: 0,
        colEnd: 10,
      },
    })
  })

  it('updates durable and turn context only for the active turn owner', () => {
    const sessionState = createThreadState(null)
    const nextContext = createContext({ rendered: undefined })

    updateWorkbookAgentDurableUiContextFromUser({
      sessionState,
      context: nextContext,
      userId: 'alex@example.com',
    })

    expect(sessionState.durable.context).toEqual(nextContext)
    expect(sessionState.live.turnContextByTurn.get('turn-1')).toEqual(nextContext)

    sessionState.live.turnActorUserIdByTurn.set('turn-1', 'casey@example.com')
    const caseyContext = createContext({ selection: { sheetName: 'Sheet2', address: 'A1' } as WorkbookAgentUiContext['selection'] })

    updateWorkbookAgentDurableUiContextFromUser({
      sessionState,
      context: caseyContext,
      userId: 'alex@example.com',
    })

    expect(sessionState.durable.context).toEqual(caseyContext)
    expect(sessionState.live.turnContextByTurn.get('turn-1')).toEqual(nextContext)
  })
})
