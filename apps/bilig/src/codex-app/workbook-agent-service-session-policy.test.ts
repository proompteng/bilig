import { describe, expect, it } from 'vitest'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'
import {
  chooseWorkbookAgentEvictionCandidates,
  resolveWorkbookAgentActiveTurnActorUserId,
  summarizeWorkbookAgentActiveTurnCounts,
} from './workbook-agent-service-session-policy.js'

function createSession(input: {
  threadId: string
  documentId?: string
  userId?: string
  activeTurnId?: string | null
  turnActorUserId?: string | null
  status?: WorkbookAgentThreadState['live']['status']
  lastAccessedAt?: number
  scope?: WorkbookAgentThreadState['scope']
}): WorkbookAgentThreadState {
  const activeTurnId = input.activeTurnId ?? null
  return {
    documentId: input.documentId ?? 'doc-1',
    userId: input.userId ?? 'alex@example.com',
    storageActorUserId: input.userId ?? 'alex@example.com',
    scope: input.scope ?? 'private',
    executionPolicy: 'autoApplyAll',
    threadId: input.threadId,
    durable: {
      context: null,
      entries: [],
      reviewQueueItems: [],
      executionRecords: [],
      workflowRuns: [],
    },
    live: {
      activeTurnId,
      status: input.status ?? 'idle',
      lastError: null,
      stagedPrivateBundleByTurn: new Map(),
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      turnActorUserIdByTurn:
        activeTurnId && input.turnActorUserId !== null
          ? new Map([[activeTurnId, input.turnActorUserId ?? input.userId ?? 'alex@example.com']])
          : new Map(),
      turnContextByTurn: new Map(),
      lastAccessedAt: input.lastAccessedAt ?? 0,
    },
  }
}

describe('workbook agent service session policy', () => {
  it('resolves the active turn actor user id', () => {
    expect(
      resolveWorkbookAgentActiveTurnActorUserId(
        createSession({
          threadId: 'thr-1',
          activeTurnId: 'turn-1',
          turnActorUserId: 'pat@example.com',
          status: 'inProgress',
        }),
      ),
    ).toBe('pat@example.com')
    expect(resolveWorkbookAgentActiveTurnActorUserId(createSession({ threadId: 'thr-2' }))).toBeNull()
  })

  it('summarizes active turn counts by user and document', () => {
    const counts = summarizeWorkbookAgentActiveTurnCounts(
      [
        createSession({ threadId: 'thr-1', activeTurnId: 'turn-1', status: 'inProgress', turnActorUserId: 'alex@example.com' }),
        createSession({ threadId: 'thr-2', activeTurnId: 'turn-2', status: 'inProgress', turnActorUserId: 'pat@example.com' }),
        createSession({ threadId: 'thr-3', documentId: 'doc-2', activeTurnId: 'turn-3', status: 'failed' }),
      ],
      { actorUserId: 'alex@example.com', documentId: 'doc-1' },
    )
    expect(counts.activeTurnsForUser).toBe(1)
    expect(counts.activeTurnsForDocument).toBe(2)
  })

  it('chooses oldest idle sessions without subscribers as eviction candidates', () => {
    const candidates = chooseWorkbookAgentEvictionCandidates(
      [
        createSession({ threadId: 'thr-1', status: 'idle', lastAccessedAt: 20 }),
        createSession({ threadId: 'thr-2', status: 'inProgress', activeTurnId: 'turn-2', lastAccessedAt: 10 }),
        createSession({ threadId: 'thr-3', status: 'idle', lastAccessedAt: 5 }),
      ],
      new Map([['thr-1', new Set()]]),
    )
    expect(candidates.map((session) => session.threadId)).toEqual(['thr-3', 'thr-1'])
  })
})
