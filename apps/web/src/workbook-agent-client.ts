import {
  WorkbookAgentThreadSnapshotSchema,
  WorkbookAgentThreadScopeSchema,
  WorkbookAgentThreadSummarySchema,
  decodeUnknownSync,
  stringifyWorkbookAgentUiContextSemanticKey,
  type WorkbookAgentThreadSnapshot,
  type WorkbookAgentThreadScope,
  type WorkbookAgentThreadSummary,
  type WorkbookAgentUiContext,
} from '@bilig/contracts'
import { Schema } from 'effect'

const WorkbookAgentThreadSummaryListSchema = Schema.Array(WorkbookAgentThreadSummarySchema)
const JSON_HEADERS = {
  'content-type': 'application/json',
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function resolvePayloadMessage(payload: unknown, fallback: string): string {
  return isRecord(payload) && typeof payload['message'] === 'string' ? payload['message'] : fallback
}

async function readJsonPayload(response: Response): Promise<unknown> {
  try {
    return (await response.json()) as unknown
  } catch {
    if (!response.ok) {
      return null
    }
    throw new Error('Workbook agent request returned malformed JSON')
  }
}

async function fetchJson(input: RequestInfo | URL, init?: RequestInit): Promise<unknown> {
  const response = init ? await fetch(input, init) : await fetch(input)
  const payload = await readJsonPayload(response)
  if (!response.ok) {
    throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
  }
  return payload
}

async function fetchOk(input: RequestInfo | URL, init?: RequestInit): Promise<void> {
  const response = init ? await fetch(input, init) : await fetch(input)
  if (response.ok) {
    return
  }
  let payload: unknown = null
  try {
    payload = await response.json()
  } catch {
    payload = null
  }
  throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
}

function threadListUrl(documentId: string): string {
  return `/v2/documents/${encodeURIComponent(documentId)}/chat/threads`
}

function threadUrl(documentId: string, threadId: string): string {
  return `${threadListUrl(documentId)}/${encodeURIComponent(threadId)}`
}

function decodeThreadSummaries(payload: unknown): readonly WorkbookAgentThreadSummary[] {
  try {
    return decodeUnknownSync(WorkbookAgentThreadSummaryListSchema, payload)
  } catch {
    throw new Error('Workbook agent request returned invalid thread summaries')
  }
}

function decodeThreadSnapshot(payload: unknown): WorkbookAgentThreadSnapshot {
  try {
    return decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload)
  } catch {
    throw new Error('Workbook agent request returned invalid thread snapshot')
  }
}

function createSessionBody(context: WorkbookAgentUiContext, scope: WorkbookAgentThreadScope) {
  return {
    context,
    scope: decodeUnknownSync(WorkbookAgentThreadScopeSchema, scope),
  }
}

interface PendingContextSync {
  context: WorkbookAgentUiContext
  key: string
  reject: (reason?: unknown) => void
  resolve: () => void
}

interface ThreadContextSyncState {
  inFlight: Promise<void> | null
  lastSyncedKey: string | null
  pending: PendingContextSync | null
}

function createThreadContextSyncState(): ThreadContextSyncState {
  return {
    inFlight: null,
    lastSyncedKey: null,
    pending: null,
  }
}

function contextSyncKey(threadId: string, context: WorkbookAgentUiContext): string {
  return `${threadId}:${stringifyWorkbookAgentUiContextSemanticKey(context)}`
}

export function createWorkbookAgentClient(documentId: string) {
  const contextSyncStates = new Map<string, ThreadContextSyncState>()

  const postThreadContext = async (threadId: string, context: WorkbookAgentUiContext): Promise<void> => {
    await fetchOk(`${threadUrl(documentId, threadId)}/context`, {
      method: 'POST',
      headers: JSON_HEADERS,
      body: JSON.stringify({
        context,
      }),
    })
  }

  const drainThreadContextSync = (threadId: string, state: ThreadContextSyncState): void => {
    const pending = state.pending
    if (!pending || state.inFlight) {
      return
    }
    state.pending = null
    if (state.lastSyncedKey === pending.key) {
      pending.resolve()
      return
    }
    state.inFlight = (async () => {
      try {
        await postThreadContext(threadId, pending.context)
        state.lastSyncedKey = pending.key
        pending.resolve()
      } catch (error: unknown) {
        pending.reject(error)
      } finally {
        state.inFlight = null
        drainThreadContextSync(threadId, state)
      }
    })()
  }

  const enqueueThreadContextSync = (threadId: string, context: WorkbookAgentUiContext): Promise<void> => {
    const key = contextSyncKey(threadId, context)
    const existingState = contextSyncStates.get(threadId)
    const state = existingState ?? createThreadContextSyncState()
    if (!existingState) {
      contextSyncStates.set(threadId, state)
    }
    if (state.lastSyncedKey === key) {
      return Promise.resolve()
    }
    if (!state.inFlight) {
      state.inFlight = (async () => {
        try {
          await postThreadContext(threadId, context)
          state.lastSyncedKey = key
        } finally {
          state.inFlight = null
          drainThreadContextSync(threadId, state)
        }
      })()
      return state.inFlight
    }

    return new Promise<void>((resolve, reject) => {
      const previousPending = state.pending
      state.pending = {
        context,
        key,
        reject(error) {
          previousPending?.reject(error)
          reject(error)
        },
        resolve() {
          previousPending?.resolve()
          resolve()
        },
      }
    })
  }

  return {
    threadEventsUrl(threadId: string): string {
      return `${threadUrl(documentId, threadId)}/events`
    },
    async loadThreadSummaries(): Promise<readonly WorkbookAgentThreadSummary[]> {
      return decodeThreadSummaries(await fetchJson(threadListUrl(documentId)))
    },
    async createSession(context: WorkbookAgentUiContext, scope: WorkbookAgentThreadScope): Promise<WorkbookAgentThreadSnapshot> {
      return decodeThreadSnapshot(
        await fetchJson(threadListUrl(documentId), {
          method: 'POST',
          headers: JSON_HEADERS,
          body: JSON.stringify(createSessionBody(context, scope)),
        }),
      )
    },
    async loadThreadSnapshot(threadId: string): Promise<WorkbookAgentThreadSnapshot> {
      return decodeThreadSnapshot(await fetchJson(threadUrl(documentId, threadId)))
    },
    async syncThreadContext(threadId: string, context: WorkbookAgentUiContext): Promise<void> {
      await enqueueThreadContextSync(threadId, context)
    },
    async sendPrompt(threadId: string, prompt: string, context: WorkbookAgentUiContext): Promise<WorkbookAgentThreadSnapshot> {
      return decodeThreadSnapshot(
        await fetchJson(`${threadUrl(documentId, threadId)}/turns`, {
          method: 'POST',
          headers: JSON_HEADERS,
          body: JSON.stringify({
            prompt,
            context,
          }),
        }),
      )
    },
    async interruptThread(threadId: string): Promise<WorkbookAgentThreadSnapshot> {
      return decodeThreadSnapshot(
        await fetchJson(`${threadUrl(documentId, threadId)}/interrupt`, {
          method: 'POST',
        }),
      )
    },
    async applyReviewItem(
      threadId: string,
      reviewItemId: string,
      input: {
        readonly appliedBy: 'user' | 'auto'
        readonly commandIndexes: readonly number[]
      },
    ): Promise<WorkbookAgentThreadSnapshot> {
      return decodeThreadSnapshot(
        await fetchJson(`${threadUrl(documentId, threadId)}/review-items/${encodeURIComponent(reviewItemId)}/apply`, {
          method: 'POST',
          headers: JSON_HEADERS,
          body: JSON.stringify(input),
        }),
      )
    },
    async dismissReviewItem(threadId: string, reviewItemId: string): Promise<WorkbookAgentThreadSnapshot> {
      return decodeThreadSnapshot(
        await fetchJson(`${threadUrl(documentId, threadId)}/review-items/${encodeURIComponent(reviewItemId)}/dismiss`, {
          method: 'POST',
          headers: JSON_HEADERS,
          body: '{}',
        }),
      )
    },
    async reviewReviewItem(
      threadId: string,
      reviewItemId: string,
      decision: 'approved' | 'rejected',
    ): Promise<WorkbookAgentThreadSnapshot> {
      return decodeThreadSnapshot(
        await fetchJson(`${threadUrl(documentId, threadId)}/review-items/${encodeURIComponent(reviewItemId)}/review`, {
          method: 'POST',
          headers: JSON_HEADERS,
          body: JSON.stringify({ decision }),
        }),
      )
    },
    async cancelWorkflowRun(threadId: string, runId: string): Promise<WorkbookAgentThreadSnapshot> {
      return decodeThreadSnapshot(
        await fetchJson(`${threadUrl(documentId, threadId)}/workflows/${encodeURIComponent(runId)}/cancel`, {
          method: 'POST',
        }),
      )
    },
  }
}
