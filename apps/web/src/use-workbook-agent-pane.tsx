import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import {
  normalizeWorkbookAgentCommandIndexes,
  projectWorkbookAgentBundle,
  toWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewSummary,
  type WorkbookAgentReviewQueueItem,
} from '@bilig/agent-api'
import {
  WorkbookAgentThreadSnapshotSchema,
  WorkbookAgentStreamEventSchema,
  WorkbookAgentThreadScopeSchema,
  WorkbookAgentThreadSummarySchema,
  decodeUnknownSync,
  type WorkbookAgentThreadSnapshot,
  type WorkbookAgentStreamEvent,
  type WorkbookAgentTimelineEntry,
  type WorkbookAgentThreadScope,
  type WorkbookAgentThreadSummary,
  type WorkbookAgentUiContext,
  type WorkbookAgentWorkflowRun,
} from '@bilig/contracts'
import { Schema } from 'effect'
import { createWorkbookPerfSession } from './perf/workbook-perf.js'
import { WorkbookAgentPanel } from './WorkbookAgentPanel.js'
import {
  useWorkbookAgentThreadSummaries,
  useWorkbookAgentWorkflowRuns,
  type ZeroWorkbookAgentSource,
} from './use-workbook-agent-durable-state.js'
import {
  clearStoredDraft,
  clearStoredSession,
  draftKey,
  loadStoredDrafts,
  loadStoredSession,
  persistStoredDrafts,
  persistStoredSession,
  type StoredWorkbookAgentThreadRef,
} from './workbook-agent-pane-storage.js'
import {
  createWorkbookAgentPreviewRequestKey,
  loadWorkbookAgentPreview,
  readCachedWorkbookAgentPreview,
} from './workbook-agent-preview-cache.js'
import {
  decodeWorkbookAgentReviewItems,
  resolveWorkbookAgentReviewOwnerUserId,
  resolvePrimaryWorkbookAgentReviewItem,
} from './workbook-agent-review-state.js'
const WorkbookAgentThreadSummaryListSchema = Schema.Array(WorkbookAgentThreadSummarySchema)

interface WorkbookAgentLiveSession {
  threadId: string
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function resolvePayloadMessage(payload: unknown, fallback: string): string {
  return isRecord(payload) && typeof payload['message'] === 'string' ? payload['message'] : fallback
}

function readMessageEventData(event: Event): string | null {
  return event instanceof MessageEvent && typeof event.data === 'string' ? event.data : null
}

function createSessionResumeBody(
  storedSession: StoredWorkbookAgentThreadRef | null,
  context: WorkbookAgentUiContext,
  scope: WorkbookAgentThreadScope,
): {
  readonly threadId?: string
  readonly context: WorkbookAgentUiContext
  readonly scope?: WorkbookAgentThreadScope
} {
  return storedSession?.threadId
    ? {
        threadId: storedSession.threadId,
        context,
      }
    : {
        context,
        scope: decodeUnknownSync(WorkbookAgentThreadScopeSchema, scope),
      }
}

function threadSnapshotUrl(documentId: string, threadId: string): string {
  return `/v2/documents/${encodeURIComponent(documentId)}/chat/threads/${encodeURIComponent(threadId)}`
}

function createTextEntryFromDelta(event: Extract<WorkbookAgentStreamEvent, { type: 'entryTextDelta' }>) {
  return {
    id: event.itemId,
    kind: event.entryKind,
    turnId: event.turnId,
    text: event.delta,
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [],
  } satisfies WorkbookAgentTimelineEntry
}

function updateSnapshotFromDelta(
  snapshot: WorkbookAgentThreadSnapshot | null,
  event: Extract<WorkbookAgentStreamEvent, { type: 'entryTextDelta' }>,
): WorkbookAgentThreadSnapshot | null {
  if (!snapshot) {
    return snapshot
  }
  let matched = false
  return {
    ...snapshot,
    entries: (() => {
      const nextEntries = snapshot.entries.map((entry) => {
        if (entry.id !== event.itemId) {
          return entry
        }
        matched = true
        return {
          ...entry,
          kind: event.entryKind,
          turnId: event.turnId,
          text: `${entry.text ?? ''}${event.delta}`,
        }
      })
      return matched ? nextEntries : [...nextEntries, createTextEntryFromDelta(event)]
    })(),
  }
}

function normalizeWorkbookAgentErrorMessage(error: string): string {
  if (error.includes('thread/start.dynamicTools requires experimentalApi capability')) {
    return 'Retry in a moment.'
  }
  if (error.includes('Invalid Codex initialize response')) {
    return 'Retry in a moment.'
  }
  return error
}

export function useWorkbookAgentPane(input: {
  readonly currentUserId: string
  readonly documentId: string
  readonly enabled: boolean
  readonly getContext: () => WorkbookAgentUiContext
  readonly previewCommandBundle: (bundle: WorkbookAgentCommandBundle) => Promise<WorkbookAgentPreviewSummary>
  readonly zero?: ZeroWorkbookAgentSource
  readonly zeroEnabled?: boolean
}) {
  const { currentUserId, documentId, enabled, getContext, previewCommandBundle, zero, zeroEnabled = false } = input
  const [snapshot, setSnapshot] = useState<WorkbookAgentThreadSnapshot | null>(null)
  const [draft, setDraft] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)
  const [isApplyingReviewItem, setIsApplyingReviewItem] = useState(false)
  const [cancellingWorkflowRunId, setCancellingWorkflowRunId] = useState<string | null>(null)
  const [pendingUserPrompt, setPendingUserPrompt] = useState<string | null>(null)
  const [preview, setPreview] = useState<WorkbookAgentPreviewSummary | null>(null)
  const [selectedCommandIndexes, setSelectedCommandIndexes] = useState<number[] | null>(null)
  const [fetchedThreadSummaries, setFetchedThreadSummaries] = useState<readonly WorkbookAgentThreadSummary[]>([])
  const [threadScope, setThreadScope] = useState<WorkbookAgentThreadScope>('private')
  const previewRequestKeyRef = useRef<string | null>(null)
  const sessionRef = useRef<WorkbookAgentLiveSession | null>(null)
  const eventSourceRef = useRef<EventSource | null>(null)
  const recoveringStreamRef = useRef(false)
  const lastContextKeyRef = useRef<string>('')
  const lastDraftKeyRef = useRef<string | null>(null)
  const getContextRef = useRef(getContext)
  const currentContext = getContextRef.current()
  const activeDraftKey = draftKey(snapshot?.threadId ?? null, threadScope)
  const perfSession = useMemo(
    () =>
      createWorkbookPerfSession({
        documentId,
        scope: `bilig:${documentId}:agent-perf`,
      }),
    [documentId],
  )
  const zeroSource = useMemo<ZeroWorkbookAgentSource>(
    () =>
      zero ??
      ({
        materialize() {
          return {
            data: [],
            addListener() {
              return () => undefined
            },
            destroy() {},
          }
        },
      } satisfies ZeroWorkbookAgentSource),
    [zero],
  )
  const usesLiveThreadSummaries = zeroEnabled && Boolean(zero)

  useEffect(() => {
    getContextRef.current = getContext
  }, [getContext])

  useEffect(() => {
    if (lastDraftKeyRef.current === activeDraftKey) {
      return
    }
    lastDraftKeyRef.current = activeDraftKey
    setDraft(loadStoredDrafts(documentId)[activeDraftKey] ?? '')
  }, [activeDraftKey, documentId])

  useEffect(() => {
    const drafts = loadStoredDrafts(documentId)
    if (draft.length === 0) {
      if (!(activeDraftKey in drafts)) {
        return
      }
      delete drafts[activeDraftKey]
    } else {
      drafts[activeDraftKey] = draft
    }
    persistStoredDrafts(documentId, drafts)
  }, [activeDraftKey, documentId, draft])

  const loadThreadSummaries = useCallback(async (): Promise<readonly WorkbookAgentThreadSummary[]> => {
    const response = await fetch(`/v2/documents/${encodeURIComponent(documentId)}/chat/threads`)
    const payload = (await response.json()) as unknown
    if (!response.ok) {
      throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
    }
    return decodeUnknownSync(WorkbookAgentThreadSummaryListSchema, payload)
  }, [documentId])

  const activeThreadId = snapshot?.threadId ?? sessionRef.current?.threadId ?? null
  const zeroThreadSummaries = useWorkbookAgentThreadSummaries({
    documentId,
    zero: zeroSource,
    enabled: enabled && zeroEnabled && Boolean(zero),
  })
  const zeroWorkflowRuns = useWorkbookAgentWorkflowRuns({
    documentId,
    threadId: activeThreadId,
    zero: zeroSource,
    enabled: enabled && zeroEnabled && Boolean(zero) && activeThreadId !== null,
  })

  const reviewQueueItems = useMemo<WorkbookAgentReviewQueueItem[]>(() => {
    return decodeWorkbookAgentReviewItems(snapshot?.reviewQueueItems)
  }, [snapshot?.reviewQueueItems])
  const primaryReviewItem = useMemo(() => resolvePrimaryWorkbookAgentReviewItem(reviewQueueItems), [reviewQueueItems])
  const visibleReviewItem = useMemo(
    () => (primaryReviewItem && (snapshot?.scope ?? 'private') === 'shared' ? primaryReviewItem : null),
    [primaryReviewItem, snapshot?.scope],
  )
  const activeReviewBundle = useMemo<WorkbookAgentCommandBundle | null>(
    () => (visibleReviewItem ? toWorkbookAgentCommandBundle(visibleReviewItem) : null),
    [visibleReviewItem],
  )
  const reviewCommandCount = visibleReviewItem?.commands.length ?? 0

  const normalizedCommandIndexes = useMemo(
    () => (activeReviewBundle ? normalizeWorkbookAgentCommandIndexes(activeReviewBundle, selectedCommandIndexes) : []),
    [activeReviewBundle, selectedCommandIndexes],
  )

  const selectedReviewBundle = useMemo<WorkbookAgentCommandBundle | null>(
    () =>
      activeReviewBundle
        ? projectWorkbookAgentBundle({
            bundle: activeReviewBundle,
            commandIndexes: normalizedCommandIndexes,
            bundleId: activeReviewBundle.id,
          })
        : null,
    [normalizedCommandIndexes, activeReviewBundle],
  )
  const previewRequestKey = useMemo(() => {
    if (!activeReviewBundle || !selectedReviewBundle) {
      return null
    }
    return createWorkbookAgentPreviewRequestKey({
      bundle: activeReviewBundle,
      commandIndexes: normalizedCommandIndexes,
    })
  }, [normalizedCommandIndexes, activeReviewBundle, selectedReviewBundle])

  const threadSummaries = useMemo<readonly WorkbookAgentThreadSummary[]>(() => {
    const merged = new Map<string, WorkbookAgentThreadSummary>()
    for (const summary of fetchedThreadSummaries) {
      merged.set(summary.threadId, summary)
    }
    for (const summary of zeroThreadSummaries) {
      merged.set(summary.threadId, summary)
    }
    return [...merged.values()].toSorted((left, right) => right.updatedAtUnixMs - left.updatedAtUnixMs)
  }, [fetchedThreadSummaries, zeroThreadSummaries])
  const workflowRuns = useMemo<readonly WorkbookAgentWorkflowRun[]>(() => {
    const merged = new Map<string, WorkbookAgentWorkflowRun>()
    for (const run of snapshot?.workflowRuns ?? []) {
      merged.set(run.runId, run)
    }
    for (const run of zeroWorkflowRuns) {
      merged.set(run.runId, run)
    }
    return [...merged.values()].toSorted((left, right) => right.updatedAtUnixMs - left.updatedAtUnixMs)
  }, [snapshot?.workflowRuns, zeroWorkflowRuns])
  const activeThreadSummary = useMemo(
    () =>
      threadSummaries.find((threadSummary) => threadSummary.threadId === (snapshot?.threadId ?? sessionRef.current?.threadId ?? null)) ??
      null,
    [snapshot?.threadId, threadSummaries],
  )
  const sharedApplyRequiresOwnerApproval =
    snapshot?.scope === 'shared' &&
    visibleReviewItem?.riskClass !== undefined &&
    visibleReviewItem.riskClass !== 'low' &&
    activeThreadSummary !== null &&
    activeThreadSummary.ownerUserId !== currentUserId
  const sharedReviewOwnerUserId = resolveWorkbookAgentReviewOwnerUserId({
    reviewItem: visibleReviewItem,
    sessionScope: snapshot?.scope ?? 'private',
    activeThreadOwnerUserId: activeThreadSummary?.ownerUserId ?? null,
  })
  const sharedReviewStatus = sharedReviewOwnerUserId !== null ? (visibleReviewItem?.status ?? 'pending') : null
  const sharedReviewDecidedByUserId = sharedReviewOwnerUserId !== null ? (visibleReviewItem?.decidedByUserId ?? null) : null
  const sharedReviewRecommendations = useMemo(
    () => (sharedReviewOwnerUserId !== null ? (visibleReviewItem?.recommendations ?? []) : []),
    [sharedReviewOwnerUserId, visibleReviewItem?.recommendations],
  )
  const currentUserSharedRecommendation =
    sharedReviewRecommendations.find((recommendation) => recommendation.userId === currentUserId)?.decision ?? null
  const optimisticEntries = useMemo<readonly WorkbookAgentTimelineEntry[]>(() => {
    const entries: WorkbookAgentTimelineEntry[] = []
    const activeTurnId = snapshot?.activeTurnId ?? null
    const showOptimisticUser =
      pendingUserPrompt !== null && !snapshot?.entries.some((entry) => entry.kind === 'user' && entry.text === pendingUserPrompt)
    if (showOptimisticUser) {
      entries.push({
        id: 'optimistic-user:local',
        kind: 'user',
        turnId: activeTurnId,
        text: pendingUserPrompt,
        phase: null,
        toolName: null,
        toolStatus: null,
        argumentsText: null,
        outputText: null,
        success: null,
        citations: [],
      })
    }
    return entries
  }, [pendingUserPrompt, snapshot])
  const showAssistantProgress = snapshot?.status === 'inProgress'
  const activeResponseTurnId = showAssistantProgress ? (snapshot?.activeTurnId ?? null) : null
  const canFinalizeSharedBundle =
    sharedReviewOwnerUserId !== null &&
    sharedReviewOwnerUserId === currentUserId &&
    sharedReviewStatus === 'pending' &&
    !isApplyingReviewItem
  const canRecommendSharedBundle =
    sharedReviewOwnerUserId !== null &&
    sharedReviewOwnerUserId !== currentUserId &&
    sharedReviewStatus === 'pending' &&
    !isApplyingReviewItem

  const closeStream = useCallback(() => {
    eventSourceRef.current?.close()
    eventSourceRef.current = null
  }, [])

  const persistSessionSnapshot = useCallback(
    (nextSnapshot: WorkbookAgentThreadSnapshot) => {
      setSnapshot(nextSnapshot)
      setThreadScope(nextSnapshot.scope)
      persistStoredSession(documentId, {
        threadId: nextSnapshot.threadId,
      })
      sessionRef.current = {
        threadId: nextSnapshot.threadId,
      }
    },
    [documentId],
  )

  const connectStream = useCallback(
    (threadId: string) => {
      closeStream()
      const source = new EventSource(`/v2/documents/${encodeURIComponent(documentId)}/chat/threads/${encodeURIComponent(threadId)}/events`)
      source.addEventListener('message', (message) => {
        try {
          const payloadText = readMessageEventData(message)
          if (payloadText === null) {
            return
          }
          const event = decodeUnknownSync(WorkbookAgentStreamEventSchema, JSON.parse(payloadText))
          if (event.type === 'snapshot') {
            recoveringStreamRef.current = false
            persistSessionSnapshot(event.snapshot)
            setError(null)
            return
          }
          setSnapshot((current: WorkbookAgentThreadSnapshot | null) => updateSnapshotFromDelta(current, event))
          if (event.type === 'entryTextDelta' && event.entryKind === 'assistant') {
            perfSession.markFirstAssistantDeltaVisible?.()
          }
        } catch (nextError) {
          setError(nextError instanceof Error ? nextError.message : String(nextError))
        }
      })
      source.addEventListener('error', () => {
        if (eventSourceRef.current === source) {
          source.close()
          eventSourceRef.current = null
        }
        if (recoveringStreamRef.current) {
          return
        }
        const storedSession = sessionRef.current
        if (!storedSession) {
          setError('Assistant stream disconnected.')
          return
        }
        recoveringStreamRef.current = true
        setError(null)
        void (async () => {
          try {
            setIsLoading(true)
            const response = await fetch(threadSnapshotUrl(documentId, storedSession.threadId))
            const payload = (await response.json()) as unknown
            if (!response.ok) {
              throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
            }
            const nextSnapshot = decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload)
            persistSessionSnapshot(nextSnapshot)
            connectStream(nextSnapshot.threadId)
          } catch (nextError) {
            recoveringStreamRef.current = false
            setError(nextError instanceof Error ? nextError.message : String(nextError))
          } finally {
            setIsLoading(false)
          }
        })()
      })
      eventSourceRef.current = source
    },
    [closeStream, documentId, perfSession, persistSessionSnapshot],
  )

  const createSession = useCallback(
    async (context: WorkbookAgentUiContext, scope: WorkbookAgentThreadScope) => {
      const response = await fetch(`/v2/documents/${encodeURIComponent(documentId)}/chat/threads`, {
        method: 'POST',
        headers: {
          'content-type': 'application/json',
        },
        body: JSON.stringify(createSessionResumeBody(null, context, scope)),
      })
      const payload = (await response.json()) as unknown
      if (!response.ok) {
        throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
      }
      return decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload)
    },
    [documentId],
  )

  const loadThreadSnapshot = useCallback(
    async (threadId: string) => {
      const response = await fetch(threadSnapshotUrl(documentId, threadId))
      const payload = (await response.json()) as unknown
      if (!response.ok) {
        throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
      }
      return decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload)
    },
    [documentId],
  )

  const ensureSession = useCallback(async (): Promise<WorkbookAgentLiveSession> => {
    const activeSession = sessionRef.current
    if (activeSession) {
      return activeSession
    }
    setIsLoading(true)
    try {
      const nextSnapshot = await createSession(getContextRef.current(), threadScope)
      persistSessionSnapshot(nextSnapshot)
      connectStream(nextSnapshot.threadId)
      setError(null)
      const nextSession = {
        threadId: nextSnapshot.threadId,
      }
      sessionRef.current = nextSession
      return nextSession
    } finally {
      setIsLoading(false)
    }
  }, [connectStream, createSession, persistSessionSnapshot, threadScope])

  useEffect(() => {
    setSelectedCommandIndexes((current) => (current === null ? current : null))
  }, [activeReviewBundle?.id, reviewCommandCount])

  useEffect(() => {
    if (!enabled || selectedReviewBundle === null || previewRequestKey === null) {
      previewRequestKeyRef.current = null
      setPreview(null)
      return
    }
    const cachedPreview = readCachedWorkbookAgentPreview(previewRequestKey)
    if (cachedPreview) {
      previewRequestKeyRef.current = previewRequestKey
      setPreview(cachedPreview)
      return
    }
    if (previewRequestKeyRef.current === previewRequestKey) {
      return
    }
    previewRequestKeyRef.current = previewRequestKey
    let cancelled = false
    void (async () => {
      try {
        const nextPreview = await loadWorkbookAgentPreview({
          requestKey: previewRequestKey,
          load: () => previewCommandBundle(selectedReviewBundle),
        })
        if (!cancelled) {
          setPreview(nextPreview)
        }
      } catch (nextError: unknown) {
        if (!cancelled) {
          setPreview(null)
          setError(nextError instanceof Error ? nextError.message : String(nextError))
        }
      }
    })()
    return () => {
      cancelled = true
    }
  }, [enabled, previewCommandBundle, previewRequestKey, selectedReviewBundle])

  useEffect(() => {
    if (!preview) {
      return
    }
    perfSession.markFirstPreviewVisible?.()
  }, [perfSession, preview])

  const applyReviewItem = useCallback(
    async (appliedBy: 'user' | 'auto' = 'user') => {
      const activeSession = sessionRef.current
      if (!activeSession || !activeReviewBundle || !selectedReviewBundle || !preview) {
        return
      }
      if (sharedApplyRequiresOwnerApproval) {
        setError('Only the shared thread owner can approve medium/high-risk workbook changes.')
        return
      }
      if (sharedReviewStatus !== null && sharedReviewStatus !== 'approved') {
        setError('Approve the shared workbook changes before apply.')
        return
      }
      try {
        setIsApplyingReviewItem(true)
        const response = await fetch(
          `/v2/documents/${encodeURIComponent(documentId)}/chat/threads/${encodeURIComponent(activeSession.threadId)}/review-items/${encodeURIComponent(activeReviewBundle.id)}/apply`,
          {
            method: 'POST',
            headers: {
              'content-type': 'application/json',
            },
            body: JSON.stringify({
              appliedBy,
              commandIndexes: normalizedCommandIndexes,
              preview,
            }),
          },
        )
        const payload = (await response.json()) as unknown
        if (!response.ok) {
          throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
        }
        persistSessionSnapshot(decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload))
        perfSession.markFirstAgentApplyVisible?.()
        setError(null)
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError))
      } finally {
        setIsApplyingReviewItem(false)
      }
    },
    [
      documentId,
      normalizedCommandIndexes,
      activeReviewBundle,
      perfSession,
      persistSessionSnapshot,
      preview,
      selectedReviewBundle,
      sharedReviewStatus,
      sharedApplyRequiresOwnerApproval,
    ],
  )

  const toggleReviewCommand = useCallback(
    (commandIndex: number) => {
      setSelectedCommandIndexes((current) => {
        if (!activeReviewBundle || commandIndex < 0 || commandIndex >= activeReviewBundle.commands.length) {
          return current
        }
        const selected = new Set(normalizeWorkbookAgentCommandIndexes(activeReviewBundle, current))
        if (selected.has(commandIndex)) {
          selected.delete(commandIndex)
        } else {
          selected.add(commandIndex)
        }
        return activeReviewBundle.commands.flatMap((_command, index) => (selected.has(index) ? [index] : []))
      })
    },
    [activeReviewBundle],
  )

  const selectAllReviewCommands = useCallback(() => {
    setSelectedCommandIndexes(activeReviewBundle ? null : [])
  }, [activeReviewBundle])

  useEffect(() => {
    if (!enabled) {
      closeStream()
      sessionRef.current = null
      recoveringStreamRef.current = false
      setFetchedThreadSummaries([])
      setSnapshot(null)
      setIsLoading(false)
      return
    }
    let cancelled = false
    lastContextKeyRef.current = ''
    const storedSession = loadStoredSession(documentId)
    sessionRef.current = null

    const bootstrapThreadSummaries = async () => {
      try {
        const nextThreadSummaries = await loadThreadSummaries()
        if (!cancelled) {
          setFetchedThreadSummaries(nextThreadSummaries)
        }
      } catch {
        if (!cancelled) {
          setFetchedThreadSummaries([])
        }
      }
    }
    if (usesLiveThreadSummaries) {
      setFetchedThreadSummaries([])
    } else {
      void bootstrapThreadSummaries()
    }

    if (!storedSession) {
      setIsLoading(false)
      return () => {
        cancelled = true
        closeStream()
      }
    }

    const bootstrapStoredSession = async () => {
      try {
        setIsLoading(true)
        const nextSnapshot = await loadThreadSnapshot(storedSession.threadId)
        if (cancelled) {
          return
        }
        persistSessionSnapshot(nextSnapshot)
        connectStream(nextSnapshot.threadId)
        setError(null)
      } catch (nextError) {
        if (!cancelled) {
          setError(nextError instanceof Error ? nextError.message : String(nextError))
        }
      } finally {
        if (!cancelled) {
          setIsLoading(false)
        }
      }
    }
    void bootstrapStoredSession()
    return () => {
      cancelled = true
      closeStream()
    }
  }, [
    closeStream,
    connectStream,
    documentId,
    enabled,
    loadThreadSnapshot,
    loadThreadSummaries,
    persistSessionSnapshot,
    usesLiveThreadSummaries,
  ])

  const selectThread = useCallback(
    async (threadId: string) => {
      if (sessionRef.current?.threadId === threadId) {
        return
      }
      try {
        setIsLoading(true)
        setError(null)
        const nextSnapshot = await loadThreadSnapshot(threadId)
        persistSessionSnapshot(nextSnapshot)
        connectStream(nextSnapshot.threadId)
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError))
      } finally {
        setIsLoading(false)
      }
    },
    [connectStream, loadThreadSnapshot, persistSessionSnapshot],
  )

  const startNewThread = useCallback(() => {
    closeStream()
    clearStoredSession(documentId)
    recoveringStreamRef.current = false
    sessionRef.current = null
    setSnapshot(null)
    setPendingUserPrompt(null)
    setPreview(null)
    setSelectedCommandIndexes([])
    setError(null)
  }, [closeStream, documentId])

  useEffect(() => {
    if (!enabled || !snapshot) {
      return
    }
    const nextContextKey = JSON.stringify(currentContext)
    if (lastContextKeyRef.current === nextContextKey) {
      return
    }
    lastContextKeyRef.current = nextContextKey
    const timeout = window.setTimeout(() => {
      const activeSession = sessionRef.current
      if (!activeSession) {
        return
      }
      void (async () => {
        try {
          await fetch(
            `/v2/documents/${encodeURIComponent(documentId)}/chat/threads/${encodeURIComponent(activeSession.threadId)}/context`,
            {
              method: 'POST',
              headers: {
                'content-type': 'application/json',
              },
              body: JSON.stringify({
                context: currentContext,
              }),
            },
          )
        } catch {}
      })()
    }, 150)
    return () => {
      window.clearTimeout(timeout)
    }
  }, [currentContext, documentId, enabled, snapshot])

  const sendPrompt = useCallback(async () => {
    const prompt = draft.trim()
    if (prompt.length === 0) {
      return
    }
    try {
      setError(null)
      setPendingUserPrompt(prompt)
      clearStoredDraft(documentId, activeDraftKey)
      setDraft('')
      const activeSession = await ensureSession()
      const response = await fetch(
        `/v2/documents/${encodeURIComponent(documentId)}/chat/threads/${encodeURIComponent(activeSession.threadId)}/turns`,
        {
          method: 'POST',
          headers: {
            'content-type': 'application/json',
          },
          body: JSON.stringify({
            prompt,
            context: getContextRef.current(),
          }),
        },
      )
      const payload = (await response.json()) as unknown
      if (!response.ok) {
        throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
      }
      persistSessionSnapshot(decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload))
      setPendingUserPrompt(null)
    } catch (nextError) {
      setPendingUserPrompt(null)
      setDraft(prompt)
      setError(nextError instanceof Error ? nextError.message : String(nextError))
    }
  }, [activeDraftKey, documentId, draft, ensureSession, persistSessionSnapshot])

  const interrupt = useCallback(async () => {
    const activeSession = sessionRef.current
    if (!activeSession) {
      return
    }
    try {
      const response = await fetch(
        `/v2/documents/${encodeURIComponent(documentId)}/chat/threads/${encodeURIComponent(activeSession.threadId)}/interrupt`,
        {
          method: 'POST',
        },
      )
      const payload = (await response.json()) as unknown
      if (!response.ok) {
        throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
      }
      setSnapshot(decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload))
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : String(nextError))
    }
  }, [documentId])

  const dismissReviewItem = useCallback(async () => {
    const activeSession = sessionRef.current
    if (!activeSession || !activeReviewBundle) {
      return
    }
    try {
      const response = await fetch(
        `/v2/documents/${encodeURIComponent(documentId)}/chat/threads/${encodeURIComponent(activeSession.threadId)}/review-items/${encodeURIComponent(activeReviewBundle.id)}/dismiss`,
        {
          method: 'POST',
          headers: {
            'content-type': 'application/json',
          },
          body: '{}',
        },
      )
      const payload = (await response.json()) as unknown
      if (!response.ok) {
        throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
      }
      persistSessionSnapshot(decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload))
      setError(null)
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : String(nextError))
    }
  }, [documentId, activeReviewBundle, persistSessionSnapshot])

  const reviewReviewItem = useCallback(
    async (decision: 'approved' | 'rejected') => {
      const activeSession = sessionRef.current
      if (!activeSession || !activeReviewBundle) {
        return
      }
      try {
        const response = await fetch(
          `/v2/documents/${encodeURIComponent(documentId)}/chat/threads/${encodeURIComponent(activeSession.threadId)}/review-items/${encodeURIComponent(activeReviewBundle.id)}/review`,
          {
            method: 'POST',
            headers: {
              'content-type': 'application/json',
            },
            body: JSON.stringify({ decision }),
          },
        )
        const payload = (await response.json()) as unknown
        if (!response.ok) {
          throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
        }
        persistSessionSnapshot(decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload))
        setError(null)
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError))
      }
    },
    [documentId, activeReviewBundle, persistSessionSnapshot],
  )

  const cancelWorkflowRun = useCallback(
    async (runId: string) => {
      const activeSession = sessionRef.current
      if (!activeSession) {
        return
      }
      try {
        setError(null)
        setCancellingWorkflowRunId(runId)
        const response = await fetch(
          `/v2/documents/${encodeURIComponent(documentId)}/chat/threads/${encodeURIComponent(activeSession.threadId)}/workflows/${encodeURIComponent(runId)}/cancel`,
          {
            method: 'POST',
          },
        )
        const payload = (await response.json()) as unknown
        if (!response.ok) {
          throw new Error(resolvePayloadMessage(payload, `Workbook agent request failed with status ${response.status}`))
        }
        persistSessionSnapshot(decodeUnknownSync(WorkbookAgentThreadSnapshotSchema, payload))
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError))
      } finally {
        setCancellingWorkflowRunId((currentRunId) => (currentRunId === runId ? null : currentRunId))
      }
    },
    [documentId, persistSessionSnapshot],
  )

  const clearAgentError = useCallback(() => {
    setError(null)
  }, [])

  const agentPanel = useMemo(
    () => (
      <WorkbookAgentPanel
        activeThreadId={snapshot?.threadId ?? sessionRef.current?.threadId ?? null}
        activeResponseTurnId={activeResponseTurnId}
        draft={draft}
        cancellingWorkflowRunId={cancellingWorkflowRunId}
        isApplyingReviewItem={isApplyingReviewItem}
        isLoading={isLoading}
        activeReviewBundle={activeReviewBundle}
        preview={preview}
        showAssistantProgress={showAssistantProgress}
        sharedApprovalOwnerUserId={sharedApplyRequiresOwnerApproval ? (activeThreadSummary?.ownerUserId ?? null) : null}
        sharedReviewOwnerUserId={sharedReviewOwnerUserId}
        sharedReviewStatus={sharedReviewStatus}
        sharedReviewDecidedByUserId={sharedReviewDecidedByUserId}
        sharedReviewRecommendations={sharedReviewRecommendations}
        currentUserSharedRecommendation={currentUserSharedRecommendation}
        canFinalizeSharedBundle={canFinalizeSharedBundle}
        canRecommendSharedBundle={canRecommendSharedBundle}
        optimisticEntries={optimisticEntries}
        selectedCommandIndexes={normalizedCommandIndexes}
        snapshot={snapshot}
        threadSummaries={threadSummaries}
        workflowRuns={workflowRuns}
        onApplyReviewItem={() => {
          void applyReviewItem('user')
        }}
        onDraftChange={setDraft}
        onDismissReviewItem={() => {
          void dismissReviewItem()
        }}
        onReviewReviewItem={(decision) => {
          void reviewReviewItem(decision)
        }}
        onInterrupt={() => {
          void interrupt()
        }}
        onSelectAllReviewCommands={selectAllReviewCommands}
        onToggleReviewCommand={toggleReviewCommand}
        onCancelWorkflowRun={(runId) => {
          void cancelWorkflowRun(runId)
        }}
        onSelectThread={(threadId) => {
          void selectThread(threadId)
        }}
        onSubmit={() => {
          void sendPrompt()
        }}
      />
    ),
    [
      applyReviewItem,
      cancelWorkflowRun,
      cancellingWorkflowRunId,
      dismissReviewItem,
      draft,
      interrupt,
      isApplyingReviewItem,
      isLoading,
      normalizedCommandIndexes,
      optimisticEntries,
      activeReviewBundle,
      preview,
      activeThreadSummary?.ownerUserId,
      activeResponseTurnId,
      canFinalizeSharedBundle,
      canRecommendSharedBundle,
      currentUserSharedRecommendation,
      reviewReviewItem,
      sendPrompt,
      selectThread,
      snapshot,
      sharedReviewDecidedByUserId,
      sharedReviewOwnerUserId,
      sharedReviewRecommendations,
      sharedReviewStatus,
      sharedApplyRequiresOwnerApproval,
      showAssistantProgress,
      selectAllReviewCommands,
      threadSummaries,
      workflowRuns,
      toggleReviewCommand,
    ],
  )

  return {
    agentPanel,
    agentError: error ? normalizeWorkbookAgentErrorMessage(error) : null,
    clearAgentError,
    pendingCommandCount: reviewCommandCount,
    previewRanges: preview?.ranges ?? activeReviewBundle?.affectedRanges ?? [],
    startNewThread,
  }
}
