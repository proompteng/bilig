import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import {
  canCancelWorkbookAgentWorkflowRun,
  canInterruptWorkbookAgentTurn,
  normalizeWorkbookAgentCommandIndexes,
  projectWorkbookAgentBundle,
  toWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewSummary,
  type WorkbookAgentReviewQueueItem,
} from '@bilig/agent-api'
import type {
  WorkbookAgentThreadSnapshot,
  WorkbookAgentTimelineEntry,
  WorkbookAgentThreadScope,
  WorkbookAgentThreadSummary,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowRun,
} from '@bilig/contracts'
import { createWorkbookPerfSession } from './perf/workbook-perf.js'
import { WorkbookAgentPanel } from './WorkbookAgentPanel.js'
import {
  useWorkbookAgentThreadSummaries,
  useWorkbookAgentWorkflowRuns,
  type ZeroWorkbookAgentSource,
} from './use-workbook-agent-durable-state.js'
import { useWorkbookAgentContextSync } from './use-workbook-agent-context-sync.js'
import {
  clearStoredDraft,
  clearStoredSession,
  draftKey,
  loadStoredDrafts,
  loadStoredSession,
  persistStoredDrafts,
  persistStoredSession,
} from './workbook-agent-pane-storage.js'
import { createWorkbookAgentClient } from './workbook-agent-client.js'
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
import {
  formatWorkbookAgentContextLabel,
  normalizeWorkbookAgentErrorMessage,
  readAppliedRevision,
  stringifyWorkbookAgentContextSyncKey,
} from './workbook-agent-pane-helpers.js'
import { useWorkbookAgentStream, type WorkbookAgentLiveSession } from './workbook-agent-stream.js'

export function useWorkbookAgentPane(input: {
  readonly currentUserId: string
  readonly documentId: string
  readonly enabled: boolean
  readonly apiEnabled?: boolean
  readonly getContext: () => WorkbookAgentUiContext
  readonly contextVersion?: number | string
  readonly contextProofVersion?: number | string
  readonly applyContext?: (context: WorkbookAgentUiContext) => void
  readonly previewCommandBundle: (bundle: WorkbookAgentCommandBundle) => Promise<WorkbookAgentPreviewSummary>
  readonly syncAuthoritativeRevision?: (revision: number) => Promise<void> | void
  readonly activeContextLabel?: string
  readonly zero?: ZeroWorkbookAgentSource
  readonly zeroEnabled?: boolean
}) {
  const {
    currentUserId,
    documentId,
    enabled,
    apiEnabled = true,
    getContext,
    contextVersion,
    contextProofVersion,
    applyContext,
    previewCommandBundle,
    syncAuthoritativeRevision,
    activeContextLabel = formatWorkbookAgentContextLabel(getContext()),
    zero,
    zeroEnabled = false,
  } = input
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
  const lastAppliedSnapshotContextKeyRef = useRef<string>('')
  const lastRequestedAuthoritativeRevisionRef = useRef(0)
  const lastDraftKeyRef = useRef<string | null>(null)
  const getContextRef = useRef(getContext)
  const applyContextRef = useRef(applyContext)
  const syncAuthoritativeRevisionRef = useRef(syncAuthoritativeRevision)
  const pendingSessionRequestRef = useRef<Promise<WorkbookAgentLiveSession> | null>(null)
  const promptSubmissionInFlightRef = useRef(false)
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
  const storageScope = useMemo(
    () => ({
      documentId,
      userId: currentUserId,
    }),
    [currentUserId, documentId],
  )
  const usesLiveThreadSummaries = zeroEnabled && Boolean(zero)
  const client = useMemo(() => createWorkbookAgentClient(documentId), [documentId])
  const { clearPendingContextSync, resetContextSync, scheduleContextSync } = useWorkbookAgentContextSync({
    client,
    documentId,
    enabled,
    getContextRef,
    sessionRef,
    snapshot,
  })

  useEffect(() => {
    getContextRef.current = getContext
  }, [getContext])

  useEffect(() => {
    applyContextRef.current = applyContext
  }, [applyContext])

  useEffect(() => {
    syncAuthoritativeRevisionRef.current = syncAuthoritativeRevision
  }, [syncAuthoritativeRevision])

  useEffect(() => {
    if (lastDraftKeyRef.current === activeDraftKey) {
      return
    }
    lastDraftKeyRef.current = activeDraftKey
    setDraft(loadStoredDrafts(storageScope)[activeDraftKey] ?? '')
  }, [activeDraftKey, storageScope])

  useEffect(() => {
    const drafts = loadStoredDrafts(storageScope)
    if (draft.length === 0) {
      if (!(activeDraftKey in drafts)) {
        return
      }
      delete drafts[activeDraftKey]
    } else {
      drafts[activeDraftKey] = draft
    }
    persistStoredDrafts(storageScope, drafts)
  }, [activeDraftKey, draft, storageScope])

  const activeThreadId = snapshot?.threadId ?? sessionRef.current?.threadId ?? null
  const zeroThreadSummaries = useWorkbookAgentThreadSummaries({
    currentUserId,
    documentId,
    zero: zeroSource,
    enabled: enabled && zeroEnabled && Boolean(zero),
  })
  const zeroWorkflowRuns = useWorkbookAgentWorkflowRuns({
    currentUserId,
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
  const canInterruptTurn =
    snapshot?.status === 'inProgress' &&
    canInterruptWorkbookAgentTurn({
      scope: snapshot.scope,
      ownerUserId: activeThreadSummary?.ownerUserId ?? null,
      actorUserId: currentUserId,
      turnActorUserId: snapshot.activeTurnActorUserId ?? null,
    })
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
  const canDismissReviewItem = sharedReviewOwnerUserId === null || sharedReviewOwnerUserId === currentUserId
  const canCancelWorkflowRun = useCallback(
    (run: WorkbookAgentWorkflowRun) =>
      canCancelWorkbookAgentWorkflowRun({
        scope: snapshot?.scope ?? 'private',
        ownerUserId: activeThreadSummary?.ownerUserId ?? null,
        actorUserId: currentUserId,
        startedByUserId: run.startedByUserId,
      }),
    [activeThreadSummary?.ownerUserId, currentUserId, snapshot?.scope],
  )

  const persistSessionSnapshot = useCallback(
    (nextSnapshot: WorkbookAgentThreadSnapshot) => {
      setSnapshot(nextSnapshot)
      setThreadScope(nextSnapshot.scope)
      const nextApplyContext = applyContextRef.current
      if (nextSnapshot.context && nextApplyContext) {
        const nextContextKey = stringifyWorkbookAgentContextSyncKey(nextSnapshot.context)
        if (lastAppliedSnapshotContextKeyRef.current !== nextContextKey) {
          lastAppliedSnapshotContextKeyRef.current = nextContextKey
          nextApplyContext(nextSnapshot.context)
        }
      }
      const appliedRevision = nextSnapshot.executionRecords.reduce<number>((maxRevision, record) => {
        const revision = readAppliedRevision(record)
        return revision === null ? maxRevision : Math.max(maxRevision, revision)
      }, 0)
      const nextSyncAuthoritativeRevision = syncAuthoritativeRevisionRef.current
      if (nextSyncAuthoritativeRevision && appliedRevision > lastRequestedAuthoritativeRevisionRef.current) {
        lastRequestedAuthoritativeRevisionRef.current = appliedRevision
        void nextSyncAuthoritativeRevision(appliedRevision)
      }
      persistStoredSession(storageScope, {
        threadId: nextSnapshot.threadId,
      })
      sessionRef.current = {
        threadId: nextSnapshot.threadId,
      }
    },
    [storageScope],
  )

  const { closeStream, connectStream, resetRecoveringStream } = useWorkbookAgentStream({
    client,
    perfSession,
    persistSessionSnapshot,
    sessionRef,
    setError,
    setIsLoading,
    setSnapshot,
  })

  const ensureSession = useCallback(async (): Promise<WorkbookAgentLiveSession> => {
    if (!apiEnabled) {
      throw new Error('Workbook assistant service is not configured for this app session.')
    }
    const activeSession = sessionRef.current
    if (activeSession) {
      return activeSession
    }
    const pendingSessionRequest = pendingSessionRequestRef.current
    if (pendingSessionRequest) {
      return await pendingSessionRequest
    }
    const nextSessionRequest = (async () => {
      setIsLoading(true)
      try {
        const nextSnapshot = await client.createSession(getContextRef.current(), threadScope)
        persistSessionSnapshot(nextSnapshot)
        connectStream(nextSnapshot.threadId)
        setError(null)
        const nextSession = {
          threadId: nextSnapshot.threadId,
        }
        sessionRef.current = nextSession
        return nextSession
      } finally {
        pendingSessionRequestRef.current = null
        setIsLoading(false)
      }
    })()
    pendingSessionRequestRef.current = nextSessionRequest
    return await nextSessionRequest
  }, [apiEnabled, client, connectStream, persistSessionSnapshot, threadScope])

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
        const nextSnapshot = await client.applyReviewItem(activeSession.threadId, activeReviewBundle.id, {
          appliedBy,
          commandIndexes: normalizedCommandIndexes,
        })
        persistSessionSnapshot(nextSnapshot)
        perfSession.markFirstAgentApplyVisible?.()
        setError(null)
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError))
      } finally {
        setIsApplyingReviewItem(false)
      }
    },
    [
      client,
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
    if (!enabled || !apiEnabled) {
      closeStream()
      resetContextSync()
      sessionRef.current = null
      pendingSessionRequestRef.current = null
      promptSubmissionInFlightRef.current = false
      resetRecoveringStream()
      setFetchedThreadSummaries([])
      setSnapshot(null)
      setIsLoading(false)
      return
    }
    let cancelled = false
    resetContextSync()
    const storedSession = loadStoredSession(storageScope)
    sessionRef.current = null
    pendingSessionRequestRef.current = null

    const bootstrapThreadSummaries = async () => {
      try {
        const nextThreadSummaries = await client.loadThreadSummaries()
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
        clearPendingContextSync()
      }
    }

    const bootstrapStoredSession = async () => {
      try {
        setIsLoading(true)
        const nextSnapshot = await client.loadThreadSnapshot(storedSession.threadId)
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
      clearPendingContextSync()
    }
  }, [
    clearPendingContextSync,
    closeStream,
    connectStream,
    client,
    documentId,
    enabled,
    persistSessionSnapshot,
    apiEnabled,
    resetContextSync,
    resetRecoveringStream,
    storageScope,
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
        const nextSnapshot = await client.loadThreadSnapshot(threadId)
        persistSessionSnapshot(nextSnapshot)
        connectStream(nextSnapshot.threadId)
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError))
      } finally {
        setIsLoading(false)
      }
    },
    [client, connectStream, persistSessionSnapshot],
  )

  const startNewThread = useCallback(() => {
    closeStream()
    clearStoredSession(storageScope)
    resetRecoveringStream()
    sessionRef.current = null
    pendingSessionRequestRef.current = null
    promptSubmissionInFlightRef.current = false
    setSnapshot(null)
    setPendingUserPrompt(null)
    setPreview(null)
    setSelectedCommandIndexes([])
    setError(null)
  }, [closeStream, resetRecoveringStream, storageScope])

  useEffect(() => {
    if (contextVersion !== undefined) {
      return
    }
    scheduleContextSync()
  })

  useEffect(() => {
    if (contextVersion === undefined) {
      return
    }
    scheduleContextSync()
  }, [activeContextLabel, contextVersion, enabled, scheduleContextSync, snapshot?.threadId])

  useEffect(() => {
    if (contextProofVersion === undefined || snapshot?.status !== 'inProgress' || snapshot.activeTurnId === null) {
      return
    }
    scheduleContextSync()
  }, [contextProofVersion, enabled, scheduleContextSync, snapshot?.activeTurnId, snapshot?.status, snapshot?.threadId])

  useEffect(() => {
    return () => {
      clearPendingContextSync()
    }
  }, [clearPendingContextSync])

  const sendPrompt = useCallback(async () => {
    if (promptSubmissionInFlightRef.current) {
      return
    }
    const prompt = draft.trim()
    if (prompt.length === 0) {
      return
    }
    promptSubmissionInFlightRef.current = true
    try {
      setError(null)
      setPendingUserPrompt(prompt)
      clearStoredDraft(storageScope, activeDraftKey)
      setDraft('')
      const activeSession = await ensureSession()
      const nextSnapshot = await client.sendPrompt(activeSession.threadId, prompt, getContextRef.current())
      persistSessionSnapshot(nextSnapshot)
      setPendingUserPrompt(null)
    } catch (nextError) {
      setPendingUserPrompt(null)
      setDraft(prompt)
      setError(nextError instanceof Error ? nextError.message : String(nextError))
    } finally {
      promptSubmissionInFlightRef.current = false
    }
  }, [activeDraftKey, client, draft, ensureSession, persistSessionSnapshot, storageScope])

  const interrupt = useCallback(async () => {
    const activeSession = sessionRef.current
    if (!activeSession) {
      return
    }
    try {
      const nextSnapshot = await client.interruptThread(activeSession.threadId)
      setSnapshot(nextSnapshot)
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : String(nextError))
    }
  }, [client])

  const dismissReviewItem = useCallback(async () => {
    const activeSession = sessionRef.current
    if (!activeSession || !activeReviewBundle) {
      return
    }
    try {
      const nextSnapshot = await client.dismissReviewItem(activeSession.threadId, activeReviewBundle.id)
      persistSessionSnapshot(nextSnapshot)
      setError(null)
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : String(nextError))
    }
  }, [client, activeReviewBundle, persistSessionSnapshot])

  const reviewReviewItem = useCallback(
    async (decision: 'approved' | 'rejected') => {
      const activeSession = sessionRef.current
      if (!activeSession || !activeReviewBundle) {
        return
      }
      try {
        const nextSnapshot = await client.reviewReviewItem(activeSession.threadId, activeReviewBundle.id, decision)
        persistSessionSnapshot(nextSnapshot)
        setError(null)
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError))
      }
    },
    [client, activeReviewBundle, persistSessionSnapshot],
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
        const nextSnapshot = await client.cancelWorkflowRun(activeSession.threadId, runId)
        persistSessionSnapshot(nextSnapshot)
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError))
      } finally {
        setCancellingWorkflowRunId((currentRunId) => (currentRunId === runId ? null : currentRunId))
      }
    },
    [client, persistSessionSnapshot],
  )

  const clearAgentError = useCallback(() => {
    setError(null)
  }, [])

  const agentPanel = useMemo(
    () => (
      <WorkbookAgentPanel
        activeThreadId={snapshot?.threadId ?? sessionRef.current?.threadId ?? null}
        activeContextLabel={activeContextLabel}
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
        canDismissReviewItem={canDismissReviewItem}
        optimisticEntries={optimisticEntries}
        selectedCommandIndexes={normalizedCommandIndexes}
        snapshot={snapshot}
        threadSummaries={threadSummaries}
        workflowRuns={workflowRuns}
        canInterruptTurn={canInterruptTurn}
        canCancelWorkflowRun={canCancelWorkflowRun}
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
      activeContextLabel,
      preview,
      activeThreadSummary?.ownerUserId,
      activeResponseTurnId,
      canFinalizeSharedBundle,
      canDismissReviewItem,
      canInterruptTurn,
      canRecommendSharedBundle,
      currentUserSharedRecommendation,
      canCancelWorkflowRun,
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
    activeThreadId,
    agentPanel,
    agentError: error ? normalizeWorkbookAgentErrorMessage(error) : null,
    clearAgentError,
    pendingCommandCount: reviewCommandCount,
    previewRanges: preview?.ranges ?? activeReviewBundle?.affectedRanges ?? [],
    selectThread,
    startNewThread,
    threadSummaries,
  }
}
