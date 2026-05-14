import { useCallback, useRef, type MutableRefObject } from 'react'
import type { WorkbookAgentThreadSnapshot, WorkbookAgentUiContext } from '@bilig/contracts'
import { logDebug } from './runtime-logger.js'
import { stringifyWorkbookAgentContextSyncKey } from './workbook-agent-pane-helpers.js'

const AGENT_CONTEXT_SYNC_DEBOUNCE_MS = 150
const AGENT_CONTEXT_SYNC_MIN_INTERVAL_MS = 750

interface WorkbookAgentContextSyncClient {
  syncThreadContext(threadId: string, context: WorkbookAgentUiContext): Promise<void>
}

interface WorkbookAgentContextSyncSession {
  readonly threadId: string
}

function stringifyWorkbookAgentImmediateContextKey(context: WorkbookAgentUiContext): string {
  return JSON.stringify({
    selection: context.selection,
    viewport: context.viewport,
  })
}

export function useWorkbookAgentContextSync(input: {
  readonly client: WorkbookAgentContextSyncClient
  readonly documentId: string
  readonly enabled: boolean
  readonly getContextRef: MutableRefObject<() => WorkbookAgentUiContext>
  readonly sessionRef: MutableRefObject<WorkbookAgentContextSyncSession | null>
  readonly snapshot: WorkbookAgentThreadSnapshot | null
}) {
  const lastContextKeyRef = useRef<string>('')
  const lastImmediateContextKeyRef = useRef<string>('')
  const hasSyncedContextRef = useRef(false)
  const lastContextSyncAtRef = useRef(0)
  const pendingContextSyncTimeoutRef = useRef<number | null>(null)
  const pendingContextSyncRef = useRef<{
    readonly context: WorkbookAgentUiContext
    readonly key: string
    readonly immediateKey: string
    readonly threadId: string
  } | null>(null)

  const clearPendingContextSync = useCallback(() => {
    const pendingTimeout = pendingContextSyncTimeoutRef.current
    if (pendingTimeout !== null) {
      window.clearTimeout(pendingTimeout)
      pendingContextSyncTimeoutRef.current = null
    }
    pendingContextSyncRef.current = null
  }, [])

  const resetContextSync = useCallback(() => {
    clearPendingContextSync()
    lastContextKeyRef.current = ''
    lastImmediateContextKeyRef.current = ''
    hasSyncedContextRef.current = false
    lastContextSyncAtRef.current = 0
  }, [clearPendingContextSync])

  const flushPendingContextSync = useCallback(() => {
    const pending = pendingContextSyncRef.current
    pendingContextSyncTimeoutRef.current = null
    pendingContextSyncRef.current = null
    if (!pending || lastContextKeyRef.current === pending.key) {
      return
    }
    lastContextKeyRef.current = pending.key
    lastImmediateContextKeyRef.current = pending.immediateKey
    hasSyncedContextRef.current = true
    lastContextSyncAtRef.current = window.performance.now()
    void (async () => {
      try {
        await input.client.syncThreadContext(pending.threadId, pending.context)
      } catch (syncError) {
        logDebug('Failed to sync agent context update', { documentId: input.documentId, error: syncError })
      }
    })()
  }, [input.client, input.documentId])

  const scheduleContextSync = useCallback(() => {
    const activeSession = input.sessionRef.current
    if (!input.enabled || !input.snapshot || !activeSession) {
      clearPendingContextSync()
      return
    }
    const nextContext = input.getContextRef.current()
    const nextContextKey = `${activeSession.threadId}:${stringifyWorkbookAgentContextSyncKey(nextContext)}`
    const nextImmediateContextKey = `${activeSession.threadId}:${stringifyWorkbookAgentImmediateContextKey(nextContext)}`
    if (lastContextKeyRef.current === nextContextKey) {
      return
    }
    pendingContextSyncRef.current = {
      context: nextContext,
      key: nextContextKey,
      immediateKey: nextImmediateContextKey,
      threadId: activeSession.threadId,
    }
    const shouldPrioritizeSync = !hasSyncedContextRef.current || lastImmediateContextKeyRef.current !== nextImmediateContextKey
    const elapsedSinceLastSync = window.performance.now() - lastContextSyncAtRef.current
    const delayMs = shouldPrioritizeSync
      ? AGENT_CONTEXT_SYNC_DEBOUNCE_MS
      : Math.max(AGENT_CONTEXT_SYNC_DEBOUNCE_MS, AGENT_CONTEXT_SYNC_MIN_INTERVAL_MS - elapsedSinceLastSync)
    if (pendingContextSyncTimeoutRef.current !== null) {
      window.clearTimeout(pendingContextSyncTimeoutRef.current)
    }
    pendingContextSyncTimeoutRef.current = window.setTimeout(flushPendingContextSync, delayMs)
  }, [clearPendingContextSync, flushPendingContextSync, input.enabled, input.getContextRef, input.sessionRef, input.snapshot])

  return {
    clearPendingContextSync,
    resetContextSync,
    scheduleContextSync,
  }
}
