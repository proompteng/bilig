import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { mutators } from '@bilig/zero-sync'
import { WorkbookChangesPanel } from './WorkbookChangesPanel.js'
import { useWorkbookChanges, type ZeroWorkbookChangeQuerySource } from './use-workbook-changes.js'
import { selectWorkbookHistoryState } from './workbook-changes-model.js'

const QUEUED_HISTORY_SHORTCUT_TIMEOUT_MS = 10_000

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function observeZeroMutationResult(result: unknown): Promise<unknown> | null {
  if (!isRecord(result)) {
    return null
  }
  const observer = result['server'] ?? result['client']
  return observer instanceof Promise ? observer : null
}

export interface ZeroWorkbookChangeSource extends ZeroWorkbookChangeQuerySource {
  mutate(mutation: unknown): unknown
}

export function useWorkbookChangesPane(input: {
  readonly documentId: string
  readonly currentUserId: string
  readonly sheetNames: readonly string[]
  readonly zero: ZeroWorkbookChangeSource
  readonly enabled: boolean
  readonly pendingMutationSummary?:
    | {
        readonly activeCount: number
        readonly failedCount: number
      }
    | undefined
  readonly localMutationEpoch?: number | undefined
  readonly localMutationEpochRef?: { readonly current: number } | undefined
  readonly onHistoryMutationApplied?: (() => void | Promise<void>) | undefined
  readonly onJump: (sheetName: string, address: string) => void
}) {
  const {
    currentUserId,
    documentId,
    enabled,
    localMutationEpoch = 0,
    localMutationEpochRef: externalLocalMutationEpochRef,
    onHistoryMutationApplied,
    onJump,
    pendingMutationSummary,
    sheetNames,
    zero,
  } = input
  const changes = useWorkbookChanges({
    documentId,
    sheetNames,
    zero,
    enabled,
  })
  const [isUndoPending, setIsUndoPending] = useState(false)
  const [isRedoPending, setIsRedoPending] = useState(false)
  const [pendingRevertRevision, setPendingRevertRevision] = useState<number | null>(null)
  const [queuedHistoryShortcut, setQueuedHistoryShortcut] = useState<'undo' | 'redo' | null>(null)
  const changeCount = changes.entries.length
  const historyState = useMemo(() => selectWorkbookHistoryState({ rows: changes.rows, currentUserId }), [changes.rows, currentUserId])
  const hasActivePendingMutation = (pendingMutationSummary?.activeCount ?? 0) > 0
  const localMutationEpochRef = useRef(localMutationEpoch)
  const [historyReadyEpoch, setHistoryReadyEpoch] = useState(localMutationEpoch)
  localMutationEpochRef.current = localMutationEpoch
  const historyRevisionSignature = `${changeCount}:${historyState.undoRevision ?? 'none'}:${historyState.redoRevision ?? 'none'}`
  const hasUnmaterializedLocalMutation = localMutationEpoch > historyReadyEpoch
  const readLocalMutationEpoch = useCallback(
    () => externalLocalMutationEpochRef?.current ?? localMutationEpochRef.current,
    [externalLocalMutationEpochRef],
  )

  useEffect(() => {
    setHistoryReadyEpoch(readLocalMutationEpoch())
  }, [historyRevisionSignature, readLocalMutationEpoch])

  const runUndoLatestChange = useCallback(() => {
    setIsUndoPending(true)
    const observer = observeZeroMutationResult(
      zero.mutate(
        mutators.workbook.undoLatestChange({
          documentId,
        }),
      ),
    )
    void (async () => {
      try {
        await (observer ?? Promise.resolve())
        await onHistoryMutationApplied?.()
      } finally {
        setIsUndoPending(false)
      }
    })()
  }, [documentId, onHistoryMutationApplied, zero])

  const runRedoLatestChange = useCallback(() => {
    setIsRedoPending(true)
    const observer = observeZeroMutationResult(
      zero.mutate(
        mutators.workbook.redoLatestChange({
          documentId,
        }),
      ),
    )
    void (async () => {
      try {
        await (observer ?? Promise.resolve())
        await onHistoryMutationApplied?.()
      } finally {
        setIsRedoPending(false)
      }
    })()
  }, [documentId, onHistoryMutationApplied, zero])

  const undoLatestChange = useCallback(() => {
    if (!enabled || isUndoPending) {
      return
    }
    const hasLiveUnmaterializedLocalMutation = readLocalMutationEpoch() > historyReadyEpoch
    if (historyState.undoRevision === null || hasLiveUnmaterializedLocalMutation) {
      if (hasActivePendingMutation || hasLiveUnmaterializedLocalMutation) {
        setQueuedHistoryShortcut('undo')
      }
      return
    }
    setQueuedHistoryShortcut(null)
    runUndoLatestChange()
  }, [
    enabled,
    hasActivePendingMutation,
    historyReadyEpoch,
    historyState.undoRevision,
    isUndoPending,
    readLocalMutationEpoch,
    runUndoLatestChange,
  ])

  const redoLatestChange = useCallback(() => {
    if (!enabled || isRedoPending) {
      return
    }
    const hasLiveUnmaterializedLocalMutation = readLocalMutationEpoch() > historyReadyEpoch
    if (historyState.redoRevision === null || hasLiveUnmaterializedLocalMutation) {
      if (hasActivePendingMutation || hasLiveUnmaterializedLocalMutation) {
        setQueuedHistoryShortcut('redo')
      }
      return
    }
    setQueuedHistoryShortcut(null)
    runRedoLatestChange()
  }, [
    enabled,
    hasActivePendingMutation,
    historyReadyEpoch,
    historyState.redoRevision,
    isRedoPending,
    readLocalMutationEpoch,
    runRedoLatestChange,
  ])

  useEffect(() => {
    if (queuedHistoryShortcut === null) {
      return
    }
    const timeout = window.setTimeout(() => setQueuedHistoryShortcut(null), QUEUED_HISTORY_SHORTCUT_TIMEOUT_MS)
    return () => window.clearTimeout(timeout)
  }, [queuedHistoryShortcut])

  useEffect(() => {
    if (!enabled || queuedHistoryShortcut === null || hasUnmaterializedLocalMutation) {
      return
    }
    if (queuedHistoryShortcut === 'undo') {
      if (historyState.undoRevision !== null && !isUndoPending) {
        setQueuedHistoryShortcut(null)
        runUndoLatestChange()
        return
      }
      return
    }
    if (historyState.redoRevision !== null && !isRedoPending) {
      setQueuedHistoryShortcut(null)
      runRedoLatestChange()
      return
    }
  }, [
    enabled,
    historyState.redoRevision,
    historyState.undoRevision,
    hasUnmaterializedLocalMutation,
    isRedoPending,
    isUndoPending,
    queuedHistoryShortcut,
    runRedoLatestChange,
    runUndoLatestChange,
  ])

  const revertChangeRevision = useCallback(
    (revision: number) => {
      const targetChange = changes.entries.find((change) => change.revision === revision)
      if (!enabled || pendingRevertRevision !== null || !targetChange?.canRevert) {
        return
      }
      setPendingRevertRevision(revision)
      const observer = observeZeroMutationResult(
        zero.mutate(
          mutators.workbook.revertChange({
            documentId,
            revision,
          }),
        ),
      )
      void (async () => {
        try {
          await (observer ?? Promise.resolve())
        } finally {
          setPendingRevertRevision(null)
        }
      })()
    },
    [changes.entries, documentId, enabled, pendingRevertRevision, zero],
  )

  const changesPanel = useMemo(
    () => (
      <WorkbookChangesPanel
        changes={changes.entries}
        pendingRevertRevision={pendingRevertRevision}
        onJump={onJump}
        onRevert={revertChangeRevision}
      />
    ),
    [changes.entries, onJump, pendingRevertRevision, revertChangeRevision],
  )

  return {
    canRedo: historyState.canRedo && !isRedoPending,
    canUndo: historyState.canUndo && !isUndoPending,
    changeCount,
    changesPanel,
    redoLatestChange,
    undoLatestChange,
  }
}
