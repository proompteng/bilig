import { useCallback, useMemo, useState } from 'react'
import { mutators } from '@bilig/zero-sync'
import { WorkbookChangesPanel } from './WorkbookChangesPanel.js'
import { useWorkbookChanges, type ZeroWorkbookChangeQuerySource } from './use-workbook-changes.js'
import { selectWorkbookHistoryState, type WorkbookChangeEntry } from './workbook-changes-model.js'

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
  readonly onJump: (sheetName: string, address: string) => void
}) {
  const { currentUserId, documentId, enabled, onJump, sheetNames, zero } = input
  const changes = useWorkbookChanges({
    documentId,
    sheetNames,
    zero,
    enabled,
  })
  const [pendingRevisions, setPendingRevisions] = useState<readonly number[]>([])
  const [isUndoPending, setIsUndoPending] = useState(false)
  const [isRedoPending, setIsRedoPending] = useState(false)
  const changeCount = changes.length
  const historyState = useMemo(() => selectWorkbookHistoryState({ entries: changes, currentUserId }), [changes, currentUserId])

  const revertChange = useCallback(
    (change: WorkbookChangeEntry) => {
      if (!enabled || !change.canRevert || pendingRevisions.includes(change.revision)) {
        return
      }
      setPendingRevisions((current) => [...current, change.revision])
      const observer = observeZeroMutationResult(
        zero.mutate(
          mutators.workbook.revertChange({
            documentId,
            revision: change.revision,
          }),
        ),
      )
      void (observer ?? Promise.resolve()).finally(() => {
        setPendingRevisions((current) => current.filter((revision) => revision !== change.revision))
      })
    },
    [documentId, enabled, pendingRevisions, zero],
  )

  const changesPanel = useMemo(
    () => <WorkbookChangesPanel changes={changes} onJump={onJump} onRevert={revertChange} pendingRevisions={pendingRevisions} />,
    [changes, onJump, pendingRevisions, revertChange],
  )

  const undoLatestChange = useCallback(() => {
    if (!enabled || isUndoPending || historyState.undoRevision === null) {
      return
    }
    setIsUndoPending(true)
    const observer = observeZeroMutationResult(
      zero.mutate(
        mutators.workbook.undoLatestChange({
          documentId,
        }),
      ),
    )
    void (observer ?? Promise.resolve()).finally(() => {
      setIsUndoPending(false)
    })
  }, [documentId, enabled, historyState.undoRevision, isUndoPending, zero])

  const redoLatestChange = useCallback(() => {
    if (!enabled || isRedoPending || historyState.redoRevision === null) {
      return
    }
    setIsRedoPending(true)
    const observer = observeZeroMutationResult(
      zero.mutate(
        mutators.workbook.redoLatestChange({
          documentId,
        }),
      ),
    )
    void (observer ?? Promise.resolve()).finally(() => {
      setIsRedoPending(false)
    })
  }, [documentId, enabled, historyState.redoRevision, isRedoPending, zero])

  return {
    canRedo: historyState.canRedo && !isRedoPending,
    canUndo: historyState.canUndo && !isUndoPending,
    changeCount,
    changesPanel,
    redoLatestChange,
    undoLatestChange,
  }
}
