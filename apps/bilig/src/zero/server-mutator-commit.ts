import { isDeepStrictEqual } from 'node:util'
import type { SpreadsheetEngine } from '@bilig/core'
import type { EngineOp } from '@bilig/workbook-domain'
import type { WorkbookChangeUndoBundle, WorkbookEventPayload } from '@bilig/zero-sync'
import type { SessionIdentity } from '../http/session.js'
import type { WorkbookRuntimeManager } from '../workbook-runtime/runtime-manager.js'
import type { Queryable } from './store.js'
import { acquireWorkbookMutationLock } from './workbook-runtime-store.js'
import {
  loadAppliedWorkbookClientMutation,
  persistWorkbookMutation,
  type AppliedWorkbookClientMutation,
} from './workbook-mutation-store.js'
import type { WorkbookChangeRange } from './workbook-change-store.js'

export interface ServerTransactionLike {
  dbTransaction: {
    wrappedTransaction: Queryable
  }
}

export function requireServerTransaction(tx: unknown): ServerTransactionLike {
  if (!isRecord(tx) || !isRecord(tx['dbTransaction']) || !isQueryable(tx['dbTransaction']['wrappedTransaction'])) {
    throw new Error('Expected a server-side Zero transaction')
  }

  return {
    dbTransaction: {
      wrappedTransaction: tx['dbTransaction']['wrappedTransaction'],
    },
  }
}

export function toEngineUndoBundle(undoOps: readonly EngineOp[] | null): WorkbookChangeUndoBundle | null {
  if (!undoOps || undoOps.length === 0) {
    return null
  }
  return {
    kind: 'engineOps',
    ops: structuredClone([...undoOps]),
  }
}

export function captureEngineUndoBundle(
  engine: SpreadsheetEngine,
  mutate: (engine: SpreadsheetEngine) => void,
): WorkbookChangeUndoBundle | null {
  return toEngineUndoBundle(
    engine.captureUndoOps(() => {
      mutate(engine)
    }).undoOps,
  )
}

export async function commitWorkbookMutation(
  documentId: string,
  tx: ServerTransactionLike,
  eventPayload: WorkbookEventPayload,
  runtimeManager: WorkbookRuntimeManager,
  mutate: (engine: SpreadsheetEngine) => WorkbookChangeUndoBundle | null,
  clientMutationId?: string,
  session?: SessionIdentity,
  updatedBy = session?.userID ?? 'system',
) {
  return await runtimeManager.runExclusive(documentId, async () => {
    const db = tx.dbTransaction.wrappedTransaction
    await acquireWorkbookMutationLock(db, documentId)
    const appliedClientMutation = await loadAppliedWorkbookClientMutation(db, documentId, clientMutationId)
    if (appliedClientMutation) {
      assertClientMutationReplayMatches(appliedClientMutation, eventPayload)
      return {
        documentId,
        revision: appliedClientMutation.revision,
        updatedAt: appliedClientMutation.createdAt,
      }
    }
    const state = await runtimeManager.loadRuntime(db, documentId)
    try {
      const undoBundle = mutate(state.engine)
      const ownerUserId = resolveOwnerUserId(state, session)
      const result = await persistWorkbookMutation(db, documentId, {
        previousState: state,
        nextEngine: state.engine,
        updatedBy,
        ownerUserId,
        eventPayload,
        undoBundle,
        ...(clientMutationId !== undefined ? { clientMutationId } : {}),
      })
      runtimeManager.commitMutation(documentId, {
        projectionCommit: result.projectionCommit,
        headRevision: result.revision,
        calculatedRevision: result.calculatedRevision,
        ownerUserId,
      })
      return {
        documentId,
        revision: result.revision,
        updatedAt: result.updatedAt,
      }
    } catch (error) {
      runtimeManager.invalidate(documentId)
      throw error
    }
  })
}

export async function commitWorkbookHistoryMutation(input: {
  documentId: string
  serverTx: ServerTransactionLike
  runtimeManager: WorkbookRuntimeManager
  session?: SessionIdentity
  clientMutationId?: string
  targetChange: {
    revision: number
    summary: string
    sheetName: string | null
    anchorAddress: string | null
    range: WorkbookChangeRange | null
    undoBundle: WorkbookChangeUndoBundle
  }
  eventKind: 'revertChange' | 'redoChange'
}): Promise<void> {
  const { clientMutationId, documentId, eventKind, runtimeManager, serverTx, session, targetChange } = input
  await runtimeManager.runExclusive(documentId, async () => {
    const db = serverTx.dbTransaction.wrappedTransaction
    await acquireWorkbookMutationLock(db, documentId)
    const eventPayload: WorkbookEventPayload = {
      kind: eventKind,
      targetRevision: targetChange.revision,
      targetSummary: targetChange.summary,
      ...(targetChange.sheetName ? { sheetName: targetChange.sheetName } : {}),
      ...(targetChange.anchorAddress ? { address: targetChange.anchorAddress } : {}),
      ...(targetChange.range ? { range: targetChange.range } : {}),
      appliedBundle: targetChange.undoBundle,
    }
    const appliedClientMutation = await loadAppliedWorkbookClientMutation(db, documentId, clientMutationId)
    if (appliedClientMutation) {
      assertClientMutationReplayMatches(appliedClientMutation, eventPayload)
      return
    }
    const state = await runtimeManager.loadRuntime(db, documentId)
    try {
      const undoBundle = applyWorkbookChangeUndoBundle(state.engine, targetChange.undoBundle)
      const ownerUserId = resolveOwnerUserId(state, session)
      const result = await persistWorkbookMutation(db, documentId, {
        previousState: state,
        nextEngine: state.engine,
        updatedBy: session?.userID ?? 'system',
        ownerUserId,
        eventPayload,
        undoBundle,
        ...(clientMutationId !== undefined ? { clientMutationId } : {}),
      })
      runtimeManager.commitMutation(documentId, {
        projectionCommit: result.projectionCommit,
        headRevision: result.revision,
        calculatedRevision: result.calculatedRevision,
        ownerUserId,
      })
    } catch (error) {
      runtimeManager.invalidate(documentId)
      throw error
    }
  })
}

export async function acknowledgeAppliedWorkbookHistoryMutationReplay(input: {
  documentId: string
  serverTx: ServerTransactionLike
  runtimeManager: WorkbookRuntimeManager
  clientMutationId: string | undefined
  matches: (payload: WorkbookEventPayload) => boolean
}): Promise<boolean> {
  if (input.clientMutationId === undefined) {
    return false
  }
  return await input.runtimeManager.runExclusive(input.documentId, async () => {
    const db = input.serverTx.dbTransaction.wrappedTransaction
    await acquireWorkbookMutationLock(db, input.documentId)
    const appliedClientMutation = await loadAppliedWorkbookClientMutation(db, input.documentId, input.clientMutationId)
    if (!appliedClientMutation) {
      return false
    }
    if (!input.matches(appliedClientMutation.payload)) {
      throw new Error(
        `Client mutation ${appliedClientMutation.clientMutationId} for workbook ${appliedClientMutation.documentId} was already applied with a different payload`,
      )
    }
    return true
  })
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isQueryable(value: unknown): value is Queryable {
  return isRecord(value) && typeof value['query'] === 'function'
}

function resolveOwnerUserId(state: { ownerUserId: string }, session?: SessionIdentity): string {
  if (state.ownerUserId !== 'system' || !session?.userID) {
    return state.ownerUserId
  }
  return session.userID
}

function assertClientMutationReplayMatches(appliedMutation: AppliedWorkbookClientMutation, eventPayload: WorkbookEventPayload): void {
  if (isDeepStrictEqual(appliedMutation.payload, eventPayload)) {
    return
  }
  throw new Error(
    `Client mutation ${appliedMutation.clientMutationId} for workbook ${appliedMutation.documentId} was already applied with a different payload`,
  )
}

function applyWorkbookChangeUndoBundle(engine: SpreadsheetEngine, undoBundle: WorkbookChangeUndoBundle): WorkbookChangeUndoBundle | null {
  switch (undoBundle.kind) {
    case 'engineOps':
      return toEngineUndoBundle(engine.applyOps(undoBundle.ops, { captureUndo: true }))
    case 'snapshot': {
      const redoSnapshot = engine.exportSnapshot()
      engine.importSnapshot(undoBundle.snapshot)
      return {
        kind: 'snapshot',
        snapshot: redoSnapshot,
      }
    }
    default: {
      const exhaustive: never = undoBundle
      return exhaustive
    }
  }
}
