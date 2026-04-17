import { Effect } from 'effect'
import type { EngineRuntimeState, TransactionRecord } from '../runtime-state.js'
import { EngineHistoryError } from '../errors.js'

export interface EngineHistoryService {
  readonly undo: () => Effect.Effect<boolean, EngineHistoryError>
  readonly redo: () => Effect.Effect<boolean, EngineHistoryError>
}

export function createEngineHistoryService(args: {
  readonly state: Pick<EngineRuntimeState, 'undoStack' | 'redoStack' | 'getTransactionReplayDepth' | 'setTransactionReplayDepth'>
  readonly executeTransaction: (transaction: TransactionRecord, source: 'undo' | 'redo') => void
}): EngineHistoryService {
  return {
    undo() {
      return Effect.try({
        try: () => {
          const entry = args.state.undoStack.pop()
          if (!entry) {
            return false
          }
          args.state.setTransactionReplayDepth(args.state.getTransactionReplayDepth() + 1)
          try {
            args.executeTransaction(entry.inverse, 'undo')
          } finally {
            args.state.setTransactionReplayDepth(args.state.getTransactionReplayDepth() - 1)
          }
          args.state.redoStack.push(entry)
          return true
        },
        catch: (cause) =>
          new EngineHistoryError({
            message: 'Failed to undo transaction',
            cause,
          }),
      })
    },
    redo() {
      return Effect.try({
        try: () => {
          const entry = args.state.redoStack.pop()
          if (!entry) {
            return false
          }
          args.state.setTransactionReplayDepth(args.state.getTransactionReplayDepth() + 1)
          try {
            args.executeTransaction(entry.forward, 'redo')
          } finally {
            args.state.setTransactionReplayDepth(args.state.getTransactionReplayDepth() - 1)
          }
          args.state.undoStack.push(entry)
          return true
        },
        catch: (cause) =>
          new EngineHistoryError({
            message: 'Failed to redo transaction',
            cause,
          }),
      })
    },
  }
}
