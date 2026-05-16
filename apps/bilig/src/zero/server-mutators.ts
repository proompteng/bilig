import { applyWorkbookAgentCommandBundle } from '@bilig/agent-api'
import {
  applyAgentCommandBundleArgsSchema,
  clearRangeArgsSchema,
  clearRangeNumberFormatArgsSchema,
  clearRangeStyleArgsSchema,
  clearCellArgsSchema,
  mergeCellsArgsSchema,
  parseApplyBatchArgs,
  parseRenderCommitArgs,
  redoLatestWorkbookChangeArgsSchema,
  rangeMutationArgsSchema,
  revertWorkbookChangeArgsSchema,
  setCellFormulaArgsSchema,
  setCellValueArgsSchema,
  setFreezePaneArgsSchema,
  setRangeNumberFormatArgsSchema,
  setRangeStyleArgsSchema,
  structuralAxisMutationArgsSchema,
  unmergeCellsArgsSchema,
  undoLatestWorkbookChangeArgsSchema,
  updateColumnMetadataArgsSchema,
  updatePresenceArgsSchema,
  updateColumnWidthArgsSchema,
  updateRowMetadataArgsSchema,
} from '@bilig/zero-sync'
import type { SessionIdentity } from '../http/session.js'
import type { WorkbookRuntimeManager } from '../workbook-runtime/runtime-manager.js'
import { ensureWorkbookDocumentExists } from './workbook-migration-store.js'
import { upsertWorkbookPresence } from './presence-store.js'
import { loadLatestRedoableWorkbookChange, loadLatestUndoableWorkbookChange, loadWorkbookChange } from './workbook-change-store.js'
import { normalizeNumberFormatInput, normalizeStylePatch } from './server-mutator-format-payloads.js'
import {
  acknowledgeAppliedWorkbookHistoryMutationReplay,
  captureEngineUndoBundle,
  commitWorkbookHistoryMutation,
  commitWorkbookMutation,
  requireServerTransaction,
  toEngineUndoBundle,
} from './server-mutator-commit.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export async function handleServerMutator(
  tx: unknown,
  name: string,
  args: unknown,
  runtimeManager: WorkbookRuntimeManager,
  session?: SessionIdentity,
): Promise<void> {
  const serverTx = requireServerTransaction(tx)

  switch (name) {
    case 'workbook.applyBatch': {
      const parsed = parseApplyBatchArgs(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'applyBatch',
          batch: parsed.batch,
        },
        runtimeManager,
        (engine) => {
          return toEngineUndoBundle(engine.applyOps(parsed.batch.ops, { captureUndo: true }))
        },
        parsed.clientMutationId,
        session,
        session?.userID ?? (isRecord(parsed.batch) && typeof parsed.batch['replicaId'] === 'string' ? parsed.batch['replicaId'] : 'system'),
      )
      return
    }

    case 'workbook.applyAgentCommandBundle': {
      const parsed = applyAgentCommandBundleArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'applyAgentCommandBundle',
          bundle: parsed.bundle,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            applyWorkbookAgentCommandBundle(draft, parsed.bundle)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.setCellValue': {
      const parsed = setCellValueArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'setCellValue',
          sheetName: parsed.sheetName,
          address: parsed.address,
          value: parsed.value,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setCellValue(parsed.sheetName, parsed.address, parsed.value)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.setCellFormula': {
      const parsed = setCellFormulaArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'setCellFormula',
          sheetName: parsed.sheetName,
          address: parsed.address,
          formula: parsed.formula,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setCellFormula(parsed.sheetName, parsed.address, parsed.formula)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.clearCell': {
      const parsed = clearCellArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'clearCell',
          sheetName: parsed.sheetName,
          address: parsed.address,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.clearCell(parsed.sheetName, parsed.address)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.clearRange': {
      const parsed = clearRangeArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'clearRange',
          range: parsed.range,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.clearRange(parsed.range)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.renderCommit': {
      const parsed = parseRenderCommitArgs(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'renderCommit',
          ops: parsed.ops,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.renderCommit(parsed.ops)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.fillRange': {
      const parsed = rangeMutationArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'fillRange',
          source: parsed.source,
          target: parsed.target,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.fillRange(parsed.source, parsed.target)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.copyRange': {
      const parsed = rangeMutationArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'copyRange',
          source: parsed.source,
          target: parsed.target,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.copyRange(parsed.source, parsed.target)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.moveRange': {
      const parsed = rangeMutationArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'moveRange',
          source: parsed.source,
          target: parsed.target,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.moveRange(parsed.source, parsed.target)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.updateRowMetadata': {
      const parsed = updateRowMetadataArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'updateRowMetadata',
          sheetName: parsed.sheetName,
          startRow: parsed.startRow,
          count: parsed.count,
          height: parsed.height,
          hidden: parsed.hidden,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.updateRowMetadata(parsed.sheetName, parsed.startRow, parsed.count, parsed.height, parsed.hidden)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.insertRows':
    case 'workbook.deleteRows':
    case 'workbook.insertColumns':
    case 'workbook.deleteColumns': {
      const parsed = structuralAxisMutationArgsSchema.parse(args)
      const kind =
        name === 'workbook.insertRows'
          ? 'insertRows'
          : name === 'workbook.deleteRows'
            ? 'deleteRows'
            : name === 'workbook.insertColumns'
              ? 'insertColumns'
              : 'deleteColumns'
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind,
          sheetName: parsed.sheetName,
          start: parsed.start,
          count: parsed.count,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            switch (kind) {
              case 'insertRows':
                draft.insertRows(parsed.sheetName, parsed.start, parsed.count)
                break
              case 'deleteRows':
                draft.deleteRows(parsed.sheetName, parsed.start, parsed.count)
                break
              case 'insertColumns':
                draft.insertColumns(parsed.sheetName, parsed.start, parsed.count)
                break
              case 'deleteColumns':
                draft.deleteColumns(parsed.sheetName, parsed.start, parsed.count)
                break
            }
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.updateColumnMetadata': {
      const parsed = updateColumnMetadataArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'updateColumnMetadata',
          sheetName: parsed.sheetName,
          startCol: parsed.startCol,
          count: parsed.count,
          width: parsed.width,
          hidden: parsed.hidden,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.updateColumnMetadata(parsed.sheetName, parsed.startCol, parsed.count, parsed.width, parsed.hidden)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.updateColumnWidth': {
      const parsed = updateColumnWidthArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'updateColumnWidth',
          sheetName: parsed.sheetName,
          columnIndex: parsed.columnIndex,
          width: parsed.width,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.updateColumnMetadata(parsed.sheetName, parsed.columnIndex, 1, parsed.width, null)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.setFreezePane': {
      const parsed = setFreezePaneArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'setFreezePane',
          sheetName: parsed.sheetName,
          rows: parsed.rows,
          cols: parsed.cols,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setFreezePane(parsed.sheetName, parsed.rows, parsed.cols)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.mergeCells': {
      const parsed = mergeCellsArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        { kind: 'mergeCells', range: parsed.range },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.mergeCells(parsed.range)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.unmergeCells': {
      const parsed = unmergeCellsArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        { kind: 'unmergeCells', range: parsed.range },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.unmergeCells(parsed.range)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.setRangeStyle': {
      const parsed = setRangeStyleArgsSchema.parse(args)
      const patch = normalizeStylePatch(parsed.patch)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'setRangeStyle',
          range: parsed.range,
          patch,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setRangeStyle(parsed.range, patch)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.clearRangeStyle': {
      const parsed = clearRangeStyleArgsSchema.parse(args)
      const eventPayload =
        parsed.fields === undefined
          ? {
              kind: 'clearRangeStyle' as const,
              range: parsed.range,
            }
          : {
              kind: 'clearRangeStyle' as const,
              range: parsed.range,
              fields: parsed.fields,
            }
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        eventPayload,
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.clearRangeStyle(parsed.range, parsed.fields)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.setRangeNumberFormat': {
      const parsed = setRangeNumberFormatArgsSchema.parse(args)
      const format = normalizeNumberFormatInput(parsed.format)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'setRangeNumberFormat',
          range: parsed.range,
          format,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setRangeNumberFormat(parsed.range, format)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.clearRangeNumberFormat': {
      const parsed = clearRangeNumberFormatArgsSchema.parse(args)
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: 'clearRangeNumberFormat',
          range: parsed.range,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.clearRangeNumberFormat(parsed.range)
          })
        },
        parsed.clientMutationId,
        session,
      )
      return
    }

    case 'workbook.updatePresence': {
      const parsed = updatePresenceArgsSchema.parse(args)
      await ensureWorkbookDocumentExists(serverTx.dbTransaction.wrappedTransaction, parsed.documentId, session?.userID ?? 'system')
      await upsertWorkbookPresence(serverTx.dbTransaction.wrappedTransaction, {
        documentId: parsed.documentId,
        sessionId: parsed.sessionId,
        userId: session?.userID ?? 'system',
        presenceClientId: parsed.presenceClientId ?? null,
        sheetId: parsed.sheetId ?? null,
        sheetName: parsed.sheetName ?? null,
        address: parsed.address ?? null,
        selection: parsed.selection,
      })
      return
    }

    case 'workbook.revertChange': {
      const parsed = revertWorkbookChangeArgsSchema.parse(args)
      const db = serverTx.dbTransaction.wrappedTransaction
      if (
        await acknowledgeAppliedWorkbookHistoryMutationReplay({
          documentId: parsed.documentId,
          serverTx,
          runtimeManager,
          clientMutationId: parsed.clientMutationId,
          matches: (payload) => payload.kind === 'revertChange' && payload.targetRevision === parsed.revision,
        })
      ) {
        return
      }
      const targetChange = await loadWorkbookChange(db, parsed.documentId, parsed.revision)
      if (!targetChange) {
        throw new Error('Workbook change was not found')
      }
      if (!targetChange.undoBundle) {
        throw new Error('Workbook change is not revertible')
      }
      if (targetChange.revertedByRevision !== null) {
        throw new Error(`Workbook change was already reverted in r${targetChange.revertedByRevision}`)
      }
      if (targetChange.eventKind === 'revertChange' || targetChange.revertsRevision !== null) {
        throw new Error('Reverting a revert change is not supported')
      }
      await commitWorkbookHistoryMutation({
        documentId: parsed.documentId,
        serverTx,
        runtimeManager,
        ...(session ? { session } : {}),
        ...(parsed.clientMutationId !== undefined ? { clientMutationId: parsed.clientMutationId } : {}),
        targetChange: {
          revision: targetChange.revision,
          summary: targetChange.summary,
          sheetName: targetChange.sheetName,
          anchorAddress: targetChange.anchorAddress,
          range: targetChange.range,
          undoBundle: targetChange.undoBundle,
        },
        eventKind: 'revertChange',
      })
      return
    }

    case 'workbook.undoLatestChange': {
      const parsed = undoLatestWorkbookChangeArgsSchema.parse(args)
      if (
        await acknowledgeAppliedWorkbookHistoryMutationReplay({
          documentId: parsed.documentId,
          serverTx,
          runtimeManager,
          clientMutationId: parsed.clientMutationId,
          matches: (payload) => payload.kind === 'revertChange',
        })
      ) {
        return
      }
      const targetChange = await loadLatestUndoableWorkbookChange(serverTx.dbTransaction.wrappedTransaction, {
        documentId: parsed.documentId,
        actorUserId: session?.userID ?? 'system',
      })
      if (!targetChange?.undoBundle) {
        throw new Error('No undoable workbook change was found')
      }
      await commitWorkbookHistoryMutation({
        documentId: parsed.documentId,
        serverTx,
        runtimeManager,
        ...(session ? { session } : {}),
        ...(parsed.clientMutationId !== undefined ? { clientMutationId: parsed.clientMutationId } : {}),
        targetChange: {
          revision: targetChange.revision,
          summary: targetChange.summary,
          sheetName: targetChange.sheetName,
          anchorAddress: targetChange.anchorAddress,
          range: targetChange.range,
          undoBundle: targetChange.undoBundle,
        },
        eventKind: 'revertChange',
      })
      return
    }

    case 'workbook.redoLatestChange': {
      const parsed = redoLatestWorkbookChangeArgsSchema.parse(args)
      if (
        await acknowledgeAppliedWorkbookHistoryMutationReplay({
          documentId: parsed.documentId,
          serverTx,
          runtimeManager,
          clientMutationId: parsed.clientMutationId,
          matches: (payload) => payload.kind === 'redoChange',
        })
      ) {
        return
      }
      const targetChange = await loadLatestRedoableWorkbookChange(serverTx.dbTransaction.wrappedTransaction, {
        documentId: parsed.documentId,
        actorUserId: session?.userID ?? 'system',
      })
      if (!targetChange?.undoBundle) {
        throw new Error('No redoable workbook change was found')
      }
      await commitWorkbookHistoryMutation({
        documentId: parsed.documentId,
        serverTx,
        runtimeManager,
        ...(session ? { session } : {}),
        ...(parsed.clientMutationId !== undefined ? { clientMutationId: parsed.clientMutationId } : {}),
        targetChange: {
          revision: targetChange.revision,
          summary: targetChange.summary,
          sheetName: targetChange.sheetName,
          anchorAddress: targetChange.anchorAddress,
          range: targetChange.range,
          undoBundle: targetChange.undoBundle,
        },
        eventKind: 'redoChange',
      })
      return
    }

    default:
      throw new Error(`Unknown Zero mutator: ${name}`)
  }
}
