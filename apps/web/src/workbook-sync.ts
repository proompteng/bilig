import type { Zero } from '@rocicorp/zero'
import { createRenderCommitArgs, mutators } from '@bilig/zero-sync'
import {
  isCellNumberFormatInputValue,
  isCellRangeRef,
  isCellStyleFieldList,
  isCellStylePatchValue,
  isCommitOps,
  isLiteralInput,
  isPendingWorkbookMutationInput,
  isWorkbookSheetName,
  isWorkbookStructuralCount,
  isWorkbookStructuralIndex,
  isWorkbookStructuralSize,
} from './workbook-mutation-guards.js'

export {
  isCellNumberFormatInputValue,
  isCellRangeRef,
  isCellStyleFieldList,
  isCellStylePatchValue,
  isCommitOps,
  isLiteralInput,
  isPendingWorkbookMutation,
  isPendingWorkbookMutationInput,
  isPendingWorkbookMutationList,
  isWorkbookMutationMethod,
  isWorkbookSheetName,
  isWorkbookStructuralCount,
  isWorkbookStructuralIndex,
  isWorkbookStructuralSize,
} from './workbook-mutation-guards.js'

export type WorkbookMutationMethod =
  | 'setCellValue'
  | 'setCellFormula'
  | 'clearCell'
  | 'clearRange'
  | 'renderCommit'
  | 'fillRange'
  | 'copyRange'
  | 'moveRange'
  | 'insertRows'
  | 'deleteRows'
  | 'insertColumns'
  | 'deleteColumns'
  | 'updateRowMetadata'
  | 'updateColumnMetadata'
  | 'setFreezePane'
  | 'mergeCells'
  | 'unmergeCells'
  | 'setRangeStyle'
  | 'clearRangeStyle'
  | 'setRangeNumberFormat'
  | 'clearRangeNumberFormat'

type KnownPendingWorkbookMutationMethod = WorkbookMutationMethod | 'updateColumnWidth'

export interface PendingWorkbookMutationInput {
  readonly method: KnownPendingWorkbookMutationMethod
  readonly args: unknown[]
}

export interface PendingWorkbookMutation extends PendingWorkbookMutationInput {
  readonly id: string
  readonly localSeq: number
  readonly baseRevision: number
  readonly enqueuedAtUnixMs: number
  readonly submittedAtUnixMs: number | null
  readonly lastAttemptedAtUnixMs: number | null
  readonly ackedAtUnixMs: number | null
  readonly rebasedAtUnixMs: number | null
  readonly failedAtUnixMs: number | null
  readonly attemptCount: number
  readonly failureMessage: string | null
  readonly status: 'local' | 'submitted' | 'acked' | 'rebased' | 'failed'
}

export function buildZeroWorkbookMutation(
  documentId: string,
  mutation: PendingWorkbookMutationInput | PendingWorkbookMutation,
): Parameters<Zero['mutate']>[0] {
  const mutationMethod = mutation.method
  if (!isPendingWorkbookMutationInput(mutation)) {
    throw new Error(`Invalid ${mutationMethod} args`)
  }
  const { method, args } = mutation
  const clientMutationId = 'id' in mutation ? mutation.id : undefined
  switch (method) {
    case 'setCellValue': {
      const [sheetName, address, value] = args
      if (typeof sheetName !== 'string' || typeof address !== 'string' || !isLiteralInput(value)) {
        throw new Error('Invalid setCellValue args')
      }
      return mutators.workbook.setCellValue({
        documentId,
        clientMutationId,
        sheetName,
        address,
        value,
      })
    }
    case 'setCellFormula': {
      const [sheetName, address, formula] = args
      if (typeof sheetName !== 'string' || typeof address !== 'string' || typeof formula !== 'string') {
        throw new Error('Invalid setCellFormula args')
      }
      return mutators.workbook.setCellFormula({
        documentId,
        clientMutationId,
        sheetName,
        address,
        formula,
      })
    }
    case 'clearCell': {
      const [sheetName, address] = args
      if (typeof sheetName !== 'string' || typeof address !== 'string') {
        throw new Error('Invalid clearCell args')
      }
      return mutators.workbook.clearCell({ documentId, clientMutationId, sheetName, address })
    }
    case 'clearRange': {
      const [range] = args
      if (!isCellRangeRef(range)) {
        throw new Error('Invalid clearRange args')
      }
      return mutators.workbook.clearRange({ documentId, clientMutationId, range })
    }
    case 'renderCommit': {
      const [ops] = args
      if (!isCommitOps(ops)) {
        throw new Error('Invalid renderCommit args')
      }
      return mutators.workbook.renderCommit(createRenderCommitArgs({ documentId, clientMutationId, ops }))
    }
    case 'fillRange':
    case 'copyRange':
    case 'moveRange': {
      const [source, target] = args
      if (!isCellRangeRef(source) || !isCellRangeRef(target)) {
        throw new Error(`Invalid ${method} args`)
      }
      return mutators.workbook[method]({ documentId, clientMutationId, source, target })
    }
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns': {
      const [sheetName, start, count] = args
      if (!isWorkbookSheetName(sheetName) || !isWorkbookStructuralIndex(start) || !isWorkbookStructuralCount(count)) {
        throw new Error(`Invalid ${method} args`)
      }
      if (method === 'insertRows') {
        return mutators.workbook.insertRows({
          documentId,
          clientMutationId,
          sheetName,
          start,
          count,
        })
      }
      if (method === 'deleteRows') {
        return mutators.workbook.deleteRows({
          documentId,
          clientMutationId,
          sheetName,
          start,
          count,
        })
      }
      if (method === 'insertColumns') {
        return mutators.workbook.insertColumns({
          documentId,
          clientMutationId,
          sheetName,
          start,
          count,
        })
      }
      return mutators.workbook.deleteColumns({
        documentId,
        clientMutationId,
        sheetName,
        start,
        count,
      })
    }
    case 'updateRowMetadata': {
      const [sheetName, startRow, count, height, hidden] = args
      if (
        !isWorkbookSheetName(sheetName) ||
        !isWorkbookStructuralIndex(startRow) ||
        !isWorkbookStructuralCount(count) ||
        (height !== null && !isWorkbookStructuralSize(height)) ||
        (hidden !== null && typeof hidden !== 'boolean')
      ) {
        throw new Error('Invalid updateRowMetadata args')
      }
      return mutators.workbook.updateRowMetadata({
        documentId,
        clientMutationId,
        sheetName,
        startRow,
        count,
        height,
        hidden,
      })
    }
    case 'updateColumnMetadata': {
      const [sheetName, startCol, count, width, hidden] = args
      if (
        !isWorkbookSheetName(sheetName) ||
        !isWorkbookStructuralIndex(startCol) ||
        !isWorkbookStructuralCount(count) ||
        (width !== null && !isWorkbookStructuralSize(width)) ||
        (hidden !== null && typeof hidden !== 'boolean')
      ) {
        throw new Error('Invalid updateColumnMetadata args')
      }
      return mutators.workbook.updateColumnMetadata({
        documentId,
        clientMutationId,
        sheetName,
        startCol,
        count,
        width,
        hidden,
      })
    }
    case 'updateColumnWidth': {
      const [sheetName, columnIndex, width] = args
      if (!isWorkbookSheetName(sheetName) || !isWorkbookStructuralIndex(columnIndex) || !isWorkbookStructuralSize(width)) {
        throw new Error('Invalid updateColumnWidth args')
      }
      return mutators.workbook.updateColumnMetadata({
        documentId,
        clientMutationId,
        sheetName,
        startCol: columnIndex,
        count: 1,
        width,
        hidden: null,
      })
    }
    case 'setFreezePane': {
      const [sheetName, rows, cols] = args
      if (!isWorkbookSheetName(sheetName) || !isWorkbookStructuralIndex(rows) || !isWorkbookStructuralIndex(cols)) {
        throw new Error('Invalid setFreezePane args')
      }
      return mutators.workbook.setFreezePane({
        documentId,
        clientMutationId,
        sheetName,
        rows,
        cols,
      })
    }
    case 'mergeCells': {
      const [range] = args
      if (!isCellRangeRef(range)) {
        throw new Error('Invalid mergeCells args')
      }
      return mutators.workbook.mergeCells({ documentId, clientMutationId, range })
    }
    case 'unmergeCells': {
      const [range] = args
      if (!isCellRangeRef(range)) {
        throw new Error('Invalid unmergeCells args')
      }
      return mutators.workbook.unmergeCells({ documentId, clientMutationId, range })
    }
    case 'setRangeStyle': {
      const [range, patch] = args
      if (!isCellRangeRef(range) || !isCellStylePatchValue(patch)) {
        throw new Error('Invalid setRangeStyle args')
      }
      return mutators.workbook.setRangeStyle({ documentId, clientMutationId, range, patch })
    }
    case 'clearRangeStyle': {
      const [range, fields] = args
      if (!isCellRangeRef(range) || (fields !== undefined && !isCellStyleFieldList(fields))) {
        throw new Error('Invalid clearRangeStyle args')
      }
      return mutators.workbook.clearRangeStyle({ documentId, clientMutationId, range, fields })
    }
    case 'setRangeNumberFormat': {
      const [range, format] = args
      if (!isCellRangeRef(range) || !isCellNumberFormatInputValue(format)) {
        throw new Error('Invalid setRangeNumberFormat args')
      }
      return mutators.workbook.setRangeNumberFormat({
        documentId,
        clientMutationId,
        range,
        format,
      })
    }
    case 'clearRangeNumberFormat': {
      const [range] = args
      if (!isCellRangeRef(range)) {
        throw new Error('Invalid clearRangeNumberFormat args')
      }
      return mutators.workbook.clearRangeNumberFormat({ documentId, clientMutationId, range })
    }
  }
}
