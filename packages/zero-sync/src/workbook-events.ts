import type { CellNumberFormatInput, CellRangeRef, CellStyleField, CellStylePatch, LiteralInput, WorkbookSnapshot } from '@bilig/protocol'
import { isWorkbookSnapshot } from '@bilig/protocol'
import { parseCellAddress } from '@bilig/formula'
import { applyWorkbookAgentCommandBundle, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import type { CommitOp, SpreadsheetEngine } from '@bilig/core'
import { isEngineOps, type EngineOp, type EngineOpBatch } from '@bilig/workbook-domain'
import {
  applyAgentCommandBundleArgsSchema,
  applyBatchArgsSchema,
  clearCellArgsSchema,
  clearRangeArgsSchema,
  clearRangeNumberFormatArgsSchema,
  clearRangeStyleArgsSchema,
  mergeCellsArgsSchema,
  rangeMutationArgsSchema,
  renderCommitArgsSchema,
  setCellFormulaArgsSchema,
  setCellValueArgsSchema,
  setFreezePaneArgsSchema,
  setRangeNumberFormatArgsSchema,
  setRangeStyleArgsSchema,
  structuralAxisMutationArgsSchema,
  unmergeCellsArgsSchema,
  updateColumnMetadataArgsSchema,
  updateColumnWidthArgsSchema,
  updateRowMetadataArgsSchema,
} from './mutators.js'
import { isWorkbookChangeRange, type WorkbookChangeRange } from './workbook-change-range.js'

export type WorkbookChangeUndoBundle =
  | {
      kind: 'engineOps'
      ops: EngineOp[]
    }
  | {
      kind: 'snapshot'
      snapshot: WorkbookSnapshot
    }

export interface DirtyRegion {
  sheetName: string
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
}

export type WorkbookEventPayload =
  | {
      kind: 'applyBatch'
      batch: EngineOpBatch
    }
  | {
      kind: 'applyAgentCommandBundle'
      bundle: WorkbookAgentCommandBundle
    }
  | {
      kind: 'setCellValue'
      sheetName: string
      address: string
      value: LiteralInput
    }
  | {
      kind: 'setCellFormula'
      sheetName: string
      address: string
      formula: string
    }
  | {
      kind: 'clearCell'
      sheetName: string
      address: string
    }
  | {
      kind: 'clearRange'
      range: CellRangeRef
    }
  | {
      kind: 'renderCommit'
      ops: CommitOp[]
    }
  | {
      kind: 'fillRange'
      source: CellRangeRef
      target: CellRangeRef
    }
  | {
      kind: 'copyRange'
      source: CellRangeRef
      target: CellRangeRef
    }
  | {
      kind: 'moveRange'
      source: CellRangeRef
      target: CellRangeRef
    }
  | {
      kind: 'insertRows'
      sheetName: string
      start: number
      count: number
    }
  | {
      kind: 'deleteRows'
      sheetName: string
      start: number
      count: number
    }
  | {
      kind: 'insertColumns'
      sheetName: string
      start: number
      count: number
    }
  | {
      kind: 'deleteColumns'
      sheetName: string
      start: number
      count: number
    }
  | {
      kind: 'updateRowMetadata'
      sheetName: string
      startRow: number
      count: number
      height: number | null
      hidden: boolean | null
    }
  | {
      kind: 'updateColumnMetadata'
      sheetName: string
      startCol: number
      count: number
      width: number | null
      hidden: boolean | null
    }
  | {
      kind: 'updateColumnWidth'
      sheetName: string
      columnIndex: number
      width: number
    }
  | {
      kind: 'setFreezePane'
      sheetName: string
      rows: number
      cols: number
    }
  | {
      kind: 'mergeCells'
      range: CellRangeRef
    }
  | {
      kind: 'unmergeCells'
      range: CellRangeRef
    }
  | {
      kind: 'setRangeStyle'
      range: CellRangeRef
      patch: CellStylePatch
    }
  | {
      kind: 'clearRangeStyle'
      range: CellRangeRef
      fields?: readonly CellStyleField[]
    }
  | {
      kind: 'setRangeNumberFormat'
      range: CellRangeRef
      format: CellNumberFormatInput
    }
  | {
      kind: 'clearRangeNumberFormat'
      range: CellRangeRef
    }
  | {
      kind: 'restoreVersion'
      versionId: string
      versionName: string
      sheetName?: string
      address?: string
      snapshot: WorkbookSnapshot
    }
  | {
      kind: 'revertChange'
      targetRevision: number
      targetSummary: string
      sheetName?: string
      address?: string
      range?: WorkbookChangeRange
      appliedBundle: WorkbookChangeUndoBundle
    }
  | {
      kind: 'redoChange'
      targetRevision: number
      targetSummary: string
      sheetName?: string
      address?: string
      range?: WorkbookChangeRange
      appliedBundle: WorkbookChangeUndoBundle
    }

export type WorkbookEventKind = WorkbookEventPayload['kind']

export const WORKBOOK_EVENT_KINDS = [
  'applyBatch',
  'applyAgentCommandBundle',
  'setCellValue',
  'setCellFormula',
  'clearCell',
  'clearRange',
  'renderCommit',
  'fillRange',
  'copyRange',
  'moveRange',
  'insertRows',
  'deleteRows',
  'insertColumns',
  'deleteColumns',
  'updateRowMetadata',
  'updateColumnMetadata',
  'updateColumnWidth',
  'setFreezePane',
  'mergeCells',
  'unmergeCells',
  'setRangeStyle',
  'clearRangeStyle',
  'setRangeNumberFormat',
  'clearRangeNumberFormat',
  'restoreVersion',
  'revertChange',
  'redoChange',
] as const satisfies readonly WorkbookEventKind[]

const WORKBOOK_EVENT_KIND_SET: ReadonlySet<string> = new Set(WORKBOOK_EVENT_KINDS)

export interface WorkbookEventRecord {
  workbookId: string
  revision: number
  actorUserId: string
  clientMutationId: string | null
  payload: WorkbookEventPayload
  createdAt: string
}

export interface AuthoritativeWorkbookEventRecord {
  revision: number
  clientMutationId: string | null
  payload: WorkbookEventPayload
}

export interface AuthoritativeWorkbookEventBatch {
  afterRevision: number
  headRevision: number
  calculatedRevision: number
  events: readonly AuthoritativeWorkbookEventRecord[]
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isSafePositiveInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0
}

export function isWorkbookChangeUndoBundle(value: unknown): value is WorkbookChangeUndoBundle {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'engineOps':
      return isEngineOps(value['ops'])
    case 'snapshot':
      return isWorkbookSnapshot(value['snapshot'])
    default:
      return false
  }
}

export function isWorkbookEventKind(value: unknown): value is WorkbookEventKind {
  return typeof value === 'string' && WORKBOOK_EVENT_KIND_SET.has(value)
}

interface EventMutationArgsSchema {
  safeParse: (value: unknown) => { success: boolean }
}

function matchesMutationArgsSchema(value: Record<string, unknown>, schema: EventMutationArgsSchema): boolean {
  return schema.safeParse({ ...value, documentId: '__event_payload__' }).success
}

export function isWorkbookEventPayload(value: unknown): value is WorkbookEventPayload {
  if (!isRecord(value) || !isWorkbookEventKind(value['kind'])) {
    return false
  }

  switch (value['kind']) {
    case 'applyBatch':
      return matchesMutationArgsSchema(value, applyBatchArgsSchema)
    case 'applyAgentCommandBundle':
      return matchesMutationArgsSchema(value, applyAgentCommandBundleArgsSchema)
    case 'setCellValue':
      return matchesMutationArgsSchema(value, setCellValueArgsSchema)
    case 'setCellFormula':
      return matchesMutationArgsSchema(value, setCellFormulaArgsSchema)
    case 'clearCell':
      return matchesMutationArgsSchema(value, clearCellArgsSchema)
    case 'clearRange':
      return matchesMutationArgsSchema(value, clearRangeArgsSchema)
    case 'renderCommit':
      return matchesMutationArgsSchema(value, renderCommitArgsSchema)
    case 'fillRange':
    case 'copyRange':
    case 'moveRange':
      return matchesMutationArgsSchema(value, rangeMutationArgsSchema)
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns':
      return matchesMutationArgsSchema(value, structuralAxisMutationArgsSchema)
    case 'updateRowMetadata':
      return matchesMutationArgsSchema(value, updateRowMetadataArgsSchema)
    case 'updateColumnMetadata':
      return matchesMutationArgsSchema(value, updateColumnMetadataArgsSchema)
    case 'updateColumnWidth':
      return matchesMutationArgsSchema(value, updateColumnWidthArgsSchema)
    case 'setFreezePane':
      return matchesMutationArgsSchema(value, setFreezePaneArgsSchema)
    case 'mergeCells':
      return matchesMutationArgsSchema(value, mergeCellsArgsSchema)
    case 'unmergeCells':
      return matchesMutationArgsSchema(value, unmergeCellsArgsSchema)
    case 'setRangeStyle':
      return matchesMutationArgsSchema(value, setRangeStyleArgsSchema)
    case 'clearRangeStyle':
      return matchesMutationArgsSchema(value, clearRangeStyleArgsSchema)
    case 'setRangeNumberFormat':
      return matchesMutationArgsSchema(value, setRangeNumberFormatArgsSchema)
    case 'clearRangeNumberFormat':
      return matchesMutationArgsSchema(value, clearRangeNumberFormatArgsSchema)
    case 'restoreVersion':
      return (
        typeof value['versionId'] === 'string' &&
        typeof value['versionName'] === 'string' &&
        (value['sheetName'] === undefined || typeof value['sheetName'] === 'string') &&
        (value['address'] === undefined || typeof value['address'] === 'string') &&
        isWorkbookSnapshot(value['snapshot'])
      )
    case 'revertChange':
    case 'redoChange':
      return (
        isSafePositiveInteger(value['targetRevision']) &&
        typeof value['targetSummary'] === 'string' &&
        (value['sheetName'] === undefined || typeof value['sheetName'] === 'string') &&
        (value['address'] === undefined || typeof value['address'] === 'string') &&
        (value['range'] === undefined || isWorkbookChangeRange(value['range'])) &&
        isWorkbookChangeUndoBundle(value['appliedBundle'])
      )
    default:
      return false
  }
}

export function isAuthoritativeWorkbookEventRecord(value: unknown): value is AuthoritativeWorkbookEventRecord {
  return (
    isRecord(value) &&
    isSafePositiveInteger(value['revision']) &&
    (typeof value['clientMutationId'] === 'string' || value['clientMutationId'] === null) &&
    isWorkbookEventPayload(value['payload'])
  )
}

export function isAuthoritativeWorkbookEventBatch(value: unknown): value is AuthoritativeWorkbookEventBatch {
  if (
    !isRecord(value) ||
    !isSafeNonNegativeInteger(value['afterRevision']) ||
    !isSafeNonNegativeInteger(value['headRevision']) ||
    !isSafeNonNegativeInteger(value['calculatedRevision']) ||
    !Array.isArray(value['events']) ||
    !value['events'].every((event) => isAuthoritativeWorkbookEventRecord(event))
  ) {
    return false
  }
  return hasContiguousAuthoritativeEventRevisions({
    afterRevision: value['afterRevision'],
    headRevision: value['headRevision'],
    calculatedRevision: value['calculatedRevision'],
    events: value['events'],
  })
}

export function isAuthoritativeWorkbookEventBatchAfterRevision(
  value: unknown,
  afterRevision: unknown,
): value is AuthoritativeWorkbookEventBatch {
  return isSafeNonNegativeInteger(afterRevision) && isAuthoritativeWorkbookEventBatch(value) && value.afterRevision === afterRevision
}

function hasContiguousAuthoritativeEventRevisions(batch: AuthoritativeWorkbookEventBatch): boolean {
  if (batch.afterRevision > batch.headRevision || batch.calculatedRevision > batch.headRevision) {
    return false
  }
  if (batch.events.length === 0) {
    return batch.afterRevision === batch.headRevision
  }
  let expectedRevision = batch.afterRevision + 1
  for (const event of batch.events) {
    if (event.revision !== expectedRevision) {
      return false
    }
    expectedRevision += 1
  }
  return batch.events.at(-1)?.revision === batch.headRevision
}

function singleCellRegion(sheetName: string, address: string): DirtyRegion {
  const parsed = parseCellAddress(address, sheetName)
  return {
    sheetName,
    rowStart: parsed.row,
    rowEnd: parsed.row,
    colStart: parsed.col,
    colEnd: parsed.col,
  }
}

function rangeRegion(range: CellRangeRef): DirtyRegion {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    sheetName: range.sheetName,
    rowStart: Math.min(start.row, end.row),
    rowEnd: Math.max(start.row, end.row),
    colStart: Math.min(start.col, end.col),
    colEnd: Math.max(start.col, end.col),
  }
}

export function deriveDirtyRegions(payload: WorkbookEventPayload): DirtyRegion[] | null {
  switch (payload.kind) {
    case 'setCellValue':
    case 'setCellFormula':
    case 'clearCell':
      return [singleCellRegion(payload.sheetName, payload.address)]
    case 'clearRange':
      return [rangeRegion(payload.range)]
    case 'fillRange':
    case 'copyRange':
    case 'moveRange':
      return [rangeRegion(payload.source), rangeRegion(payload.target)]
    case 'setRangeStyle':
    case 'clearRangeStyle':
    case 'setRangeNumberFormat':
    case 'clearRangeNumberFormat':
    case 'mergeCells':
    case 'unmergeCells':
      return [rangeRegion(payload.range)]
    case 'applyBatch':
    case 'applyAgentCommandBundle':
    case 'renderCommit':
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns':
    case 'updateRowMetadata':
    case 'updateColumnMetadata':
    case 'updateColumnWidth':
    case 'setFreezePane':
    case 'restoreVersion':
    case 'revertChange':
    case 'redoChange':
      return null
    default: {
      const exhaustive: never = payload
      return exhaustive
    }
  }
}

export function applyWorkbookEvent(engine: SpreadsheetEngine, payload: WorkbookEventPayload): void {
  switch (payload.kind) {
    case 'applyBatch':
      engine.applyRemoteBatch(payload.batch)
      return
    case 'applyAgentCommandBundle':
      applyWorkbookAgentCommandBundle(engine, payload.bundle)
      return
    case 'setCellValue':
      engine.setCellValue(payload.sheetName, payload.address, payload.value)
      return
    case 'setCellFormula':
      engine.setCellFormula(payload.sheetName, payload.address, payload.formula)
      return
    case 'clearCell':
      engine.clearCell(payload.sheetName, payload.address)
      return
    case 'clearRange':
      engine.clearRange(payload.range)
      return
    case 'renderCommit':
      engine.renderCommit(payload.ops)
      return
    case 'fillRange':
      engine.fillRange(payload.source, payload.target)
      return
    case 'copyRange':
      engine.copyRange(payload.source, payload.target)
      return
    case 'moveRange':
      engine.moveRange(payload.source, payload.target)
      return
    case 'insertRows':
      engine.insertRows(payload.sheetName, payload.start, payload.count)
      return
    case 'deleteRows':
      engine.deleteRows(payload.sheetName, payload.start, payload.count)
      return
    case 'insertColumns':
      engine.insertColumns(payload.sheetName, payload.start, payload.count)
      return
    case 'deleteColumns':
      engine.deleteColumns(payload.sheetName, payload.start, payload.count)
      return
    case 'updateRowMetadata':
      engine.updateRowMetadata(payload.sheetName, payload.startRow, payload.count, payload.height, payload.hidden)
      return
    case 'updateColumnMetadata':
      engine.updateColumnMetadata(payload.sheetName, payload.startCol, payload.count, payload.width, payload.hidden)
      return
    case 'updateColumnWidth':
      engine.updateColumnMetadata(payload.sheetName, payload.columnIndex, 1, payload.width, null)
      return
    case 'setFreezePane':
      engine.setFreezePane(payload.sheetName, payload.rows, payload.cols)
      return
    case 'mergeCells':
      engine.mergeCells(payload.range)
      return
    case 'unmergeCells':
      engine.unmergeCells(payload.range)
      return
    case 'setRangeStyle':
      engine.setRangeStyle(payload.range, payload.patch)
      return
    case 'clearRangeStyle':
      engine.clearRangeStyle(payload.range, payload.fields)
      return
    case 'setRangeNumberFormat':
      engine.setRangeNumberFormat(payload.range, payload.format)
      return
    case 'clearRangeNumberFormat':
      engine.clearRangeNumberFormat(payload.range)
      return
    case 'restoreVersion':
      engine.importSnapshot(payload.snapshot)
      return
    case 'revertChange':
    case 'redoChange':
      if (payload.appliedBundle.kind === 'engineOps') {
        engine.applyOps(payload.appliedBundle.ops)
        return
      }
      engine.importSnapshot(payload.appliedBundle.snapshot)
      return
    default: {
      const exhaustive: never = payload
      throw new Error(`Unhandled workbook event: ${JSON.stringify(exhaustive)}`)
    }
  }
}
