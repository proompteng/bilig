import { normalizeWorkbookChangeRange, type WorkbookChangeRange } from './workbook-change-range.js'
import { isWorkbookChangeUndoBundle, type WorkbookChangeUndoBundle } from './workbook-events.js'
import type { WorkbookHistoryRangeSource } from './workbook-history-state.js'

export interface WorkbookChangeRowModel {
  readonly revision: number
  readonly actorUserId: string
  readonly clientMutationId: string | null
  readonly eventKind: string
  readonly summary: string
  readonly sheetId: number | null
  readonly sheetName: string | null
  readonly anchorAddress: string | null
  readonly rangeJson: WorkbookChangeRange | null
  readonly rangeJsonInvalid: boolean
  readonly undoBundleJson: WorkbookChangeUndoBundle | null
  readonly revertedByRevision: number | null
  readonly revertsRevision: number | null
  readonly createdAt: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function parseInteger(value: unknown): number | null {
  if (typeof value === 'number' && Number.isSafeInteger(value)) {
    return value
  }
  if (typeof value === 'bigint') {
    const parsed = Number(value)
    return Number.isSafeInteger(parsed) ? parsed : null
  }
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (/^-?\d+$/u.test(trimmed)) {
      const parsed = Number(trimmed)
      return Number.isSafeInteger(parsed) ? parsed : null
    }
  }
  return null
}

export function normalizeWorkbookChangeRowModel(value: unknown): WorkbookChangeRowModel | null {
  if (!isRecord(value)) {
    return null
  }
  const revision = parseInteger(value['revision'])
  const createdAt = parseInteger(value['createdAt'])
  const actorUserId = value['actorUserId']
  const eventKind = value['eventKind']
  const summary = value['summary']
  if (
    revision === null ||
    createdAt === null ||
    typeof actorUserId !== 'string' ||
    typeof eventKind !== 'string' ||
    typeof summary !== 'string'
  ) {
    return null
  }

  const clientMutationId = value['clientMutationId']
  const sheetId = parseInteger(value['sheetId'])
  const sheetName = value['sheetName']
  const anchorAddress = value['anchorAddress']
  const rawRangeJson = value['rangeJson']
  const rangeJson = normalizeWorkbookChangeRange(rawRangeJson)
  const revertedByRevision = parseInteger(value['revertedByRevision'])
  const revertsRevision = parseInteger(value['revertsRevision'])

  return {
    revision,
    actorUserId,
    clientMutationId: typeof clientMutationId === 'string' ? clientMutationId : null,
    eventKind,
    summary,
    sheetId,
    sheetName: typeof sheetName === 'string' ? sheetName : null,
    anchorAddress: typeof anchorAddress === 'string' ? anchorAddress : null,
    rangeJson,
    rangeJsonInvalid: rawRangeJson !== undefined && rawRangeJson !== null && rangeJson === null,
    undoBundleJson: isWorkbookChangeUndoBundle(value['undoBundleJson']) ? value['undoBundleJson'] : null,
    revertedByRevision,
    revertsRevision,
    createdAt,
  }
}

export function workbookChangeRowHistoryRangeSource(
  row: Pick<WorkbookChangeRowModel, 'sheetName' | 'anchorAddress' | 'rangeJson' | 'rangeJsonInvalid'>,
): WorkbookHistoryRangeSource {
  return {
    sheetName: row.sheetName,
    anchorAddress: row.anchorAddress,
    rangeJson: row.rangeJson,
    rangeJsonInvalid: row.rangeJsonInvalid,
  }
}
