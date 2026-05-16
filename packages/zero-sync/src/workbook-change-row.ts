import { normalizeWorkbookChangeRange, type WorkbookChangeRange } from './workbook-change-range.js'
import {
  isWorkbookChangeUndoBundle,
  isWorkbookEventKind,
  type WorkbookChangeUndoBundle,
  type WorkbookEventKind,
} from './workbook-events.js'
import type { WorkbookHistoryRangeSource } from './workbook-history-state.js'

export interface WorkbookChangeRowModel {
  readonly revision: number
  readonly actorUserId: string
  readonly clientMutationId: string | null
  readonly eventKind: WorkbookEventKind
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

function parseSafeInteger(value: unknown): number | null {
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

function parseSafeNonNegativeInteger(value: unknown): number | null {
  const parsed = parseSafeInteger(value)
  return parsed !== null && parsed >= 0 ? parsed : null
}

function parseSafePositiveInteger(value: unknown): number | null {
  const parsed = parseSafeInteger(value)
  return parsed !== null && parsed > 0 ? parsed : null
}

export function normalizeWorkbookChangeRowModel(value: unknown): WorkbookChangeRowModel | null {
  if (!isRecord(value)) {
    return null
  }
  const revision = parseSafePositiveInteger(value['revision'])
  const createdAt = parseSafeNonNegativeInteger(value['createdAt'])
  const actorUserId = value['actorUserId']
  const eventKind = value['eventKind']
  const summary = value['summary']
  if (
    revision === null ||
    createdAt === null ||
    typeof actorUserId !== 'string' ||
    !isWorkbookEventKind(eventKind) ||
    typeof summary !== 'string'
  ) {
    return null
  }

  const clientMutationId = value['clientMutationId']
  const sheetId = value['sheetId'] === undefined || value['sheetId'] === null ? null : parseSafePositiveInteger(value['sheetId'])
  if (sheetId === null && value['sheetId'] !== undefined && value['sheetId'] !== null) {
    return null
  }
  const sheetName = value['sheetName']
  const anchorAddress = value['anchorAddress']
  const rawRangeJson = value['rangeJson']
  const rangeJson = normalizeWorkbookChangeRange(rawRangeJson)
  const revertedByRevision =
    value['revertedByRevision'] === undefined || value['revertedByRevision'] === null
      ? null
      : parseSafePositiveInteger(value['revertedByRevision'])
  if (revertedByRevision === null && value['revertedByRevision'] !== undefined && value['revertedByRevision'] !== null) {
    return null
  }
  const revertsRevision =
    value['revertsRevision'] === undefined || value['revertsRevision'] === null ? null : parseSafePositiveInteger(value['revertsRevision'])
  if (revertsRevision === null && value['revertsRevision'] !== undefined && value['revertsRevision'] !== null) {
    return null
  }

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
