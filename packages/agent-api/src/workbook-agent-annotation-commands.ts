import type { SpreadsheetEngine } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { WorkbookCommentThreadSnapshot, WorkbookNoteSnapshot } from '@bilig/protocol'
import type { WorkbookAgentCommand, WorkbookAgentPreviewRange } from './workbook-agent-bundles.js'

export type WorkbookAgentAnnotationCommand = Extract<
  WorkbookAgentCommand,
  { kind: 'upsertCommentThread' } | { kind: 'deleteCommentThread' } | { kind: 'upsertNote' } | { kind: 'deleteNote' }
>

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isCommentEntry(value: unknown): value is WorkbookCommentThreadSnapshot['comments'][number] {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    value['id'].trim().length > 0 &&
    typeof value['body'] === 'string' &&
    value['body'].trim().length > 0 &&
    (value['authorUserId'] === undefined || typeof value['authorUserId'] === 'string') &&
    (value['authorDisplayName'] === undefined || typeof value['authorDisplayName'] === 'string') &&
    (value['createdAtUnixMs'] === undefined || typeof value['createdAtUnixMs'] === 'number')
  )
}

function isCommentThread(value: unknown): value is WorkbookCommentThreadSnapshot {
  return (
    isRecord(value) &&
    typeof value['threadId'] === 'string' &&
    value['threadId'].trim().length > 0 &&
    typeof value['sheetName'] === 'string' &&
    typeof value['address'] === 'string' &&
    Array.isArray(value['comments']) &&
    value['comments'].length > 0 &&
    value['comments'].every((entry) => isCommentEntry(entry)) &&
    (value['resolved'] === undefined || typeof value['resolved'] === 'boolean') &&
    (value['resolvedByUserId'] === undefined || typeof value['resolvedByUserId'] === 'string') &&
    (value['resolvedAtUnixMs'] === undefined || typeof value['resolvedAtUnixMs'] === 'number')
  )
}

function isNote(value: unknown): value is WorkbookNoteSnapshot {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['address'] === 'string' &&
    typeof value['text'] === 'string' &&
    value['text'].trim().length > 0
  )
}

function normalizeAddress(sheetName: string, address: string): string {
  const parsed = parseCellAddress(address, sheetName)
  return formatAddress(parsed.row, parsed.col)
}

function normalizeCommentThread(thread: WorkbookCommentThreadSnapshot): WorkbookCommentThreadSnapshot {
  return {
    threadId: thread.threadId.trim(),
    sheetName: thread.sheetName,
    address: normalizeAddress(thread.sheetName, thread.address),
    comments: thread.comments.map((entry) => ({
      id: entry.id.trim(),
      body: entry.body.trim(),
      ...(entry.authorUserId !== undefined ? { authorUserId: entry.authorUserId } : {}),
      ...(entry.authorDisplayName !== undefined ? { authorDisplayName: entry.authorDisplayName } : {}),
      ...(entry.createdAtUnixMs !== undefined ? { createdAtUnixMs: entry.createdAtUnixMs } : {}),
    })),
    ...(thread.resolved !== undefined ? { resolved: thread.resolved } : {}),
    ...(thread.resolvedByUserId !== undefined ? { resolvedByUserId: thread.resolvedByUserId } : {}),
    ...(thread.resolvedAtUnixMs !== undefined ? { resolvedAtUnixMs: thread.resolvedAtUnixMs } : {}),
  }
}

function normalizeNote(note: WorkbookNoteSnapshot): WorkbookNoteSnapshot {
  return {
    sheetName: note.sheetName,
    address: normalizeAddress(note.sheetName, note.address),
    text: note.text.trim(),
  }
}

function rangeForAddress(sheetName: string, address: string): WorkbookAgentPreviewRange {
  const normalizedAddress = normalizeAddress(sheetName, address)
  return {
    sheetName,
    startAddress: normalizedAddress,
    endAddress: normalizedAddress,
    role: 'target',
  }
}

export function isWorkbookAgentAnnotationCommandKind(kind: string): kind is WorkbookAgentAnnotationCommand['kind'] {
  return kind === 'upsertCommentThread' || kind === 'deleteCommentThread' || kind === 'upsertNote' || kind === 'deleteNote'
}

export function isWorkbookAgentAnnotationCommand(command: WorkbookAgentCommand): command is WorkbookAgentAnnotationCommand {
  return isWorkbookAgentAnnotationCommandKind(command.kind)
}

export function isWorkbookAgentAnnotationCommandValue(value: unknown): value is WorkbookAgentAnnotationCommand {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'upsertCommentThread':
      return isCommentThread(value['thread'])
    case 'deleteCommentThread':
      return typeof value['sheetName'] === 'string' && typeof value['address'] === 'string'
    case 'upsertNote':
      return isNote(value['note'])
    case 'deleteNote':
      return typeof value['sheetName'] === 'string' && typeof value['address'] === 'string'
    default:
      return false
  }
}

export function isHighRiskWorkbookAgentAnnotationCommand(_command: WorkbookAgentAnnotationCommand): boolean {
  return false
}

export function isWorkbookScopeAnnotationCommand(_command: WorkbookAgentAnnotationCommand): boolean {
  return false
}

export function describeWorkbookAgentAnnotationCommand(command: WorkbookAgentAnnotationCommand): string {
  switch (command.kind) {
    case 'upsertCommentThread':
      return `Set comment thread on ${command.thread.sheetName}!${normalizeAddress(command.thread.sheetName, command.thread.address)}`
    case 'deleteCommentThread':
      return `Delete comment thread on ${command.sheetName}!${normalizeAddress(command.sheetName, command.address)}`
    case 'upsertNote':
      return `Set note on ${command.note.sheetName}!${normalizeAddress(command.note.sheetName, command.note.address)}`
    case 'deleteNote':
      return `Delete note on ${command.sheetName}!${normalizeAddress(command.sheetName, command.address)}`
  }
}

export function estimateWorkbookAgentAnnotationCommandAffectedCells(_command: WorkbookAgentAnnotationCommand): number {
  return 1
}

export function deriveWorkbookAgentAnnotationCommandPreviewRanges(command: WorkbookAgentAnnotationCommand): WorkbookAgentPreviewRange[] {
  switch (command.kind) {
    case 'upsertCommentThread':
      return [rangeForAddress(command.thread.sheetName, command.thread.address)]
    case 'deleteCommentThread':
      return [rangeForAddress(command.sheetName, command.address)]
    case 'upsertNote':
      return [rangeForAddress(command.note.sheetName, command.note.address)]
    case 'deleteNote':
      return [rangeForAddress(command.sheetName, command.address)]
  }
}

export function applyWorkbookAgentAnnotationCommand(engine: SpreadsheetEngine, command: WorkbookAgentAnnotationCommand): void {
  switch (command.kind) {
    case 'upsertCommentThread':
      engine.setCommentThread(normalizeCommentThread(command.thread))
      return
    case 'deleteCommentThread':
      engine.deleteCommentThread(command.sheetName, command.address)
      return
    case 'upsertNote':
      engine.setNote(normalizeNote(command.note))
      return
    case 'deleteNote':
      engine.deleteNote(command.sheetName, command.address)
      return
  }
}
