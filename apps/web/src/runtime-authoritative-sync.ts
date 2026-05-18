import { isAuthoritativeWorkbookEventBatchAfterRevision, type AuthoritativeWorkbookEventBatch } from '@bilig/zero-sync'
import { isWorkbookSnapshot, type WorkbookSnapshot } from '@bilig/protocol'
import type { WorkbookRevisionState } from './runtime-zero-revision-sync.js'

export type { AuthoritativeWorkbookEventBatch } from '@bilig/zero-sync'

export interface LatestWorkbookSnapshot {
  readonly snapshot: WorkbookSnapshot
  readonly revision: number | null
}

export function parseSnapshotRevisionHeader(value: string | null): number | null {
  const trimmed = value?.trim()
  if (!trimmed || !/^(0|[1-9]\d*)$/u.test(trimmed)) {
    return null
  }
  const parsed = Number(trimmed)
  return Number.isSafeInteger(parsed) ? parsed : null
}

async function parseJsonResponse(response: Response, context: string): Promise<unknown> {
  try {
    return JSON.parse(await response.text()) as unknown
  } catch {
    throw new Error(`${context} response returned malformed JSON`)
  }
}

export async function loadAuthoritativeEventBatch(input: {
  readonly documentId: string
  readonly afterRevision: number
  readonly fetchImpl: typeof fetch
}): Promise<AuthoritativeWorkbookEventBatch> {
  const response = await input.fetchImpl(
    `/v2/documents/${encodeURIComponent(input.documentId)}/events?afterRevision=${String(input.afterRevision)}`,
    {
      headers: {
        accept: 'application/json',
      },
      cache: 'no-store',
    },
  )
  if (!response.ok) {
    throw new Error(`Failed to load authoritative events (${response.status})`)
  }
  const parsed = await parseJsonResponse(response, 'Authoritative events')
  if (!isAuthoritativeWorkbookEventBatchAfterRevision(parsed, input.afterRevision)) {
    throw new Error('Authoritative event payload does not match the expected schema')
  }
  return parsed
}

export async function loadLatestWorkbookSnapshot(input: {
  readonly documentId: string
  readonly fetchImpl: typeof fetch
}): Promise<LatestWorkbookSnapshot | null> {
  const response = await input.fetchImpl(`/v2/documents/${encodeURIComponent(input.documentId)}/snapshot/latest`, {
    headers: {
      accept: 'application/json, application/vnd.bilig.workbook+json',
    },
    cache: 'no-store',
  })
  if (response.status === 204 || response.status === 404) {
    return null
  }
  if (!response.ok) {
    throw new Error(`Failed to load workbook snapshot (${response.status})`)
  }
  const parsed = await parseJsonResponse(response, 'Workbook snapshot')
  if (!isWorkbookSnapshot(parsed)) {
    throw new Error('Workbook snapshot payload does not match the expected schema')
  }
  return {
    snapshot: parsed,
    revision: parseSnapshotRevisionHeader(response.headers.get('x-bilig-snapshot-cursor')),
  }
}

export function shouldApplyAuthoritativeEventBatch(input: {
  readonly eventBatch: Pick<AuthoritativeWorkbookEventBatch, 'events' | 'headRevision' | 'calculatedRevision'>
  readonly currentAuthoritativeRevision: number
}): boolean {
  return (
    input.eventBatch.events.length > 0 &&
    input.eventBatch.headRevision > input.currentAuthoritativeRevision &&
    input.eventBatch.calculatedRevision >= input.eventBatch.headRevision
  )
}

export function shouldQueueAuthoritativeRebase(input: {
  readonly revisionState: WorkbookRevisionState | null
  readonly currentAuthoritativeRevision: number
  readonly currentCalculatedRevision: number
}): input is {
  readonly revisionState: WorkbookRevisionState
  readonly currentAuthoritativeRevision: number
  readonly currentCalculatedRevision: number
} {
  return (
    input.revisionState !== null &&
    input.revisionState.calculatedRevision >= input.revisionState.headRevision &&
    input.revisionState.headRevision >= input.currentAuthoritativeRevision &&
    (input.revisionState.headRevision > input.currentAuthoritativeRevision ||
      input.revisionState.calculatedRevision > input.currentCalculatedRevision)
  )
}

export function shouldInstallBootstrapSnapshot(input: {
  readonly snapshotRevision: number | null
  readonly currentAuthoritativeRevision: number
}): boolean {
  return input.snapshotRevision === null || input.snapshotRevision >= input.currentAuthoritativeRevision
}

export function isAuthoritativeSnapshotNewerForRebase(input: {
  readonly snapshotRevision: number
  readonly snapshotCalculatedRevision: number
  readonly currentAuthoritativeRevision: number
  readonly currentCalculatedRevision: number
}): boolean {
  return (
    input.snapshotRevision > input.currentAuthoritativeRevision ||
    (input.snapshotRevision === input.currentAuthoritativeRevision && input.snapshotCalculatedRevision > input.currentCalculatedRevision)
  )
}
