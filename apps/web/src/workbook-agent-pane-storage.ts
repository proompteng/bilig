import type { WorkbookAgentThreadScope } from '@bilig/contracts'

const STORAGE_KEY_PREFIX = 'bilig:workbook-agent:'
const DRAFT_STORAGE_KEY_PREFIX = 'bilig:workbook-agent-drafts:'

export interface StoredWorkbookAgentThreadRef {
  threadId: string
}

function storageKey(documentId: string): string {
  return `${STORAGE_KEY_PREFIX}${documentId}`
}

function draftStorageKey(documentId: string): string {
  return `${DRAFT_STORAGE_KEY_PREFIX}${documentId}`
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isStoredWorkbookAgentSession(value: unknown): value is StoredWorkbookAgentThreadRef {
  return isRecord(value) && typeof value['threadId'] === 'string'
}

export function loadStoredSession(documentId: string): StoredWorkbookAgentThreadRef | null {
  try {
    const raw = window.sessionStorage.getItem(storageKey(documentId))
    if (!raw) {
      return null
    }
    const parsed = JSON.parse(raw) as unknown
    if (isStoredWorkbookAgentSession(parsed)) {
      return parsed
    }
  } catch {}
  return null
}

export function persistStoredSession(documentId: string, value: StoredWorkbookAgentThreadRef): void {
  window.sessionStorage.setItem(storageKey(documentId), JSON.stringify(value))
}

export function clearStoredSession(documentId: string): void {
  window.sessionStorage.removeItem(storageKey(documentId))
}

export function loadStoredDrafts(documentId: string): Record<string, string> {
  try {
    const raw = window.sessionStorage.getItem(draftStorageKey(documentId))
    if (!raw) {
      return {}
    }
    const parsed = JSON.parse(raw) as unknown
    if (!isRecord(parsed)) {
      return {}
    }
    return Object.fromEntries(
      Object.entries(parsed).flatMap(([key, value]) => (typeof value === 'string' ? ([[key, value]] as const) : [])),
    )
  } catch {
    return {}
  }
}

export function persistStoredDrafts(documentId: string, drafts: Record<string, string>): void {
  const entries = Object.entries(drafts).filter((entry) => entry[1].length > 0)
  if (entries.length === 0) {
    window.sessionStorage.removeItem(draftStorageKey(documentId))
    return
  }
  window.sessionStorage.setItem(draftStorageKey(documentId), JSON.stringify(Object.fromEntries(entries)))
}

export function clearStoredDraft(documentId: string, key: string): void {
  const drafts = loadStoredDrafts(documentId)
  if (!(key in drafts)) {
    return
  }
  delete drafts[key]
  persistStoredDrafts(documentId, drafts)
}

export function draftKey(threadId: string | null, scope: WorkbookAgentThreadScope): string {
  return threadId ? `thread:${threadId}` : `new:${scope}`
}
