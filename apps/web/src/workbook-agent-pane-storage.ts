import type { WorkbookAgentThreadScope } from '@bilig/contracts'
import { logDebug } from './runtime-logger.js'

const STORAGE_KEY_PREFIX = 'bilig:workbook-agent:'
const DRAFT_STORAGE_KEY_PREFIX = 'bilig:workbook-agent-drafts:'

export interface WorkbookAgentPaneStorageScope {
  readonly documentId: string
  readonly userId: string
}

export interface StoredWorkbookAgentThreadRef {
  threadId: string
}

function storageKey(scope: WorkbookAgentPaneStorageScope): string {
  return `${STORAGE_KEY_PREFIX}${encodeURIComponent(scope.documentId)}:${encodeURIComponent(scope.userId)}`
}

function draftStorageKey(scope: WorkbookAgentPaneStorageScope): string {
  return `${DRAFT_STORAGE_KEY_PREFIX}${encodeURIComponent(scope.documentId)}:${encodeURIComponent(scope.userId)}`
}

function legacyStorageKey(documentId: string): string {
  return `${STORAGE_KEY_PREFIX}${documentId}`
}

function legacyDraftStorageKey(documentId: string): string {
  return `${DRAFT_STORAGE_KEY_PREFIX}${documentId}`
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function getSessionStorage(documentId: string): Storage | null {
  try {
    return window.sessionStorage
  } catch (error) {
    logDebug('Failed to access workbook agent storage', { documentId, error })
    return null
  }
}

function readSessionStorageItem(documentId: string, key: string): string | null {
  const storage = getSessionStorage(documentId)
  if (!storage) {
    return null
  }
  try {
    return storage.getItem(key)
  } catch (error) {
    logDebug('Failed to read workbook agent storage', { documentId, key, error })
    return null
  }
}

function writeSessionStorageItem(documentId: string, key: string, value: string): void {
  const storage = getSessionStorage(documentId)
  if (!storage) {
    return
  }
  try {
    storage.setItem(key, value)
  } catch (error) {
    logDebug('Failed to persist workbook agent storage', { documentId, key, error })
  }
}

function removeSessionStorageItem(documentId: string, key: string): void {
  const storage = getSessionStorage(documentId)
  if (!storage) {
    return
  }
  try {
    storage.removeItem(key)
  } catch (error) {
    logDebug('Failed to clear workbook agent storage', { documentId, key, error })
  }
}

function parseStoredWorkbookAgentSession(value: unknown): StoredWorkbookAgentThreadRef | null {
  if (!isRecord(value) || typeof value['threadId'] !== 'string') {
    return null
  }
  const threadId = value['threadId'].trim()
  return threadId.length > 0 ? { threadId } : null
}

export function loadStoredSession(scope: WorkbookAgentPaneStorageScope): StoredWorkbookAgentThreadRef | null {
  const key = storageKey(scope)
  try {
    removeSessionStorageItem(scope.documentId, legacyStorageKey(scope.documentId))
    const raw = readSessionStorageItem(scope.documentId, key)
    if (!raw) {
      return null
    }
    const parsed = JSON.parse(raw) as unknown
    const storedSession = parseStoredWorkbookAgentSession(parsed)
    if (storedSession) {
      return storedSession
    }
    removeSessionStorageItem(scope.documentId, key)
  } catch (error) {
    logDebug('Failed to load stored workbook agent session', { documentId: scope.documentId, error })
    removeSessionStorageItem(scope.documentId, key)
  }
  return null
}

export function persistStoredSession(scope: WorkbookAgentPaneStorageScope, value: StoredWorkbookAgentThreadRef): void {
  removeSessionStorageItem(scope.documentId, legacyStorageKey(scope.documentId))
  const storedSession = parseStoredWorkbookAgentSession(value)
  if (!storedSession) {
    removeSessionStorageItem(scope.documentId, storageKey(scope))
    return
  }
  writeSessionStorageItem(scope.documentId, storageKey(scope), JSON.stringify(storedSession))
}

export function clearStoredSession(scope: WorkbookAgentPaneStorageScope): void {
  removeSessionStorageItem(scope.documentId, storageKey(scope))
  removeSessionStorageItem(scope.documentId, legacyStorageKey(scope.documentId))
}

export function loadStoredDrafts(scope: WorkbookAgentPaneStorageScope): Record<string, string> {
  const key = draftStorageKey(scope)
  try {
    removeSessionStorageItem(scope.documentId, legacyDraftStorageKey(scope.documentId))
    const raw = readSessionStorageItem(scope.documentId, key)
    if (!raw) {
      return {}
    }
    const parsed = JSON.parse(raw) as unknown
    if (!isRecord(parsed)) {
      removeSessionStorageItem(scope.documentId, key)
      return {}
    }
    const drafts = Object.fromEntries(
      Object.entries(parsed).flatMap(([storedDraftKey, value]) => (typeof value === 'string' ? ([[storedDraftKey, value]] as const) : [])),
    )
    if (Object.keys(drafts).length !== Object.keys(parsed).length) {
      persistStoredDrafts(scope, drafts)
    }
    return drafts
  } catch (error) {
    logDebug('Failed to load stored workbook agent draft', { documentId: scope.documentId, error })
    removeSessionStorageItem(scope.documentId, key)
    return {}
  }
}

export function persistStoredDrafts(scope: WorkbookAgentPaneStorageScope, drafts: Record<string, string>): void {
  removeSessionStorageItem(scope.documentId, legacyDraftStorageKey(scope.documentId))
  const entries = Object.entries(drafts).filter((entry) => entry[1].length > 0)
  if (entries.length === 0) {
    removeSessionStorageItem(scope.documentId, draftStorageKey(scope))
    return
  }
  writeSessionStorageItem(scope.documentId, draftStorageKey(scope), JSON.stringify(Object.fromEntries(entries)))
}

export function clearStoredDraft(scope: WorkbookAgentPaneStorageScope, key: string): void {
  const drafts = loadStoredDrafts(scope)
  if (!(key in drafts)) {
    return
  }
  delete drafts[key]
  persistStoredDrafts(scope, drafts)
}

export function draftKey(threadId: string | null, scope: WorkbookAgentThreadScope): string {
  return threadId ? `thread:${threadId}` : `new:${scope}`
}
