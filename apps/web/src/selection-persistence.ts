import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { WorkerRuntimeSelection } from './runtime-session.js'
import {
  legacyWorkbookDocumentStorageKey,
  scopedWorkbookStorageKey,
  type WorkbookBrowserStorageScope,
} from './workbook-browser-storage-scope.js'

const DEFAULT_SELECTION: WorkerRuntimeSelection = {
  sheetName: 'Sheet1',
  address: 'A1',
}
const SHEET_QUERY_PARAM = 'sheet'
const CELL_QUERY_PARAM = 'cell'
const SELECTION_PERSIST_DEBOUNCE_MS = 120
const SELECTION_URL_CHANGE_EVENT = 'bilig-selection-url-change'
const SELECTION_HISTORY_STATE_KEY = '__biligSelectionHistoryInstrumentation'

interface PendingSelectionPersist {
  readonly scope: SelectionPersistenceScope
  readonly selection: WorkerRuntimeSelection
  readonly timeoutId: ReturnType<typeof globalThis.setTimeout>
}

const pendingPersists = new Map<string, PendingSelectionPersist>()
let flushListenersInstalled = false
let suppressNextSelectionUrlChange = false

interface SelectionHistoryInstrumentation {
  readonly history: History
  readonly originalPushState: History['pushState'] | null
  readonly originalReplaceState: History['replaceState']
}

type SelectionHistoryWindow = Window & {
  [SELECTION_HISTORY_STATE_KEY]?: SelectionHistoryInstrumentation | undefined
}

export type SelectionPersistenceScope = WorkbookBrowserStorageScope

const SELECTION_STORAGE_KEY_PREFIX = 'bilig:selection:'

function storageKey(scope: SelectionPersistenceScope): string {
  return scopedWorkbookStorageKey(SELECTION_STORAGE_KEY_PREFIX, scope)
}

function legacyStorageKey(documentId: string): string {
  return legacyWorkbookDocumentStorageKey(SELECTION_STORAGE_KEY_PREFIX, documentId)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeSelection(sheetName: string, address: string): WorkerRuntimeSelection | null {
  const trimmedSheetName = sheetName.trim()
  const trimmedAddress = address.trim().toUpperCase()
  if (trimmedSheetName.length === 0 || trimmedAddress.length === 0) {
    return null
  }
  try {
    const parsed = parseCellAddress(trimmedAddress, trimmedSheetName)
    return {
      sheetName: trimmedSheetName,
      address: formatAddress(parsed.row, parsed.col),
    }
  } catch {
    return null
  }
}

function normalizeSheetName(sheetName: string): string | null {
  const trimmedSheetName = sheetName.trim()
  return trimmedSheetName.length === 0 ? null : trimmedSheetName
}

function readSheetSelectionFromUrl(): string | null {
  if (typeof window === 'undefined') {
    return null
  }
  const searchParams = new URLSearchParams(window.location.search)
  const sheetName = searchParams.get(SHEET_QUERY_PARAM)
  return sheetName ? normalizeSheetName(sheetName) : null
}

function readCellSelectionFromUrl(): string | null {
  if (typeof window === 'undefined') {
    return null
  }
  const searchParams = new URLSearchParams(window.location.search)
  const address = searchParams.get(CELL_QUERY_PARAM)
  if (!address) {
    return null
  }
  try {
    return normalizeSelection('Sheet1', address)?.address ?? null
  } catch {
    return null
  }
}

function removeStoredSelection(key: string): void {
  try {
    window.localStorage.removeItem(key)
  } catch {
    // Ignore storage cleanup failures and keep selection recovery usable.
  }
}

export function readSelectionFromUrl(): WorkerRuntimeSelection | null {
  const sheetName = readSheetSelectionFromUrl()
  if (!sheetName) {
    return null
  }
  return {
    sheetName,
    address: readCellSelectionFromUrl() ?? DEFAULT_SELECTION.address,
  }
}

function readStoredSelection(scope: SelectionPersistenceScope): WorkerRuntimeSelection | null {
  if (typeof window === 'undefined') {
    return null
  }
  const key = storageKey(scope)
  try {
    removeStoredSelection(legacyStorageKey(scope.documentId))
    const raw = window.localStorage.getItem(key)
    if (!raw) {
      return null
    }
    const parsed = JSON.parse(raw)
    if (
      !isRecord(parsed) ||
      typeof parsed['sheetName'] !== 'string' ||
      parsed['sheetName'].trim().length === 0 ||
      typeof parsed['address'] !== 'string' ||
      parsed['address'].trim().length === 0
    ) {
      removeStoredSelection(key)
      return null
    }
    const normalizedSelection = normalizeSelection(parsed['sheetName'], parsed['address'])
    if (!normalizedSelection) {
      removeStoredSelection(key)
    }
    return normalizedSelection
  } catch {
    removeStoredSelection(key)
    return null
  }
}

export function loadPersistedSelection(scope: SelectionPersistenceScope): WorkerRuntimeSelection {
  const storedSelection = readStoredSelection(scope)
  const urlSheetSelection = readSheetSelectionFromUrl()
  const urlCellSelection = readCellSelectionFromUrl()
  if (urlSheetSelection) {
    return {
      sheetName: urlSheetSelection,
      address: urlCellSelection ?? (storedSelection?.sheetName === urlSheetSelection ? storedSelection.address : DEFAULT_SELECTION.address),
    }
  }
  if (typeof window === 'undefined') {
    return DEFAULT_SELECTION
  }
  return storedSelection ?? DEFAULT_SELECTION
}

function emitSelectionUrlChange(): void {
  if (typeof window === 'undefined' || typeof window.dispatchEvent !== 'function') {
    return
  }
  if (suppressNextSelectionUrlChange) {
    suppressNextSelectionUrlChange = false
    return
  }
  window.dispatchEvent(new Event(SELECTION_URL_CHANGE_EVENT))
}

function installHistorySelectionListeners(): void {
  if (typeof window === 'undefined') {
    return
  }
  const history = window.history
  const selectionWindow = window as SelectionHistoryWindow
  if (selectionWindow[SELECTION_HISTORY_STATE_KEY]?.history === history) {
    return
  }
  if (!history || typeof history.replaceState !== 'function') {
    return
  }
  const state: SelectionHistoryInstrumentation = {
    history,
    originalPushState: typeof history.pushState === 'function' ? history.pushState.bind(history) : null,
    originalReplaceState: history.replaceState.bind(history),
  }
  selectionWindow[SELECTION_HISTORY_STATE_KEY] = state
  if (state.originalPushState) {
    history.pushState = ((...args: Parameters<History['pushState']>) => {
      const result = state.originalPushState?.(...args)
      emitSelectionUrlChange()
      return result
    }) as History['pushState']
  }
  history.replaceState = ((...args: Parameters<History['replaceState']>) => {
    const result = state.originalReplaceState(...args)
    emitSelectionUrlChange()
    return result
  }) as History['replaceState']
}

export function subscribeSelectionUrlChanges(listener: () => void): () => void {
  if (typeof window === 'undefined' || typeof window.addEventListener !== 'function') {
    return () => undefined
  }
  installHistorySelectionListeners()
  window.addEventListener(SELECTION_URL_CHANGE_EVENT, listener)
  window.addEventListener('popstate', listener)
  window.addEventListener('hashchange', listener)
  return () => {
    window.removeEventListener(SELECTION_URL_CHANGE_EVENT, listener)
    window.removeEventListener('popstate', listener)
    window.removeEventListener('hashchange', listener)
  }
}

function persistSelectionToUrl(selection: WorkerRuntimeSelection): void {
  const currentUrl = new URL(window.location.href)
  const currentSheet = currentUrl.searchParams.get(SHEET_QUERY_PARAM)
  const currentCell = currentUrl.searchParams.get(CELL_QUERY_PARAM)
  if (currentSheet === selection.sheetName && currentCell === selection.address) {
    return
  }
  currentUrl.searchParams.set(SHEET_QUERY_PARAM, selection.sheetName)
  currentUrl.searchParams.set(CELL_QUERY_PARAM, selection.address)
  suppressNextSelectionUrlChange = true
  window.history.replaceState(window.history.state, '', currentUrl)
  if (suppressNextSelectionUrlChange) {
    suppressNextSelectionUrlChange = false
  }
}

function persistNormalizedSelection(scope: SelectionPersistenceScope, normalizedSelection: WorkerRuntimeSelection): void {
  removeStoredSelection(legacyStorageKey(scope.documentId))
  persistSelectionToUrl(normalizedSelection)
  window.localStorage.setItem(storageKey(scope), JSON.stringify(normalizedSelection))
}

function selectionPersistenceScopesEqual(left: SelectionPersistenceScope, right: SelectionPersistenceScope): boolean {
  return left.documentId === right.documentId && left.userId === right.userId
}

function clearPendingPersist(scope: SelectionPersistenceScope): void {
  const key = storageKey(scope)
  const pendingPersist = pendingPersists.get(key)
  if (!pendingPersist || typeof window === 'undefined') {
    pendingPersists.delete(key)
    return
  }
  globalThis.clearTimeout(pendingPersist.timeoutId)
  pendingPersists.delete(key)
}

function installScheduledPersistFlushListeners(): void {
  if (flushListenersInstalled || typeof window === 'undefined' || typeof window.addEventListener !== 'function') {
    return
  }
  flushListenersInstalled = true
  window.addEventListener('pagehide', flushScheduledSelectionPersistence)
  if (typeof document !== 'undefined') {
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'hidden') {
        flushScheduledSelectionPersistence()
      }
    })
  }
}

export function persistSelection(scope: SelectionPersistenceScope, selection: WorkerRuntimeSelection): void {
  if (typeof window === 'undefined') {
    return
  }
  try {
    clearPendingPersist(scope)
    const normalizedSelection = normalizeSelection(selection.sheetName, selection.address) ?? DEFAULT_SELECTION
    persistNormalizedSelection(scope, normalizedSelection)
  } catch {
    // Ignore storage failures and keep the runtime usable.
  }
}

export function scheduleSelectionPersistence(scope: SelectionPersistenceScope, selection: WorkerRuntimeSelection): void {
  if (typeof window === 'undefined') {
    return
  }
  try {
    installScheduledPersistFlushListeners()
    const normalizedSelection = normalizeSelection(selection.sheetName, selection.address) ?? DEFAULT_SELECTION
    const key = storageKey(scope)
    const pendingPersist = pendingPersists.get(key)
    if (
      pendingPersist &&
      selectionPersistenceScopesEqual(pendingPersist.scope, scope) &&
      pendingPersist.selection.sheetName === normalizedSelection.sheetName &&
      pendingPersist.selection.address === normalizedSelection.address
    ) {
      return
    }

    clearPendingPersist(scope)
    const timeoutId = globalThis.setTimeout(() => {
      const current = pendingPersists.get(key)
      if (!current || current.timeoutId !== timeoutId) {
        return
      }
      pendingPersists.delete(key)
      try {
        persistNormalizedSelection(current.scope, current.selection)
      } catch {
        // Ignore storage failures and keep the runtime usable.
      }
    }, SELECTION_PERSIST_DEBOUNCE_MS)
    pendingPersists.set(key, {
      scope,
      selection: normalizedSelection,
      timeoutId,
    })
  } catch {
    // Ignore storage failures and keep the runtime usable.
  }
}

export function flushScheduledSelectionPersistence(): void {
  if (pendingPersists.size === 0 || typeof window === 'undefined') {
    pendingPersists.clear()
    return
  }
  const pending = [...pendingPersists.values()]
  pendingPersists.clear()
  for (const current of pending) {
    globalThis.clearTimeout(current.timeoutId)
    try {
      persistNormalizedSelection(current.scope, current.selection)
    } catch {
      // Ignore storage failures and keep the runtime usable.
    }
  }
}
