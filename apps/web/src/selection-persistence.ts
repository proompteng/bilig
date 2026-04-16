import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { WorkerRuntimeSelection } from './runtime-session.js'

const DEFAULT_SELECTION: WorkerRuntimeSelection = {
  sheetName: 'Sheet1',
  address: 'A1',
}
const SHEET_QUERY_PARAM = 'sheet'
const CELL_QUERY_PARAM = 'cell'

function storageKey(documentId: string): string {
  return `bilig:selection:${documentId}`
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

function readStoredSelection(documentId: string): WorkerRuntimeSelection | null {
  if (typeof window === 'undefined') {
    return null
  }
  try {
    const raw = window.localStorage.getItem(storageKey(documentId))
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
      return null
    }
    return normalizeSelection(parsed['sheetName'], parsed['address'])
  } catch {
    return null
  }
}

export function loadPersistedSelection(documentId: string): WorkerRuntimeSelection {
  const storedSelection = readStoredSelection(documentId)
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

function persistSelectionToUrl(selection: WorkerRuntimeSelection): void {
  const currentUrl = new URL(window.location.href)
  const currentSheet = currentUrl.searchParams.get(SHEET_QUERY_PARAM)
  const currentCell = currentUrl.searchParams.get(CELL_QUERY_PARAM)
  if (currentSheet === selection.sheetName && currentCell === selection.address) {
    return
  }
  currentUrl.searchParams.set(SHEET_QUERY_PARAM, selection.sheetName)
  currentUrl.searchParams.set(CELL_QUERY_PARAM, selection.address)
  window.history.replaceState(window.history.state, '', currentUrl)
}

export function persistSelection(documentId: string, selection: WorkerRuntimeSelection): void {
  if (typeof window === 'undefined') {
    return
  }
  try {
    const normalizedSelection = normalizeSelection(selection.sheetName, selection.address) ?? DEFAULT_SELECTION
    persistSelectionToUrl(normalizedSelection)
    window.localStorage.setItem(storageKey(documentId), JSON.stringify(normalizedSelection))
  } catch {
    // Ignore storage failures and keep the runtime usable.
  }
}
