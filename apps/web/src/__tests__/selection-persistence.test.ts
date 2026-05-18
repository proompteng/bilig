import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  flushScheduledSelectionPersistence,
  loadPersistedSelection,
  persistSelection,
  readSelectionFromUrl,
  scheduleSelectionPersistence,
  subscribeSelectionUrlChanges,
} from '../selection-persistence.js'

const alexScope = {
  documentId: 'book-1',
  userId: 'alex@example.com',
}
const caseyScope = {
  documentId: 'book-1',
  userId: 'casey@example.com',
}
const alexSelectionStorageKey = 'bilig:selection:book-1:alex%40example.com'
const caseySelectionStorageKey = 'bilig:selection:book-1:casey%40example.com'
const legacySelectionStorageKey = 'bilig:selection:book-1'

describe('selection persistence', () => {
  const storage = new Map<string, string>()
  const replaceState = vi.fn()
  let eventTarget: EventTarget

  beforeEach(() => {
    storage.clear()
    replaceState.mockReset()
    eventTarget = new EventTarget()
    vi.stubGlobal('window', {
      addEventListener: eventTarget.addEventListener.bind(eventTarget),
      dispatchEvent: eventTarget.dispatchEvent.bind(eventTarget),
      history: {
        replaceState,
        state: { from: 'test' },
      },
      localStorage: {
        getItem(key: string) {
          return storage.get(key) ?? null
        },
        setItem(key: string, value: string) {
          storage.set(key, value)
        },
        removeItem(key: string) {
          storage.delete(key)
        },
        clear() {
          storage.clear()
        },
      },
      location: new URL('https://bilig.test/'),
      removeEventListener: eventTarget.removeEventListener.bind(eventTarget),
    })
  })

  afterEach(() => {
    flushScheduledSelectionPersistence()
    vi.useRealTimers()
    vi.unstubAllGlobals()
  })

  it('falls back to Sheet1!A1 when nothing is stored', () => {
    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.has(alexSelectionStorageKey)).toBe(false)
  })

  it('removes corrupt stored selection JSON after falling back', () => {
    storage.set(alexSelectionStorageKey, '{')

    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.has(alexSelectionStorageKey)).toBe(false)
  })

  it('falls back when corrupt stored selection cleanup fails', () => {
    storage.set(alexSelectionStorageKey, '{')
    vi.stubGlobal('window', {
      history: {
        replaceState,
        state: { from: 'test' },
      },
      localStorage: {
        getItem(key: string) {
          return storage.get(key) ?? null
        },
        removeItem() {
          throw new Error('storage denied')
        },
      },
      location: new URL('https://bilig.test/'),
    })

    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.has(alexSelectionStorageKey)).toBe(true)
  })

  it('removes syntactically valid stored selections with invalid cell addresses', () => {
    storage.set(alexSelectionStorageKey, JSON.stringify({ sheetName: 'Sheet1', address: 'A0' }))

    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.has(alexSelectionStorageKey)).toBe(false)
  })

  it('restores the last stored sheet selection for a document', () => {
    persistSelection(alexScope, { sheetName: 'Sheet3', address: 'G22' })

    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet3',
      address: 'G22',
    })
  })

  it('does not restore another user selection for the same document', () => {
    persistSelection(alexScope, { sheetName: 'Sheet3', address: 'G22' })

    expect(loadPersistedSelection(caseyScope)).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.get(alexSelectionStorageKey)).toBe(JSON.stringify({ sheetName: 'Sheet3', address: 'G22' }))
    expect(storage.has(caseySelectionStorageKey)).toBe(false)
  })

  it('removes legacy document-only selection instead of restoring unscoped state', () => {
    storage.set(legacySelectionStorageKey, JSON.stringify({ sheetName: 'PrivateSheet', address: 'D4' }))

    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.has(legacySelectionStorageKey)).toBe(false)

    storage.set(legacySelectionStorageKey, JSON.stringify({ sheetName: 'PrivateSheet', address: 'D4' }))
    persistSelection(alexScope, { sheetName: 'Sheet2', address: 'C3' })
    expect(storage.has(legacySelectionStorageKey)).toBe(false)
  })

  it('ignores invalid stored values', () => {
    storage.set(alexSelectionStorageKey, '{"sheetName":"","address":42}')

    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
  })

  it('prefers a URL-backed sheet selection over local storage', () => {
    storage.set(alexSelectionStorageKey, JSON.stringify({ sheetName: 'Sheet3', address: 'G22' }))
    vi.stubGlobal('window', {
      history: {
        replaceState,
        state: { from: 'test' },
      },
      localStorage: {
        getItem(key: string) {
          return storage.get(key) ?? null
        },
        setItem(key: string, value: string) {
          storage.set(key, value)
        },
        clear() {
          storage.clear()
        },
      },
      location: new URL('https://bilig.test/?sheet=Sheet7'),
    })

    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet7',
      address: 'A1',
    })
  })

  it('reuses the stored address when the URL sheet matches it', () => {
    storage.set(alexSelectionStorageKey, JSON.stringify({ sheetName: 'Sheet7', address: 'G22' }))
    vi.stubGlobal('window', {
      history: {
        replaceState,
        state: { from: 'test' },
      },
      localStorage: {
        getItem(key: string) {
          return storage.get(key) ?? null
        },
        setItem(key: string, value: string) {
          storage.set(key, value)
        },
        clear() {
          storage.clear()
        },
      },
      location: new URL('https://bilig.test/?sheet=Sheet7'),
    })

    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet7',
      address: 'G22',
    })
  })

  it('writes only the sheet into the URL state', () => {
    persistSelection(alexScope, { sheetName: 'Sheet7', address: 'b12' })

    expect(replaceState).toHaveBeenCalledTimes(1)
    const [, , nextUrl] = replaceState.mock.calls[0]
    expect(String(nextUrl)).toBe('https://bilig.test/?sheet=Sheet7&cell=B12')
    expect(storage.get(alexSelectionStorageKey)).toBe(JSON.stringify({ sheetName: 'Sheet7', address: 'B12' }))
  })

  it('restores a URL-backed cell selection when both sheet and cell are present', () => {
    eventTarget = new EventTarget()
    vi.stubGlobal('window', {
      addEventListener: eventTarget.addEventListener.bind(eventTarget),
      dispatchEvent: eventTarget.dispatchEvent.bind(eventTarget),
      history: {
        replaceState,
        state: { from: 'test' },
      },
      localStorage: {
        getItem(key: string) {
          return storage.get(key) ?? null
        },
        setItem(key: string, value: string) {
          storage.set(key, value)
        },
        clear() {
          storage.clear()
        },
      },
      location: new URL('https://bilig.test/?sheet=Sheet9&cell=d14'),
      removeEventListener: eventTarget.removeEventListener.bind(eventTarget),
    })

    expect(loadPersistedSelection(alexScope)).toEqual({
      sheetName: 'Sheet9',
      address: 'D14',
    })
  })

  it('reads explicit URL selection for same-document navigation', () => {
    eventTarget = new EventTarget()
    vi.stubGlobal('window', {
      addEventListener: eventTarget.addEventListener.bind(eventTarget),
      dispatchEvent: eventTarget.dispatchEvent.bind(eventTarget),
      history: {
        replaceState,
        state: { from: 'test' },
      },
      localStorage: {
        getItem(key: string) {
          return storage.get(key) ?? null
        },
        setItem(key: string, value: string) {
          storage.set(key, value)
        },
        clear() {
          storage.clear()
        },
      },
      location: new URL('https://bilig.test/?sheet=Prepaid+Template&cell=f16'),
      removeEventListener: eventTarget.removeEventListener.bind(eventTarget),
    })

    expect(readSelectionFromUrl()).toEqual({
      sheetName: 'Prepaid Template',
      address: 'F16',
    })
  })

  it('emits selection URL changes from external history writes', () => {
    const listener = vi.fn()
    const unsubscribe = subscribeSelectionUrlChanges(listener)

    window.history.replaceState(window.history.state, '', 'https://bilig.test/?sheet=Sheet1&cell=C3')

    expect(listener).toHaveBeenCalledTimes(1)
    unsubscribe()
    window.history.replaceState(window.history.state, '', 'https://bilig.test/?sheet=Sheet1&cell=D4')
    expect(listener).toHaveBeenCalledTimes(1)
  })

  it('does not re-emit URL changes from local selection persistence writes', () => {
    const listener = vi.fn()
    const unsubscribe = subscribeSelectionUrlChanges(listener)

    persistSelection(alexScope, { sheetName: 'Sheet1', address: 'C3' })

    expect(listener).not.toHaveBeenCalled()
    unsubscribe()
  })

  it('coalesces rapid scheduled selection writes into the last selection', () => {
    vi.useFakeTimers()

    scheduleSelectionPersistence(alexScope, { sheetName: 'Sheet1', address: 'A1' })
    scheduleSelectionPersistence(alexScope, { sheetName: 'Sheet1', address: 'B1' })
    scheduleSelectionPersistence(alexScope, { sheetName: 'Sheet1', address: 'C1' })

    expect(replaceState).not.toHaveBeenCalled()
    expect(storage.get(alexSelectionStorageKey)).toBeUndefined()

    vi.advanceTimersByTime(119)
    expect(replaceState).not.toHaveBeenCalled()

    vi.advanceTimersByTime(1)
    expect(replaceState).toHaveBeenCalledTimes(1)
    const [, , nextUrl] = replaceState.mock.calls[0]
    expect(String(nextUrl)).toBe('https://bilig.test/?sheet=Sheet1&cell=C1')
    expect(storage.get(alexSelectionStorageKey)).toBe(JSON.stringify({ sheetName: 'Sheet1', address: 'C1' }))
  })

  it('flushes the latest scheduled selection before an immediate persistence write', () => {
    vi.useFakeTimers()

    scheduleSelectionPersistence(alexScope, { sheetName: 'Sheet1', address: 'B2' })
    persistSelection(alexScope, { sheetName: 'Sheet2', address: 'D4' })

    vi.runOnlyPendingTimers()

    expect(replaceState).toHaveBeenCalledTimes(1)
    const [, , nextUrl] = replaceState.mock.calls[0]
    expect(String(nextUrl)).toBe('https://bilig.test/?sheet=Sheet2&cell=D4')
    expect(storage.get(alexSelectionStorageKey)).toBe(JSON.stringify({ sheetName: 'Sheet2', address: 'D4' }))
  })

  it('keeps pending scheduled selection writes isolated by user scope', () => {
    vi.useFakeTimers()

    scheduleSelectionPersistence(alexScope, { sheetName: 'Sheet1', address: 'B2' })
    scheduleSelectionPersistence(caseyScope, { sheetName: 'Sheet4', address: 'D4' })

    vi.advanceTimersByTime(120)

    expect(storage.get(alexSelectionStorageKey)).toBe(JSON.stringify({ sheetName: 'Sheet1', address: 'B2' }))
    expect(storage.get(caseySelectionStorageKey)).toBe(JSON.stringify({ sheetName: 'Sheet4', address: 'D4' }))
  })
})
