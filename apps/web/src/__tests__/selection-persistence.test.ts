import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  flushScheduledSelectionPersistence,
  loadPersistedSelection,
  persistSelection,
  readSelectionFromUrl,
  scheduleSelectionPersistence,
  subscribeSelectionUrlChanges,
} from '../selection-persistence.js'

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
    expect(loadPersistedSelection('book-1')).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.has('bilig:selection:book-1')).toBe(false)
  })

  it('removes corrupt stored selection JSON after falling back', () => {
    storage.set('bilig:selection:book-1', '{')

    expect(loadPersistedSelection('book-1')).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.has('bilig:selection:book-1')).toBe(false)
  })

  it('falls back when corrupt stored selection cleanup fails', () => {
    storage.set('bilig:selection:book-1', '{')
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

    expect(loadPersistedSelection('book-1')).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.has('bilig:selection:book-1')).toBe(true)
  })

  it('removes syntactically valid stored selections with invalid cell addresses', () => {
    storage.set('bilig:selection:book-1', JSON.stringify({ sheetName: 'Sheet1', address: 'A0' }))

    expect(loadPersistedSelection('book-1')).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
    expect(storage.has('bilig:selection:book-1')).toBe(false)
  })

  it('restores the last stored sheet selection for a document', () => {
    persistSelection('book-1', { sheetName: 'Sheet3', address: 'G22' })

    expect(loadPersistedSelection('book-1')).toEqual({
      sheetName: 'Sheet3',
      address: 'G22',
    })
  })

  it('ignores invalid stored values', () => {
    storage.set('bilig:selection:book-1', '{"sheetName":"","address":42}')

    expect(loadPersistedSelection('book-1')).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
    })
  })

  it('prefers a URL-backed sheet selection over local storage', () => {
    storage.set('bilig:selection:book-1', JSON.stringify({ sheetName: 'Sheet3', address: 'G22' }))
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

    expect(loadPersistedSelection('book-1')).toEqual({
      sheetName: 'Sheet7',
      address: 'A1',
    })
  })

  it('reuses the stored address when the URL sheet matches it', () => {
    storage.set('bilig:selection:book-1', JSON.stringify({ sheetName: 'Sheet7', address: 'G22' }))
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

    expect(loadPersistedSelection('book-1')).toEqual({
      sheetName: 'Sheet7',
      address: 'G22',
    })
  })

  it('writes only the sheet into the URL state', () => {
    persistSelection('book-1', { sheetName: 'Sheet7', address: 'b12' })

    expect(replaceState).toHaveBeenCalledTimes(1)
    const [, , nextUrl] = replaceState.mock.calls[0]
    expect(String(nextUrl)).toBe('https://bilig.test/?sheet=Sheet7&cell=B12')
    expect(storage.get('bilig:selection:book-1')).toBe(JSON.stringify({ sheetName: 'Sheet7', address: 'B12' }))
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

    expect(loadPersistedSelection('book-1')).toEqual({
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

  it('emits selection URL changes from history writes', () => {
    const listener = vi.fn()
    const unsubscribe = subscribeSelectionUrlChanges(listener)

    persistSelection('book-1', { sheetName: 'Sheet1', address: 'C3' })

    expect(listener).toHaveBeenCalledTimes(1)
    unsubscribe()
    persistSelection('book-1', { sheetName: 'Sheet1', address: 'D4' })
    expect(listener).toHaveBeenCalledTimes(1)
  })

  it('coalesces rapid scheduled selection writes into the last selection', () => {
    vi.useFakeTimers()

    scheduleSelectionPersistence('book-1', { sheetName: 'Sheet1', address: 'A1' })
    scheduleSelectionPersistence('book-1', { sheetName: 'Sheet1', address: 'B1' })
    scheduleSelectionPersistence('book-1', { sheetName: 'Sheet1', address: 'C1' })

    expect(replaceState).not.toHaveBeenCalled()
    expect(storage.get('bilig:selection:book-1')).toBeUndefined()

    vi.advanceTimersByTime(119)
    expect(replaceState).not.toHaveBeenCalled()

    vi.advanceTimersByTime(1)
    expect(replaceState).toHaveBeenCalledTimes(1)
    const [, , nextUrl] = replaceState.mock.calls[0]
    expect(String(nextUrl)).toBe('https://bilig.test/?sheet=Sheet1&cell=C1')
    expect(storage.get('bilig:selection:book-1')).toBe(JSON.stringify({ sheetName: 'Sheet1', address: 'C1' }))
  })

  it('flushes the latest scheduled selection before an immediate persistence write', () => {
    vi.useFakeTimers()

    scheduleSelectionPersistence('book-1', { sheetName: 'Sheet1', address: 'B2' })
    persistSelection('book-1', { sheetName: 'Sheet2', address: 'D4' })

    vi.runOnlyPendingTimers()

    expect(replaceState).toHaveBeenCalledTimes(1)
    const [, , nextUrl] = replaceState.mock.calls[0]
    expect(String(nextUrl)).toBe('https://bilig.test/?sheet=Sheet2&cell=D4')
    expect(storage.get('bilig:selection:book-1')).toBe(JSON.stringify({ sheetName: 'Sheet2', address: 'D4' }))
  })
})
