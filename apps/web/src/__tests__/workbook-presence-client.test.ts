import { afterEach, describe, expect, it, vi } from 'vitest'
import { loadOrCreateWorkbookPresenceClientId } from '../workbook-presence-client.js'

describe('workbook presence client id persistence', () => {
  const storage = new Map<string, string>()

  afterEach(() => {
    storage.clear()
    vi.unstubAllGlobals()
  })

  function stubStorage() {
    vi.stubGlobal('window', {
      localStorage: {
        getItem(key: string) {
          return storage.get(key) ?? null
        },
        removeItem(key: string) {
          storage.delete(key)
        },
        setItem(key: string, value: string) {
          storage.set(key, value)
        },
      },
    })
  }

  function stubUuid(value: string) {
    vi.stubGlobal('crypto', {
      randomUUID: () => value,
    })
  }

  it('normalizes a stored presence id before returning it', () => {
    stubStorage()
    storage.set('bilig:workbook-presence-client-id', '  presence:existing-client  ')

    expect(loadOrCreateWorkbookPresenceClientId()).toBe('presence:existing-client')
  })

  it('replaces stored values outside the presence id contract', () => {
    stubStorage()
    stubUuid('new-client')
    storage.set('bilig:workbook-presence-client-id', 'not-a-presence-id')

    expect(loadOrCreateWorkbookPresenceClientId()).toBe('presence:new-client')
    expect(storage.get('bilig:workbook-presence-client-id')).toBe('presence:new-client')
  })

  it('replaces stored presence ids with embedded whitespace', () => {
    stubStorage()
    stubUuid('clean-client')
    storage.set('bilig:workbook-presence-client-id', 'presence:bad client')

    expect(loadOrCreateWorkbookPresenceClientId()).toBe('presence:clean-client')
    expect(storage.get('bilig:workbook-presence-client-id')).toBe('presence:clean-client')
  })

  it('falls back to a fresh id when localStorage cleanup or writes fail', () => {
    stubUuid('fallback-client')
    vi.stubGlobal('window', {
      localStorage: {
        getItem() {
          return 'invalid'
        },
        removeItem() {
          throw new Error('storage denied')
        },
        setItem() {
          throw new Error('storage denied')
        },
      },
    })

    expect(loadOrCreateWorkbookPresenceClientId()).toBe('presence:fallback-client')
  })
})
