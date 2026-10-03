import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  clearStoredSession,
  loadStoredDrafts,
  loadStoredSession,
  persistStoredDrafts,
  persistStoredSession,
} from '../workbook-agent-pane-storage.js'
import { captureExpectedConsoleDebug, captureExpectedConsoleDebugMessages } from './expected-console.js'

const alexScope = {
  documentId: 'doc-1',
  userId: 'alex@example.com',
}
const caseyScope = {
  documentId: 'doc-1',
  userId: 'casey@example.com',
}

describe('workbook agent pane storage', () => {
  const storage = new Map<string, string>()

  beforeEach(() => {
    storage.clear()
    vi.stubGlobal('window', {
      sessionStorage: {
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
  })

  afterEach(() => {
    vi.unstubAllGlobals()
    vi.restoreAllMocks()
  })

  it('removes corrupt stored session JSON after falling back', () => {
    const storageDebug = captureExpectedConsoleDebug('Failed to load stored workbook agent session')
    storage.set('bilig:workbook-agent:doc-1:alex%40example.com', '{')

    expect(loadStoredSession(alexScope)).toBeNull()
    expect(storage.has('bilig:workbook-agent:doc-1:alex%40example.com')).toBe(false)
    storageDebug.expectLogCount(1)
  })

  it('rejects and removes blank stored thread ids', () => {
    storage.set('bilig:workbook-agent:doc-1:alex%40example.com', JSON.stringify({ threadId: '   ' }))
    storage.set('bilig:workbook-agent:doc-1', JSON.stringify({ threadId: 'legacy-thread' }))

    expect(loadStoredSession(alexScope)).toBeNull()
    expect(storage.has('bilig:workbook-agent:doc-1:alex%40example.com')).toBe(false)
    expect(storage.has('bilig:workbook-agent:doc-1')).toBe(true)

    storage.set('bilig:workbook-agent:doc-1', JSON.stringify({ threadId: 'legacy-thread' }))
    persistStoredSession(alexScope, { threadId: '   ' })
    expect(storage.has('bilig:workbook-agent:doc-1:alex%40example.com')).toBe(false)
    expect(storage.has('bilig:workbook-agent:doc-1')).toBe(true)
  })

  it('normalizes and persists valid stored thread ids', () => {
    persistStoredSession(alexScope, { threadId: '  thr-1  ' })

    expect(loadStoredSession(alexScope)).toEqual({ threadId: 'thr-1' })
    expect(storage.get('bilig:workbook-agent:doc-1:alex%40example.com')).toBe(JSON.stringify({ threadId: 'thr-1' }))
  })

  it('does not restore another user assistant session for the same document', () => {
    persistStoredSession(alexScope, { threadId: 'alex-private-thread' })

    expect(loadStoredSession(caseyScope)).toBeNull()
  })

  it('ignores obsolete document-only assistant sessions', () => {
    storage.set('bilig:workbook-agent:doc-1', JSON.stringify({ threadId: 'legacy-thread' }))

    expect(loadStoredSession(alexScope)).toBeNull()
    expect(storage.has('bilig:workbook-agent:doc-1')).toBe(true)
  })

  it('removes corrupt stored draft JSON after falling back', () => {
    const storageDebug = captureExpectedConsoleDebug('Failed to load stored workbook agent draft')
    storage.set('bilig:workbook-agent-drafts:doc-1:alex%40example.com', '{')

    expect(loadStoredDrafts(alexScope)).toEqual({})
    expect(storage.has('bilig:workbook-agent-drafts:doc-1:alex%40example.com')).toBe(false)
    storageDebug.expectLogCount(1)
  })

  it('self-heals stored draft maps with non-string values', () => {
    storage.set('bilig:workbook-agent-drafts:doc-1:alex%40example.com', JSON.stringify({ keep: 'draft', drop: 42 }))

    expect(loadStoredDrafts(alexScope)).toEqual({ keep: 'draft' })
    expect(storage.get('bilig:workbook-agent-drafts:doc-1:alex%40example.com')).toBe(JSON.stringify({ keep: 'draft' }))
  })

  it('does not restore another user assistant drafts for the same document', () => {
    persistStoredDrafts(alexScope, { 'new:private': 'alex draft' })

    expect(loadStoredDrafts(caseyScope)).toEqual({})
  })

  it('ignores obsolete document-only assistant drafts', () => {
    storage.set('bilig:workbook-agent-drafts:doc-1', JSON.stringify({ 'new:private': 'legacy draft' }))

    expect(loadStoredDrafts(alexScope)).toEqual({})
    expect(storage.has('bilig:workbook-agent-drafts:doc-1')).toBe(true)

    storage.set('bilig:workbook-agent-drafts:doc-1', JSON.stringify({ 'new:private': 'legacy draft' }))
    persistStoredDrafts(alexScope, {})
    expect(storage.has('bilig:workbook-agent-drafts:doc-1')).toBe(true)
  })

  it('does not throw when session storage writes fail', () => {
    const storageDebug = captureExpectedConsoleDebugMessages()
    vi.stubGlobal('window', {
      sessionStorage: {
        getItem() {
          return null
        },
        removeItem() {
          throw new Error('storage denied')
        },
        setItem() {
          throw new Error('storage denied')
        },
      },
    })

    expect(() => persistStoredSession(alexScope, { threadId: 'thr-1' })).not.toThrow()
    expect(() => persistStoredDrafts(alexScope, { key: 'draft' })).not.toThrow()
    expect(() => clearStoredSession(alexScope)).not.toThrow()
    storageDebug.expectMessageCount('Failed to clear workbook agent storage', 1)
    storageDebug.expectMessageCount('Failed to persist workbook agent storage', 2)
  })
})
