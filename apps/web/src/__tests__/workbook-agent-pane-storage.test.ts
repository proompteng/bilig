import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  clearStoredSession,
  loadStoredDrafts,
  loadStoredSession,
  persistStoredDrafts,
  persistStoredSession,
} from '../workbook-agent-pane-storage.js'

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
  })

  it('removes corrupt stored session JSON after falling back', () => {
    storage.set('bilig:workbook-agent:doc-1:alex%40example.com', '{')

    expect(loadStoredSession(alexScope)).toBeNull()
    expect(storage.has('bilig:workbook-agent:doc-1:alex%40example.com')).toBe(false)
  })

  it('rejects and removes blank stored thread ids', () => {
    storage.set('bilig:workbook-agent:doc-1:alex%40example.com', JSON.stringify({ threadId: '   ' }))
    storage.set('bilig:workbook-agent:doc-1', JSON.stringify({ threadId: 'legacy-thread' }))

    expect(loadStoredSession(alexScope)).toBeNull()
    expect(storage.has('bilig:workbook-agent:doc-1:alex%40example.com')).toBe(false)
    expect(storage.has('bilig:workbook-agent:doc-1')).toBe(false)

    storage.set('bilig:workbook-agent:doc-1', JSON.stringify({ threadId: 'legacy-thread' }))
    persistStoredSession(alexScope, { threadId: '   ' })
    expect(storage.has('bilig:workbook-agent:doc-1:alex%40example.com')).toBe(false)
    expect(storage.has('bilig:workbook-agent:doc-1')).toBe(false)
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

  it('removes legacy document-only assistant sessions instead of restoring unscoped threads', () => {
    storage.set('bilig:workbook-agent:doc-1', JSON.stringify({ threadId: 'legacy-thread' }))

    expect(loadStoredSession(alexScope)).toBeNull()
    expect(storage.has('bilig:workbook-agent:doc-1')).toBe(false)
  })

  it('removes corrupt stored draft JSON after falling back', () => {
    storage.set('bilig:workbook-agent-drafts:doc-1:alex%40example.com', '{')

    expect(loadStoredDrafts(alexScope)).toEqual({})
    expect(storage.has('bilig:workbook-agent-drafts:doc-1:alex%40example.com')).toBe(false)
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

  it('removes legacy document-only assistant drafts instead of restoring unscoped text', () => {
    storage.set('bilig:workbook-agent-drafts:doc-1', JSON.stringify({ 'new:private': 'legacy draft' }))

    expect(loadStoredDrafts(alexScope)).toEqual({})
    expect(storage.has('bilig:workbook-agent-drafts:doc-1')).toBe(false)

    storage.set('bilig:workbook-agent-drafts:doc-1', JSON.stringify({ 'new:private': 'legacy draft' }))
    persistStoredDrafts(alexScope, {})
    expect(storage.has('bilig:workbook-agent-drafts:doc-1')).toBe(false)
  })

  it('does not throw when session storage writes fail', () => {
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
  })
})
