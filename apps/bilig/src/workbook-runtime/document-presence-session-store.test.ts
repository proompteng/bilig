import { describe, expect, it } from 'vitest'
import { createInMemoryDocumentPersistence } from '@bilig/storage-server'
import {
  closePresenceBackedWorkbookSession,
  countPresenceBackedWorkbookSessions,
  joinOwnedBrowserSession,
  openPresenceBackedWorkbookSession,
} from './document-presence-session-store.js'

describe('document-presence-session-store', () => {
  it('opens and closes presence-backed workbook sessions', async () => {
    const persistence = createInMemoryDocumentPersistence()

    const sessionId = await openPresenceBackedWorkbookSession(persistence, 'doc-1', 'replica-1')
    expect(sessionId).toBe('doc-1:replica-1')
    expect(await countPresenceBackedWorkbookSessions(persistence, sessionId)).toBe(1)

    await closePresenceBackedWorkbookSession(persistence, sessionId)
    expect(await countPresenceBackedWorkbookSessions(persistence, sessionId)).toBe(0)
  })

  it('joins browser sessions and claims ownership', async () => {
    const persistence = createInMemoryDocumentPersistence()

    await joinOwnedBrowserSession(persistence, 'bilig-app', 'doc-2', 'browser-1')

    expect(await persistence.presence.sessions('doc-2')).toEqual(['browser-1'])
    expect(await persistence.ownership.owner('doc-2')).toBe('bilig-app')
  })
})
