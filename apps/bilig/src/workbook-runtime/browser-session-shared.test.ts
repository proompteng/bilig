import { describe, expect, it, vi } from 'vitest'
import type { HelloFrame } from '@bilig/binary-protocol'
import { openWorkbookBrowserSession } from './browser-session-shared.js'

const helloFrame: HelloFrame = {
  kind: 'hello',
  documentId: 'doc-1',
  replicaId: 'browser-1',
  sessionId: 'session-1',
  protocolVersion: 1,
  lastServerCursor: 0,
  capabilities: [],
}

describe('browser-session-shared', () => {
  it('registers the browser session before replaying missed frames', async () => {
    const register = vi.fn(async () => {})
    const listMissedFrames = vi.fn(async () => [
      {
        kind: 'appendBatch' as const,
        documentId: 'doc-1',
        cursor: 4,
        batch: {
          id: 'batch-1',
          replicaId: 'replica-1',
          clock: { counter: 1 },
          ops: [],
        },
      },
    ])

    const frames = await openWorkbookBrowserSession(helloFrame, {
      register,
      latestCursor: 4,
      latestSnapshot: null,
      listMissedFrames,
    })

    expect(register).toHaveBeenCalledWith(helloFrame)
    expect(listMissedFrames).toHaveBeenCalledWith(0)
    expect(frames).toContainEqual(
      expect.objectContaining({
        kind: 'appendBatch',
        documentId: 'doc-1',
      }),
    )
    expect(frames.at(-1)).toEqual({
      kind: 'cursorWatermark',
      documentId: 'doc-1',
      cursor: 4,
      compactedCursor: 0,
    })
  })
})
