import { describe, expect, it, vi } from 'vitest'
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, createSnapshotChunkFrames } from '@bilig/binary-protocol'
import { createInMemoryDocumentPersistence } from '@bilig/storage-server'
import type { AgentFrame } from '@bilig/agent-api'
import { DocumentSessionManager } from './document-session-manager.js'

describe('DocumentSessionManager', () => {
  it('broadcasts sync frames to attached browser subscribers', async () => {
    const sent: unknown[] = []
    const manager = new DocumentSessionManager(createInMemoryDocumentPersistence())
    const detach = manager.attachBrowser('doc-broadcast', 'browser-1', (frame) => {
      sent.push(frame)
    })

    await manager.handleSyncFrame({
      kind: 'appendBatch',
      documentId: 'doc-broadcast',
      cursor: 0,
      batch: {
        id: 'batch-1',
        replicaId: 'replica-1',
        clock: { counter: 1 },
        ops: [],
      },
    })

    expect(sent).toContainEqual(
      expect.objectContaining({
        kind: 'appendBatch',
        documentId: 'doc-broadcast',
      }),
    )

    detach()
  })

  it('delegates open and close session requests through the worksheet executor', async () => {
    const execute = vi.fn(async (frame: AgentFrame) => {
      if (frame.kind !== 'request') {
        throw new Error('expected request frame')
      }
      if (frame.request.kind === 'openWorkbookSession') {
        return {
          kind: 'response',
          response: {
            kind: 'ok',
            id: frame.request.id,
            sessionId: `${frame.request.documentId}:${frame.request.replicaId}`,
          },
        } satisfies AgentFrame
      }
      return {
        kind: 'response',
        response: {
          kind: 'ok',
          id: frame.request.id,
        },
      } satisfies AgentFrame
    })

    const manager = new DocumentSessionManager(createInMemoryDocumentPersistence(), 'bilig-app', {
      execute,
    })

    const openResponse = await manager.handleAgentFrame({
      kind: 'request',
      request: {
        kind: 'openWorkbookSession',
        id: 'open-1',
        documentId: 'doc-1',
        replicaId: 'replica-1',
      },
    } satisfies AgentFrame)

    expect(openResponse).toEqual({
      kind: 'response',
      response: {
        kind: 'ok',
        id: 'open-1',
        sessionId: 'doc-1:replica-1',
      },
    })
    expect(execute).toHaveBeenCalledTimes(1)
    expect((await manager.getDocumentState('doc-1')).sessions).toContain('doc-1:replica-1')

    const closeResponse = await manager.handleAgentFrame({
      kind: 'request',
      request: {
        kind: 'closeWorkbookSession',
        id: 'close-1',
        sessionId: 'doc-1:replica-1',
      },
    } satisfies AgentFrame)

    expect(closeResponse).toEqual({
      kind: 'response',
      response: {
        kind: 'ok',
        id: 'close-1',
      },
    })
    expect(execute).toHaveBeenCalledTimes(2)
    expect((await manager.getDocumentState('doc-1')).sessions).toEqual([])
  })

  it('assembles uploaded snapshot chunks into persisted browser snapshots', async () => {
    const manager = new DocumentSessionManager(createInMemoryDocumentPersistence())
    const snapshot = {
      version: 1 as const,
      workbook: {
        name: 'Imported',
      },
      sheets: [
        {
          name: 'Sheet1',
          order: 0,
          cells: [{ address: 'A1', value: 42 }],
        },
      ],
    }
    const bytes = new TextEncoder().encode(JSON.stringify(snapshot))
    const chunkFrames = createSnapshotChunkFrames({
      documentId: 'doc-snapshot',
      snapshotId: 'doc-snapshot:snapshot:1',
      cursor: 6,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes,
    })

    await Promise.all(chunkFrames.map((frame) => manager.handleSyncFrame(frame)))

    const helloFrames = await manager.openBrowserSession({
      kind: 'hello',
      documentId: 'doc-snapshot',
      replicaId: 'browser-1',
      sessionId: 'browser-1',
      protocolVersion: 1,
      lastServerCursor: 0,
      capabilities: [],
    })

    expect(helloFrames.some((frame) => frame.kind === 'snapshotChunk')).toBe(true)
    expect(helloFrames).toContainEqual({
      kind: 'cursorWatermark',
      documentId: 'doc-snapshot',
      cursor: 0,
      compactedCursor: 6,
    })
  })
})
