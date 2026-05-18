import { describe, expect, it } from 'vitest'
import { createEmptyWorkbookSnapshot } from '@bilig/zero-sync'

import {
  isAuthoritativeSnapshotNewerForRebase,
  loadAuthoritativeEventBatch,
  loadLatestWorkbookSnapshot,
  parseSnapshotRevisionHeader,
  shouldApplyAuthoritativeEventBatch,
  shouldInstallBootstrapSnapshot,
  shouldQueueAuthoritativeRebase,
} from '../runtime-authoritative-sync.js'

async function mismatchedAuthoritativeEventResponse(): Promise<Response> {
  return new Response(
    JSON.stringify({
      afterRevision: 4,
      headRevision: 5,
      calculatedRevision: 5,
      events: [
        {
          revision: 5,
          clientMutationId: null,
          payload: {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'A1',
            value: 'ok',
          },
        },
      ],
    }),
  )
}

async function currentAuthoritativeSnapshotResponse(): Promise<Response> {
  return new Response(JSON.stringify(createEmptyWorkbookSnapshot('doc-1')), {
    status: 200,
    headers: { 'x-bilig-snapshot-cursor': '17' },
  })
}

describe('runtime session snapshot revision parsing', () => {
  it('accepts safe non-negative integer snapshot cursor headers', () => {
    expect(parseSnapshotRevisionHeader('0')).toBe(0)
    expect(parseSnapshotRevisionHeader('37')).toBe(37)
    expect(parseSnapshotRevisionHeader(' 451 ')).toBe(451)
    expect(parseSnapshotRevisionHeader(String(Number.MAX_SAFE_INTEGER))).toBe(Number.MAX_SAFE_INTEGER)
  })

  it.each([null, '', ' ', '-1', '+1', '01', '1.5', '37abc', String(Number.MAX_SAFE_INTEGER + 1)])(
    'rejects malformed snapshot cursor header %s',
    (value) => {
      expect(parseSnapshotRevisionHeader(value)).toBeNull()
    },
  )

  it('loads and validates authoritative event batches against the requested cursor', async () => {
    await expect(
      loadAuthoritativeEventBatch({
        documentId: 'doc-1',
        afterRevision: 3,
        fetchImpl: mismatchedAuthoritativeEventResponse,
      }),
    ).rejects.toThrow('Authoritative event payload does not match the expected schema')
  })

  it('loads authoritative snapshots with safe revision headers only', async () => {
    await expect(
      loadLatestWorkbookSnapshot({ documentId: 'doc-1', fetchImpl: currentAuthoritativeSnapshotResponse }),
    ).resolves.toMatchObject({
      revision: 17,
    })
  })

  it('keeps authoritative event and snapshot revision planning fail-closed', () => {
    expect(
      shouldApplyAuthoritativeEventBatch({
        currentAuthoritativeRevision: 5,
        eventBatch: { events: [], headRevision: 6, calculatedRevision: 6 },
      }),
    ).toBe(false)
    expect(
      shouldApplyAuthoritativeEventBatch({
        currentAuthoritativeRevision: 5,
        eventBatch: { events: [{}], headRevision: 6, calculatedRevision: 5 },
      }),
    ).toBe(false)
    expect(
      shouldQueueAuthoritativeRebase({
        revisionState: { headRevision: 5, calculatedRevision: 5 },
        currentAuthoritativeRevision: 5,
        currentCalculatedRevision: 5,
      }),
    ).toBe(false)
    expect(
      shouldQueueAuthoritativeRebase({
        revisionState: { headRevision: 5, calculatedRevision: 6 },
        currentAuthoritativeRevision: 5,
        currentCalculatedRevision: 5,
      }),
    ).toBe(true)
    expect(shouldInstallBootstrapSnapshot({ snapshotRevision: 3, currentAuthoritativeRevision: 4 })).toBe(false)
    expect(
      isAuthoritativeSnapshotNewerForRebase({
        snapshotRevision: 5,
        snapshotCalculatedRevision: 5,
        currentAuthoritativeRevision: 5,
        currentCalculatedRevision: 5,
      }),
    ).toBe(false)
    expect(
      isAuthoritativeSnapshotNewerForRebase({
        snapshotRevision: 5,
        snapshotCalculatedRevision: 6,
        currentAuthoritativeRevision: 5,
        currentCalculatedRevision: 5,
      }),
    ).toBe(true)
  })
})
