import type { AgentFrame } from '@bilig/agent-api'
import { describe, expect, it, vi } from 'vitest'
import { CSV_CONTENT_TYPE, XLSX_CONTENT_TYPE } from '@bilig/agent-api'
import * as XLSX from 'xlsx'
import {
  createWorkbookLoadOptions,
  createCloseWorkbookSessionResponse,
  createOpenWorkbookSessionResponse,
  documentIdFromSessionId,
  handleWorkbookAgentFrame,
  loadWorkbookIntoRuntime,
} from './workbook-session-shared.js'

describe('workbook-session-shared', () => {
  it('derives a document id from a session id', () => {
    expect(documentIdFromSessionId('doc-1:replica-1')).toBe('doc-1')
    expect(documentIdFromSessionId('doc-2')).toBe('doc-2')
  })

  it('creates standard open and close session responses', () => {
    expect(createOpenWorkbookSessionResponse('open-1', 'doc-1:replica-1')).toEqual({
      kind: 'ok',
      id: 'open-1',
      sessionId: 'doc-1:replica-1',
    })
    expect(createCloseWorkbookSessionResponse('close-1')).toEqual({
      kind: 'ok',
      id: 'close-1',
    })
  })

  it('prepares workbook imports once and delegates registration and publish hooks', async () => {
    const registerPreparedSession = vi.fn()
    const publishImportedSnapshot = vi.fn()
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['hello']]), 'Sheet1')
    const encodedWorkbook: unknown = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
    if (!(encodedWorkbook instanceof Uint8Array)) {
      throw new Error('Expected xlsx writer to return workbook bytes')
    }

    const response = await loadWorkbookIntoRuntime(
      {
        kind: 'loadWorkbookFile',
        id: 'load-1',
        replicaId: 'replica-1',
        fileName: 'tiny.xlsx',
        contentType: XLSX_CONTENT_TYPE,
        openMode: 'create',
        bytesBase64: Buffer.from(encodedWorkbook).toString('base64'),
      },
      {
        serverUrl: 'http://127.0.0.1:4321',
        browserAppBaseUrl: 'http://127.0.0.1:3000',
      },
      {
        registerPreparedSession,
        publishImportedSnapshot,
      },
    )

    expect(registerPreparedSession).toHaveBeenCalledTimes(1)
    expect(publishImportedSnapshot).toHaveBeenCalledTimes(1)
    expect(response).toEqual(
      expect.objectContaining({
        kind: 'workbookLoaded',
        id: 'load-1',
        sessionId: expect.stringContaining(':replica-1'),
        serverUrl: 'http://127.0.0.1:4321',
      }),
    )
  })

  it('supports csv imports through the shared workbook load path', async () => {
    const publishImportedSnapshot = vi.fn()

    const response = await loadWorkbookIntoRuntime(
      {
        kind: 'loadWorkbookFile',
        id: 'load-csv-1',
        replicaId: 'replica-1',
        fileName: 'tiny.csv',
        contentType: CSV_CONTENT_TYPE,
        openMode: 'create',
        bytesBase64: Buffer.from('Label,Value\nalpha,12').toString('base64'),
      },
      {
        serverUrl: 'http://127.0.0.1:4321',
      },
      {
        registerPreparedSession: vi.fn(),
        publishImportedSnapshot,
      },
    )

    expect(response).toEqual(
      expect.objectContaining({
        kind: 'workbookLoaded',
        id: 'load-csv-1',
        documentId: expect.stringMatching(/^csv:/),
        sheetNames: ['tiny'],
      }),
    )
    expect(publishImportedSnapshot).toHaveBeenCalledWith(
      expect.stringMatching(/^csv:/),
      expect.objectContaining({
        sheets: [
          expect.objectContaining({
            name: 'tiny',
          }),
        ],
      }),
      expect.any(Object),
    )
  })

  it('normalizes open and close workbook session lifecycle responses', async () => {
    const openResponse = await handleWorkbookAgentFrame(
      {
        kind: 'request',
        request: {
          kind: 'openWorkbookSession',
          id: 'open-1',
          documentId: 'doc-1',
          replicaId: 'replica-1',
        },
      } satisfies AgentFrame,
      {},
      {
        invalidFrameMessage: 'bad frame',
        errorCode: 'TEST_ERROR',
        loadWorkbookFile: vi.fn(),
        openWorkbookSession: async () => 'doc-1:replica-1',
        closeWorkbookSession: async () => undefined,
        getMetrics: async (request) => ({ kind: 'metrics', id: request.id, value: { ok: true } }),
      },
    )

    expect(openResponse).toEqual({
      kind: 'response',
      response: {
        kind: 'ok',
        id: 'open-1',
        sessionId: 'doc-1:replica-1',
      },
    })

    const closeResponse = await handleWorkbookAgentFrame(
      {
        kind: 'request',
        request: {
          kind: 'closeWorkbookSession',
          id: 'close-1',
          sessionId: 'doc-1:replica-1',
        },
      } satisfies AgentFrame,
      {},
      {
        invalidFrameMessage: 'bad frame',
        errorCode: 'TEST_ERROR',
        loadWorkbookFile: vi.fn(),
        openWorkbookSession: async () => 'doc-1:replica-1',
        closeWorkbookSession: async () => undefined,
        getMetrics: async (request) => ({ kind: 'metrics', id: request.id, value: { ok: true } }),
      },
    )

    expect(closeResponse).toEqual({
      kind: 'response',
      response: {
        kind: 'ok',
        id: 'close-1',
      },
    })
  })

  it('builds workbook load options without undefined keys', () => {
    const options = createWorkbookLoadOptions(
      {
        browserAppBaseUrl: 'http://127.0.0.1:3000',
      },
      {
        registerPreparedSession: vi.fn(),
        publishImportedSnapshot: vi.fn(),
      },
    )

    expect(options).toEqual({
      browserAppBaseUrl: 'http://127.0.0.1:3000',
      registerPreparedSession: expect.any(Function),
      publishImportedSnapshot: expect.any(Function),
    })
  })
})
