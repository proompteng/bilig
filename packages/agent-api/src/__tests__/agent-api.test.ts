import { describe, expect, it } from 'vitest'

import {
  CSV_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  decodeAgentFrame,
  decodeStdioMessages,
  encodeAgentFrame,
  encodeStdioMessage,
  normalizeWorkbookImportContentType,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
} from '../index.js'

describe('agent api', () => {
  it('roundtrips request and response envelopes', () => {
    const frame = {
      kind: 'request' as const,
      request: {
        kind: 'openWorkbookSession' as const,
        id: 'req-1',
        documentId: 'book-1',
        replicaId: 'agent-local',
      },
    }

    expect(decodeAgentFrame(encodeAgentFrame(frame))).toEqual(frame)
  })

  it('decodes multiple stdio messages from one buffer', () => {
    const encoded = new Uint8Array([
      ...encodeStdioMessage({
        kind: 'request',
        request: {
          kind: 'getMetrics',
          id: 'req-1',
          sessionId: 'sess-1',
        },
      }),
      ...encodeStdioMessage({
        kind: 'response',
        response: {
          kind: 'ok',
          id: 'req-1',
          value: { ok: true },
        },
      }),
    ])

    const decoded = decodeStdioMessages(encoded)
    expect(decoded.frames).toHaveLength(2)
    expect(decoded.remainder.byteLength).toBe(0)
  })

  it('roundtrips workbook file load requests and responses', () => {
    const requestFrame = {
      kind: 'request' as const,
      request: {
        kind: 'loadWorkbookFile' as const,
        id: 'upload-1',
        replicaId: 'agent-local',
        openMode: 'create' as const,
        fileName: 'report.xlsx',
        contentType: XLSX_CONTENT_TYPE,
        bytesBase64: 'QUJD',
      },
    }
    expect(decodeAgentFrame(encodeAgentFrame(requestFrame))).toEqual(requestFrame)

    const csvRequestFrame = {
      kind: 'request' as const,
      request: {
        kind: 'loadWorkbookFile' as const,
        id: 'upload-2',
        replicaId: 'agent-local',
        openMode: 'replace' as const,
        documentId: 'doc-1',
        fileName: 'report.csv',
        contentType: CSV_CONTENT_TYPE,
        bytesBase64: 'QUJD',
      },
    }
    expect(decodeAgentFrame(encodeAgentFrame(csvRequestFrame))).toEqual(csvRequestFrame)

    const csvRequestFrameWithUploadMetadata = {
      kind: 'request' as const,
      request: {
        kind: 'loadWorkbookFile' as const,
        id: 'upload-3',
        replicaId: 'agent-local',
        openMode: 'create' as const,
        fileName: 'report.csv',
        contentType: 'TEXT/CSV; charset=utf-8',
        bytesBase64: 'QUJD',
      },
    }
    expect(decodeAgentFrame(encodeAgentFrame(csvRequestFrameWithUploadMetadata))).toEqual(csvRequestFrameWithUploadMetadata)

    const responseFrame = {
      kind: 'response' as const,
      response: {
        kind: 'workbookLoaded' as const,
        id: 'upload-1',
        documentId: 'xlsx:abc123',
        sessionId: 'xlsx:abc123:agent-local',
        workbookName: 'report.xlsx',
        sheetNames: ['Sheet1'],
        serverUrl: 'http://127.0.0.1:4321',
        browserUrl: 'http://127.0.0.1:4173/?document=xlsx%3Aabc123',
        warnings: [],
      },
    }
    expect(decodeAgentFrame(encodeAgentFrame(responseFrame))).toEqual(responseFrame)
  })

  it('normalizes workbook import content types from upload metadata', () => {
    expect(normalizeWorkbookImportContentType(' text/csv; charset=utf-8 ')).toBe(CSV_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('TEXT/CSV')).toBe(CSV_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType(`${XLSX_CONTENT_TYPE}; charset=binary`)).toBe(XLSX_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType(XLSX_CONTENT_TYPE.toUpperCase())).toBe(XLSX_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('application/vnd.ms-excel.sheet.macroEnabled.12')).toBe(XLSM_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('application/vnd.ms-excel.sheet.binary.macroEnabled.12')).toBe(XLSB_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('application/vnd.ms-excel; charset=binary')).toBe(LEGACY_XLS_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('application/octet-stream')).toBeNull()
  })
})
