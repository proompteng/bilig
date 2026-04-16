// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { WorkbookLoadedResponse } from '@bilig/agent-api'
import { resolveWorkbookNavigationUrl } from '../workbook-navigation.js'
import { resolveImportedWorkbookNavigationUrl } from '../workbook-import-client.js'

afterEach(() => {
  vi.restoreAllMocks()
  window.history.replaceState({}, '', '/?existing=1')
})

describe('workbook navigation', () => {
  it('builds workbook navigation urls with normalized cell addresses and optional server routing', () => {
    window.history.replaceState({}, '', '/?existing=1')
    const url = new URL(
      resolveWorkbookNavigationUrl({
        documentId: 'doc-42',
        sheetName: 'Summary',
        address: 'b7',
        serverUrl: 'https://sync.bilig.test',
      }),
    )

    expect(url.origin).toBe(window.location.origin)
    expect(url.searchParams.get('existing')).toBe('1')
    expect(url.searchParams.get('document')).toBe('doc-42')
    expect(url.searchParams.get('server')).toBe('https://sync.bilig.test')
    expect(url.searchParams.get('sheet')).toBe('Summary')
    expect(url.searchParams.get('cell')).toBe('B7')
  })

  it('removes optional workbook navigation parameters when the target omits them', () => {
    window.history.replaceState({}, '', '/?existing=1&server=https%3A%2F%2Fold-sync&sheet=Old&cell=Z99')
    const url = new URL(
      resolveWorkbookNavigationUrl({
        documentId: 'doc-42',
      }),
    )

    expect(url.origin).toBe(window.location.origin)
    expect(url.searchParams.get('existing')).toBe('1')
    expect(url.searchParams.get('document')).toBe('doc-42')
    expect(url.searchParams.has('server')).toBe(false)
    expect(url.searchParams.has('sheet')).toBe(false)
    expect(url.searchParams.has('cell')).toBe(false)
  })

  it('routes imported workbooks through the shared workbook navigation helper', () => {
    const fallbackResult: WorkbookLoadedResponse = {
      kind: 'workbookLoaded',
      id: 'load-1',
      documentId: 'csv:abc123',
      sessionId: 'csv:abc123:browser-import',
      workbookName: 'metrics',
      sheetNames: ['metrics'],
      serverUrl: 'https://sync.bilig.test',
      warnings: [],
    }
    const browserResult: WorkbookLoadedResponse = {
      ...fallbackResult,
      browserUrl: 'https://sync.bilig.test/?document=csv%3Aabc123',
    }

    const fallbackUrl = new URL(resolveImportedWorkbookNavigationUrl(fallbackResult))

    expect(fallbackUrl.origin).toBe(window.location.origin)
    expect(fallbackUrl.searchParams.get('existing')).toBe('1')
    expect(fallbackUrl.searchParams.get('document')).toBe('csv:abc123')
    expect(fallbackUrl.searchParams.get('server')).toBe('https://sync.bilig.test')
    expect(resolveImportedWorkbookNavigationUrl(browserResult)).toBe('https://sync.bilig.test/?document=csv%3Aabc123')
  })
})
