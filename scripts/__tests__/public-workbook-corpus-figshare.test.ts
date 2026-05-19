import { afterEach, describe, expect, it, vi } from 'vitest'

import { createEmptyPublicWorkbookManifest, validatePublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import { discoverRecentComplexFigshareQueries } from '../public-workbook-corpus-figshare.ts'

describe('public workbook corpus Figshare discovery', () => {
  afterEach(() => {
    vi.restoreAllMocks()
    vi.unstubAllGlobals()
  })

  it('adds licensed recent spreadsheet files from Figshare articles', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: string | URL) => {
        const url = String(input)
        if (url === 'https://api.figshare.com/v2/articles/search') {
          return jsonResponse([
            {
              id: 32080353,
              url_public_api: 'https://api.figshare.com/v2/articles/32080353',
            },
            {
              id: 32080354,
              url_public_api: 'https://api.figshare.com/v2/articles/32080354',
            },
          ])
        }
        if (url === 'https://api.figshare.com/v2/articles/32080353') {
          return jsonResponse({
            id: 32080353,
            title: '2026 forecast workbook',
            url_public_html: 'https://figshare.com/articles/dataset/forecast_workbook/32080353',
            published_date: '2026-03-20T00:00:00Z',
            license: {
              name: 'CC BY 4.0',
              url: 'https://creativecommons.org/licenses/by/4.0/',
            },
            files: [
              {
                name: 'forecast-workbook.xlsx',
                download_url: 'https://ndownloader.figshare.com/files/61234567',
                is_link_only: false,
              },
              {
                name: 'paper.pdf',
                download_url: 'https://ndownloader.figshare.com/files/61234568',
                is_link_only: false,
              },
            ],
          })
        }
        if (url === 'https://api.figshare.com/v2/articles/32080354') {
          return jsonResponse({
            id: 32080354,
            title: '2026 closed workbook',
            url_public_html: 'https://figshare.com/articles/dataset/closed_workbook/32080354',
            published_date: '2026-03-20T00:00:00Z',
            license: {
              name: 'All rights reserved',
              url: 'https://figshare.com/articles/dataset/closed_workbook/32080354',
            },
            files: [
              {
                name: 'closed-workbook.xlsx',
                download_url: 'https://ndownloader.figshare.com/files/61234569',
                is_link_only: false,
              },
            ],
          })
        }
        throw new Error(`Unexpected Figshare API request: ${url}`)
      }),
    )

    const manifest = await discoverRecentComplexFigshareQueries({
      manifest: createEmptyPublicWorkbookManifest('2026-05-19T00:00:00.000Z', 500),
      queries: ['2026 xlsx forecast'],
      limit: 500,
      pageSize: 25,
      maxPagesPerQuery: 1,
      discoveredAt: '2026-05-19T00:00:00.000Z',
    })

    validatePublicWorkbookManifest(manifest)
    expect(manifest.sources).toHaveLength(1)
    expect(manifest.sources[0]).toMatchObject({
      kind: 'direct-url',
      sourceUrl: 'https://figshare.com/articles/dataset/forecast_workbook/32080353',
      downloadUrl: 'https://ndownloader.figshare.com/files/61234567',
      fileName: 'forecast-workbook.xlsx',
      license: {
        spdxId: 'CC-BY-4.0',
        title: 'CC BY 4.0',
        evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
      },
    })
    expect(manifest.sources[0]?.topicEvidence).toEqual(
      expect.arrayContaining([
        'recent-2026:figshare.title',
        'recent-2026:figshare.publishedDate',
        expect.stringMatching(/^figshare-query:/u),
      ]),
    )
  })
})

function jsonResponse(value: unknown): Response {
  return new Response(`${JSON.stringify(value)}\n`, {
    status: 200,
    headers: {
      'content-type': 'application/json',
    },
  })
}
