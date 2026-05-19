import { afterEach, describe, expect, it, vi } from 'vitest'

import { createEmptyPublicWorkbookManifest, validatePublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import { discoverRecentComplexZenodoQueries } from '../public-workbook-corpus-zenodo.ts'

describe('public workbook corpus Zenodo discovery', () => {
  afterEach(() => {
    vi.restoreAllMocks()
    vi.unstubAllGlobals()
  })

  it('adds licensed recent spreadsheet files from Zenodo records', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: string | URL) => {
        const url = String(input)
        if (url.startsWith('https://zenodo.org/api/records')) {
          return jsonResponse({
            hits: {
              hits: [
                {
                  id: 17658060,
                  links: {
                    html: 'https://zenodo.org/records/17658060',
                  },
                  metadata: {
                    title: '2025 forecast workbook',
                    publication_date: '2025-11-20',
                    license: {
                      id: 'cc-by-4.0',
                    },
                  },
                  files: [
                    {
                      key: 'forecast-workbook.xlsx',
                      links: {
                        self: 'https://zenodo.org/api/records/17658060/files/forecast-workbook.xlsx/content',
                      },
                    },
                    {
                      key: 'paper.pdf',
                      links: {
                        self: 'https://zenodo.org/api/records/17658060/files/paper.pdf/content',
                      },
                    },
                  ],
                },
                {
                  id: 17658061,
                  links: {
                    html: 'https://zenodo.org/records/17658061',
                  },
                  metadata: {
                    title: '2025 closed workbook',
                    publication_date: '2025-11-20',
                    license: {
                      id: 'other-closed',
                    },
                  },
                  files: [
                    {
                      key: 'closed-workbook.xlsx',
                      links: {
                        self: 'https://zenodo.org/api/records/17658061/files/closed-workbook.xlsx/content',
                      },
                    },
                  ],
                },
              ],
            },
          })
        }
        throw new Error(`Unexpected Zenodo API request: ${url}`)
      }),
    )

    const manifest = await discoverRecentComplexZenodoQueries({
      manifest: createEmptyPublicWorkbookManifest('2026-05-19T00:00:00.000Z', 500),
      queries: ['2025 xlsx forecast'],
      limit: 500,
      perPage: 25,
      maxPagesPerQuery: 1,
      discoveredAt: '2026-05-19T00:00:00.000Z',
    })

    validatePublicWorkbookManifest(manifest)
    expect(manifest.sources).toHaveLength(1)
    expect(manifest.sources[0]).toMatchObject({
      kind: 'direct-url',
      sourceUrl: 'https://zenodo.org/records/17658060',
      downloadUrl: 'https://zenodo.org/api/records/17658060/files/forecast-workbook.xlsx/content',
      fileName: 'forecast-workbook.xlsx',
      license: {
        spdxId: 'CC-BY-4.0',
        title: 'cc-by-4.0',
        evidenceUrl: 'https://zenodo.org/records/17658060',
      },
    })
    expect(manifest.sources[0]?.topicEvidence).toEqual(
      expect.arrayContaining(['recent-2025:zenodo.title', 'recent-2025:zenodo.publicationDate', expect.stringMatching(/^zenodo-query:/u)]),
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
