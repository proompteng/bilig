import { afterEach, describe, expect, it, vi } from 'vitest'

import { createEmptyPublicWorkbookManifest, validatePublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import { discoverGithubWorkbookSources, discoverRecentComplexGithubQueries } from '../public-workbook-corpus-github.ts'

describe('public workbook corpus GitHub discovery', () => {
  afterEach(() => {
    vi.restoreAllMocks()
    vi.unstubAllGlobals()
  })

  it('adds licensed recent spreadsheet files from GitHub contents search', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: string | URL) => {
        const url = String(input)
        if (url.startsWith('https://api.github.com/search/code')) {
          return jsonResponse({
            items: [
              githubSearchItem({
                repo: 'acme/budget-models',
                path: 'models/2025-budget-model.xlsx',
                name: '2025-budget-model.xlsx',
              }),
              githubSearchItem({
                repo: 'acme/budget-models',
                path: 'models/budget-model.xlsx',
                name: 'budget-model.xlsx',
              }),
              githubSearchItem({
                repo: 'acme/unlicensed-models',
                path: 'models/2026-forecast-model.xlsx',
                name: '2026-forecast-model.xlsx',
              }),
            ],
          })
        }
        if (url === 'https://api.github.com/repos/acme/budget-models/contents/models/2025-budget-model.xlsx') {
          return jsonResponse({
            download_url: 'https://raw.githubusercontent.com/acme/budget-models/main/models/2025-budget-model.xlsx',
          })
        }
        if (url === 'https://api.github.com/repos/acme/budget-models/contents/models/budget-model.xlsx') {
          return jsonResponse({
            download_url: 'https://raw.githubusercontent.com/acme/budget-models/main/models/budget-model.xlsx',
          })
        }
        if (url === 'https://api.github.com/repos/acme/unlicensed-models/contents/models/2026-forecast-model.xlsx') {
          return jsonResponse({
            download_url: 'https://raw.githubusercontent.com/acme/unlicensed-models/main/models/2026-forecast-model.xlsx',
          })
        }
        if (url === 'https://api.github.com/repos/acme/budget-models/license') {
          return jsonResponse({
            license: {
              spdx_id: 'MIT',
              name: 'MIT License',
            },
            html_url: 'https://github.com/acme/budget-models/blob/main/LICENSE',
          })
        }
        if (url === 'https://api.github.com/repos/acme/unlicensed-models/license') {
          return jsonResponse({
            license: {
              spdx_id: 'NOASSERTION',
              name: 'Other',
            },
            html_url: 'https://github.com/acme/unlicensed-models/blob/main/LICENSE',
          })
        }
        throw new Error(`Unexpected GitHub API request: ${url}`)
      }),
    )

    const manifest = await discoverGithubWorkbookSources({
      manifest: createEmptyPublicWorkbookManifest('2026-05-17T00:00:00.000Z', 500),
      queries: ['2025 budget extension:xlsx'],
      limit: 500,
      perPage: 25,
      maxPagesPerQuery: 1,
      githubToken: 'ghs_test',
      discoveredAt: '2026-05-17T00:00:00.000Z',
    })

    validatePublicWorkbookManifest(manifest)
    expect(manifest.sources).toHaveLength(1)
    expect(manifest.sources[0]).toMatchObject({
      kind: 'github-contents',
      sourceUrl: 'https://github.com/acme/budget-models/blob/main/models/2025-budget-model.xlsx',
      downloadUrl: 'https://raw.githubusercontent.com/acme/budget-models/main/models/2025-budget-model.xlsx',
      fileName: '2025-budget-model.xlsx',
      license: {
        spdxId: 'MIT',
        title: 'MIT License',
        evidenceUrl: 'https://github.com/acme/budget-models/blob/main/LICENSE',
      },
    })
    expect(manifest.sources[0]?.topicEvidence).toEqual(
      expect.arrayContaining(['recent-2025:github.path', expect.stringMatching(/^github-query:/u)]),
    )
  })

  it('walks licensed repository trees for recent spreadsheet paths', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: string | URL) => {
        const url = String(input)
        if (url.startsWith('https://api.github.com/search/repositories')) {
          return jsonResponse({
            items: [
              {
                full_name: 'acme/model-repo',
                html_url: 'https://github.com/acme/model-repo',
                default_branch: 'main',
              },
            ],
          })
        }
        if (url === 'https://api.github.com/repos/acme/model-repo/license') {
          return jsonResponse({
            license: {
              spdx_id: 'Apache-2.0',
              name: 'Apache License 2.0',
            },
            html_url: 'https://github.com/acme/model-repo/blob/main/LICENSE',
          })
        }
        if (url === 'https://api.github.com/repos/acme/model-repo/git/trees/main?recursive=1') {
          return jsonResponse({
            tree: [
              {
                type: 'blob',
                path: 'finance/2026-budget-model.xlsx',
              },
              {
                type: 'blob',
                path: 'finance/budget-model.xlsx',
              },
            ],
          })
        }
        throw new Error(`Unexpected GitHub API request: ${url}`)
      }),
    )

    const manifest = await discoverRecentComplexGithubQueries({
      manifest: createEmptyPublicWorkbookManifest('2026-05-17T00:00:00.000Z', 500),
      queries: [],
      repositoryQueries: ['2026 xlsx license:apache-2.0'],
      limit: 500,
      perPage: 25,
      maxPagesPerQuery: 1,
      maxRepositoriesPerQuery: 5,
      githubToken: 'ghs_test',
      discoveredAt: '2026-05-17T00:00:00.000Z',
    })

    validatePublicWorkbookManifest(manifest)
    expect(manifest.sources).toHaveLength(1)
    expect(manifest.sources[0]).toMatchObject({
      kind: 'github-contents',
      sourceUrl: 'https://github.com/acme/model-repo/blob/main/finance/2026-budget-model.xlsx',
      downloadUrl: 'https://raw.githubusercontent.com/acme/model-repo/main/finance/2026-budget-model.xlsx',
      fileName: '2026-budget-model.xlsx',
      license: {
        spdxId: 'Apache-2.0',
        title: 'Apache License 2.0',
        evidenceUrl: 'https://github.com/acme/model-repo/blob/main/LICENSE',
      },
    })
    expect(manifest.sources[0]?.topicEvidence).toEqual(
      expect.arrayContaining(['recent-2026:github.path', expect.stringMatching(/^github-repo-query:/u)]),
    )
  })

  it('uses per-workbook GitHub commit dates as recent evidence when paths do not carry years', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: string | URL) => {
        const url = String(input)
        if (url.startsWith('https://api.github.com/search/repositories')) {
          return jsonResponse({
            items: [
              {
                full_name: 'acme/finance-models',
                html_url: 'https://github.com/acme/finance-models',
                default_branch: 'main',
              },
            ],
          })
        }
        if (url === 'https://api.github.com/repos/acme/finance-models/license') {
          return jsonResponse({
            license: {
              spdx_id: 'MIT',
              name: 'MIT License',
            },
            html_url: 'https://github.com/acme/finance-models/blob/main/LICENSE',
          })
        }
        if (url === 'https://api.github.com/repos/acme/finance-models/git/trees/main?recursive=1') {
          return jsonResponse({
            tree: [
              {
                type: 'blob',
                path: 'models/dcf-model.xlsx',
              },
              {
                type: 'blob',
                path: 'exports/raw-data.xlsx',
              },
            ],
          })
        }
        if (url.startsWith('https://api.github.com/repos/acme/finance-models/commits?') && url.includes('path=models%2Fdcf-model.xlsx')) {
          return jsonResponse([
            {
              commit: {
                committer: {
                  date: '2026-04-30T12:00:00Z',
                },
              },
            },
          ])
        }
        if (url.startsWith('https://api.github.com/repos/acme/finance-models/commits?') && url.includes('path=exports%2Fraw-data.xlsx')) {
          return jsonResponse([
            {
              commit: {
                committer: {
                  date: '2024-12-31T12:00:00Z',
                },
              },
            },
          ])
        }
        throw new Error(`Unexpected GitHub API request: ${url}`)
      }),
    )

    const manifest = await discoverRecentComplexGithubQueries({
      manifest: createEmptyPublicWorkbookManifest('2026-05-17T00:00:00.000Z', 500),
      queries: [],
      repositoryQueries: ['financial model excel license:mit'],
      limit: 500,
      perPage: 25,
      maxPagesPerQuery: 1,
      maxRepositoriesPerQuery: 5,
      githubToken: 'ghs_test',
      discoveredAt: '2026-05-17T00:00:00.000Z',
    })

    validatePublicWorkbookManifest(manifest)
    expect(manifest.sources).toHaveLength(1)
    expect(manifest.sources[0]).toMatchObject({
      sourceUrl: 'https://github.com/acme/finance-models/blob/main/models/dcf-model.xlsx',
      downloadUrl: 'https://raw.githubusercontent.com/acme/finance-models/main/models/dcf-model.xlsx',
      fileName: 'dcf-model.xlsx',
    })
    expect(manifest.sources[0]?.topicEvidence).toEqual(
      expect.arrayContaining(['recent-2026:github.commitDate', expect.stringMatching(/^github-repo-query:/u)]),
    )
  })

  it('uses repository names and descriptions as recent evidence before per-file commit lookups', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: string | URL) => {
        const url = String(input)
        if (url.startsWith('https://api.github.com/search/repositories')) {
          return jsonResponse({
            items: [
              {
                full_name: 'acme/forecast-models-2025',
                name: 'forecast-models-2025',
                description: 'Open financial model workbooks for 2025 planning',
                html_url: 'https://github.com/acme/forecast-models-2025',
                default_branch: 'main',
              },
            ],
          })
        }
        if (url === 'https://api.github.com/repos/acme/forecast-models-2025/license') {
          return jsonResponse({
            license: {
              spdx_id: 'MIT',
              name: 'MIT License',
            },
            html_url: 'https://github.com/acme/forecast-models-2025/blob/main/LICENSE',
          })
        }
        if (url === 'https://api.github.com/repos/acme/forecast-models-2025/git/trees/main?recursive=1') {
          return jsonResponse({
            tree: [
              {
                type: 'blob',
                path: 'models/operating-model.xlsx',
              },
            ],
          })
        }
        throw new Error(`Unexpected GitHub API request: ${url}`)
      }),
    )

    const manifest = await discoverRecentComplexGithubQueries({
      manifest: createEmptyPublicWorkbookManifest('2026-05-17T00:00:00.000Z', 500),
      queries: [],
      repositoryQueries: ['financial model excel license:mit'],
      limit: 500,
      perPage: 25,
      maxPagesPerQuery: 1,
      maxRepositoriesPerQuery: 5,
      githubToken: 'ghs_test',
      discoveredAt: '2026-05-17T00:00:00.000Z',
    })

    validatePublicWorkbookManifest(manifest)
    expect(manifest.sources).toHaveLength(1)
    expect(manifest.sources[0]).toMatchObject({
      sourceUrl: 'https://github.com/acme/forecast-models-2025/blob/main/models/operating-model.xlsx',
      downloadUrl: 'https://raw.githubusercontent.com/acme/forecast-models-2025/main/models/operating-model.xlsx',
      fileName: 'operating-model.xlsx',
    })
    expect(manifest.sources[0]?.topicEvidence).toEqual(
      expect.arrayContaining(['recent-2025:github.repositoryFullName', 'recent-2025:github.repositoryDescription']),
    )
  })
})

function githubSearchItem(args: { readonly repo: string; readonly path: string; readonly name: string }): Record<string, unknown> {
  return {
    name: args.name,
    path: args.path,
    html_url: `https://github.com/${args.repo}/blob/main/${args.path}`,
    url: `https://api.github.com/repos/${args.repo}/contents/${args.path}`,
    repository: {
      full_name: args.repo,
    },
  }
}

function jsonResponse(value: unknown): Response {
  return new Response(`${JSON.stringify(value)}\n`, {
    status: 200,
    headers: {
      'content-type': 'application/json',
    },
  })
}
