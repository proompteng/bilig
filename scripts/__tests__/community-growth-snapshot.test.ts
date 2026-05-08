import { describe, expect, it } from 'vitest'

import { collectCommunityGrowthSnapshot } from '../community-growth-snapshot.ts'

function responseJson(value: unknown): Response {
  return Response.json(value)
}

function fetchInputUrl(input: string | URL | Request): string {
  if (typeof input === 'string') {
    return input
  }

  if (input instanceof URL) {
    return input.href
  }

  return input.url
}

describe('community growth snapshot', () => {
  it('collects public GitHub and npm metrics without traffic when no token is configured', async () => {
    const calls: string[] = []
    const fetchImpl = async (input: string | URL | Request): Promise<Response> => {
      const url = fetchInputUrl(input)
      calls.push(url)

      if (url === 'https://api.github.com/repos/proompteng/bilig') {
        return responseJson({
          full_name: 'proompteng/bilig',
          html_url: 'https://github.com/proompteng/bilig',
          description: 'Headless spreadsheet engine',
          stargazers_count: 8,
          forks_count: 1,
          subscribers_count: 0,
          open_issues_count: 39,
          default_branch: 'main',
        })
      }

      if (url === 'https://registry.npmjs.org/%40bilig%2Fheadless') {
        return responseJson({
          name: '@bilig/headless',
          description: 'Headless spreadsheet engine',
          license: 'MIT',
          'dist-tags': {
            latest: '0.10.62',
          },
          time: {
            modified: '2026-05-08T05:53:15.971Z',
          },
        })
      }

      if (url === 'https://api.npmjs.org/downloads/point/last-week/%40bilig%2Fheadless') {
        return responseJson({
          downloads: 2399,
          start: '2026-04-30',
          end: '2026-05-06',
          package: '@bilig/headless',
        })
      }

      if (url === 'https://api.npmjs.org/downloads/point/last-month/%40bilig%2Fheadless') {
        return responseJson({
          downloads: 16491,
          start: '2026-04-07',
          end: '2026-05-06',
          package: '@bilig/headless',
        })
      }

      if (url.startsWith('https://api.github.com/search/issues?')) {
        const query = new URL(url).searchParams.get('q')
        if (query === 'repo:proompteng/bilig is:issue is:open label:"good first issue"') {
          return responseJson({ total_count: 8 })
        }
        if (query === 'repo:proompteng/bilig is:issue is:open label:first-timers-only') {
          return responseJson({ total_count: 5 })
        }
        if (query === 'repo:proompteng/bilig is:issue is:open label:"help wanted"') {
          return responseJson({ total_count: 8 })
        }
        if (query === 'repo:proompteng/bilig is:pr is:open') {
          return responseJson({ total_count: 2 })
        }
      }

      throw new Error(`unexpected fetch ${url}`)
    }

    const snapshot = await collectCommunityGrowthSnapshot({
      fetchImpl: fetchImpl as typeof fetch,
      githubToken: '',
      now: new Date('2026-05-08T06:00:00.000Z'),
    })

    expect(snapshot).toMatchObject({
      capturedAt: '2026-05-08T06:00:00.000Z',
      github: {
        fullName: 'proompteng/bilig',
        stargazerCount: 8,
        forkCount: 1,
        openIssueCount: 39,
      },
      npm: {
        name: '@bilig/headless',
        version: '0.10.62',
        downloads: {
          lastWeek: {
            downloads: 2399,
          },
          lastMonth: {
            downloads: 16491,
          },
        },
      },
      contributorFunnel: {
        openGoodFirstIssueCount: 8,
        openFirstTimersOnlyIssueCount: 5,
        openHelpWantedIssueCount: 8,
        openPullRequestCount: 2,
      },
      traffic: {
        available: false,
      },
    })
    expect(calls).not.toContain('https://api.github.com/repos/proompteng/bilig/traffic/views')
  })

  it('collects GitHub traffic metrics when a token is provided', async () => {
    const fetchImpl = async (input: string | URL | Request): Promise<Response> => {
      const url = fetchInputUrl(input)

      if (url === 'https://api.github.com/repos/proompteng/bilig') {
        return responseJson({
          full_name: 'proompteng/bilig',
          html_url: 'https://github.com/proompteng/bilig',
          description: 'Headless spreadsheet engine',
          stargazers_count: 8,
          forks_count: 1,
          subscribers_count: 0,
          open_issues_count: 39,
          default_branch: 'main',
        })
      }

      if (url === 'https://registry.npmjs.org/%40bilig%2Fheadless') {
        return responseJson({
          name: '@bilig/headless',
          description: 'Headless spreadsheet engine',
          license: 'MIT',
          'dist-tags': {
            latest: '0.10.62',
          },
          time: {
            modified: '2026-05-08T05:53:15.971Z',
          },
        })
      }

      if (url === 'https://api.npmjs.org/downloads/point/last-week/%40bilig%2Fheadless') {
        return responseJson({ downloads: 1, start: '2026-04-30', end: '2026-05-06' })
      }

      if (url === 'https://api.npmjs.org/downloads/point/last-month/%40bilig%2Fheadless') {
        return responseJson({ downloads: 2, start: '2026-04-07', end: '2026-05-06' })
      }

      if (url.startsWith('https://api.github.com/search/issues?')) {
        const query = new URL(url).searchParams.get('q')
        if (query === 'repo:proompteng/bilig is:issue is:open label:"good first issue"') {
          return responseJson({ total_count: 8 })
        }
        if (query === 'repo:proompteng/bilig is:issue is:open label:first-timers-only') {
          return responseJson({ total_count: 5 })
        }
        if (query === 'repo:proompteng/bilig is:issue is:open label:"help wanted"') {
          return responseJson({ total_count: 8 })
        }
        if (query === 'repo:proompteng/bilig is:pr is:open') {
          return responseJson({ total_count: 2 })
        }
      }

      if (url === 'https://api.github.com/repos/proompteng/bilig/traffic/views') {
        return responseJson({ count: 120, uniques: 33 })
      }

      if (url === 'https://api.github.com/repos/proompteng/bilig/traffic/clones') {
        return responseJson({ count: 45, uniques: 12 })
      }

      if (url === 'https://api.github.com/repos/proompteng/bilig/traffic/popular/referrers') {
        return responseJson([{ referrer: 'news.ycombinator.com', count: 20, uniques: 10 }])
      }

      if (url === 'https://api.github.com/repos/proompteng/bilig/traffic/popular/paths') {
        return responseJson([{ path: '/proompteng/bilig', title: 'bilig', count: 50, uniques: 25 }])
      }

      throw new Error(`unexpected fetch ${url}`)
    }

    const snapshot = await collectCommunityGrowthSnapshot({
      fetchImpl: fetchImpl as typeof fetch,
      githubToken: 'test-token',
      now: new Date('2026-05-08T06:00:00.000Z'),
    })

    expect(snapshot.traffic).toEqual({
      available: true,
      views: {
        count: 120,
        uniques: 33,
      },
      clones: {
        count: 45,
        uniques: 12,
      },
      referrers: [
        {
          referrer: 'news.ycombinator.com',
          count: 20,
          uniques: 10,
        },
      ],
      paths: [
        {
          path: '/proompteng/bilig',
          title: 'bilig',
          count: 50,
          uniques: 25,
        },
      ],
    })
  })
})
