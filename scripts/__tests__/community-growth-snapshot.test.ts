import { describe, expect, it } from 'vitest'

import {
  collectCommunityGrowthSnapshot,
  renderCommunityGrowthSnapshotMarkdown,
  type GitHubCliApiJson,
} from '../community-growth-snapshot.ts'

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

function requestBodyQuery(init: RequestInit | undefined): string {
  const body = init?.body
  if (typeof body !== 'string') {
    throw new Error('expected GraphQL request body to be a string')
  }

  const parsed: unknown = JSON.parse(body)
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('expected GraphQL request body to be an object')
  }

  const query = Reflect.get(parsed, 'query')
  if (typeof query !== 'string') {
    throw new Error('expected GraphQL request body query to be a string')
  }

  return query
}

function growthSearchResponse(url: string): Response | undefined {
  if (!url.startsWith('https://api.github.com/search/issues?')) {
    return undefined
  }

  const query = new URL(url).searchParams.get('q')

  if (query === 'repo:proompteng/bilig is:issue is:open label:"good first issue"') {
    return responseJson({ total_count: 8 })
  }
  if (query === 'repo:proompteng/bilig is:issue is:open label:first-timers-only') {
    return responseJson({ total_count: 5 })
  }
  if (query === 'repo:proompteng/bilig is:issue is:open label:first-timers-only label:documentation') {
    return responseJson({ total_count: 4 })
  }
  if (query === 'repo:proompteng/bilig is:issue is:open label:first-timers-only -label:documentation') {
    return responseJson({ total_count: 1 })
  }
  if (query === 'repo:proompteng/bilig is:issue is:open label:"help wanted"') {
    return responseJson({ total_count: 8 })
  }
  if (query === 'repo:proompteng/bilig is:pr is:open') {
    return responseJson({ total_count: 2 })
  }
  if (query === 'repo:proompteng/bilig is:issue is:open -author:gregkonush') {
    return responseJson({ total_count: 3 })
  }
  if (query === 'repo:proompteng/bilig is:pr is:open -author:gregkonush') {
    return responseJson({ total_count: 1 })
  }
  if (
    query === 'repo:proompteng/bilig is:issue created:>=2026-05-01 -author:gregkonush' ||
    query === 'repo:proompteng/bilig is:issue created:>=2026-05-05 -author:gregkonush'
  ) {
    return responseJson({ total_count: 2 })
  }
  if (
    query === 'repo:proompteng/bilig is:pr created:>=2026-05-01 -author:gregkonush' ||
    query === 'repo:proompteng/bilig is:pr created:>=2026-05-05 -author:gregkonush'
  ) {
    return responseJson({ total_count: 1 })
  }

  return undefined
}

describe('community growth snapshot', () => {
  it('collects public GitHub and npm metrics without traffic when no token or gh CLI is configured', async () => {
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
          topics: ['spreadsheet', 'typescript'],
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

      const searchResponse = growthSearchResponse(url)
      if (searchResponse !== undefined) {
        return searchResponse
      }

      throw new Error(`unexpected fetch ${url}`)
    }

    const snapshot = await collectCommunityGrowthSnapshot({
      fetchImpl: fetchImpl as typeof fetch,
      githubToken: '',
      githubCliApiJson: false,
      now: new Date('2026-05-08T06:00:00.000Z'),
    })

    expect(snapshot).toMatchObject({
      capturedAt: '2026-05-08T06:00:00.000Z',
      github: {
        fullName: 'proompteng/bilig',
        stargazerCount: 8,
        forkCount: 1,
        openIssueCount: 37,
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
        openDocumentationStarterIssueCount: 4,
        openNonDocumentationStarterIssueCount: 1,
        openHelpWantedIssueCount: 8,
        openPullRequestCount: 2,
        externalOpenIssueCount: 3,
        externalOpenPullRequestCount: 1,
        externalIssuesOpenedLastSevenDays: 2,
        externalPullRequestsOpenedLastSevenDays: 1,
      },
      discussionActivity: {
        available: false,
      },
      traffic: {
        available: false,
      },
    })
    expect(calls).not.toContain('https://api.github.com/repos/proompteng/bilig/traffic/views')
    expect(calls).not.toContain('https://api.github.com/graphql')
  })

  it('uses authenticated gh CLI fallback for discussion and traffic metrics when no token is configured', async () => {
    const calls: string[] = []
    const cliCalls: Array<{ readonly endpoint: string; readonly fields: readonly string[] }> = []
    const fetchImpl = async (input: string | URL | Request): Promise<Response> => {
      const url = fetchInputUrl(input)
      calls.push(url)

      if (url === 'https://api.github.com/repos/proompteng/bilig') {
        return responseJson({
          full_name: 'proompteng/bilig',
          html_url: 'https://github.com/proompteng/bilig',
          description: 'Headless spreadsheet engine',
          stargazers_count: 24,
          forks_count: 4,
          subscribers_count: 1,
          open_issues_count: 29,
          default_branch: 'main',
          topics: ['headless-spreadsheet', 'typescript'],
        })
      }

      if (url === 'https://registry.npmjs.org/%40bilig%2Fheadless') {
        return responseJson({
          name: '@bilig/headless',
          description: 'Headless spreadsheet engine',
          license: 'MIT',
          'dist-tags': {
            latest: '0.11.24',
          },
          time: {
            modified: '2026-05-12T05:53:15.971Z',
          },
        })
      }

      if (url === 'https://api.npmjs.org/downloads/point/last-week/%40bilig%2Fheadless') {
        return responseJson({ downloads: 13427, start: '2026-05-05', end: '2026-05-11' })
      }

      if (url === 'https://api.npmjs.org/downloads/point/last-month/%40bilig%2Fheadless') {
        return responseJson({ downloads: 24931, start: '2026-04-12', end: '2026-05-11' })
      }

      const searchResponse = growthSearchResponse(url)
      if (searchResponse !== undefined) {
        return searchResponse
      }

      throw new Error(`unexpected fetch ${url}`)
    }
    const githubCliApiJson: GitHubCliApiJson = async (endpoint, options) => {
      cliCalls.push({
        endpoint,
        fields: (options?.fields ?? []).map((field) => field.name),
      })

      if (endpoint === 'repos/proompteng/bilig/traffic/views') {
        return { count: 393, uniques: 159 }
      }
      if (endpoint === 'repos/proompteng/bilig/traffic/clones') {
        return { count: 18287, uniques: 1907 }
      }
      if (endpoint === 'repos/proompteng/bilig/traffic/popular/referrers') {
        return [{ referrer: 'news.ycombinator.com', count: 51, uniques: 36 }]
      }
      if (endpoint === 'repos/proompteng/bilig/traffic/popular/paths') {
        return [{ path: '/proompteng/bilig', title: 'bilig', count: 199, uniques: 145 }]
      }
      if (endpoint === 'graphql') {
        const query = options?.fields?.find((field) => field.name === 'query')?.value ?? ''
        if (query.includes('CommunityGrowthContributorFunnel')) {
          return {
            data: {
              goodFirst: {
                issueCount: 22,
              },
              firstTimers: {
                issueCount: 22,
              },
              documentationStarters: {
                issueCount: 20,
              },
              nonDocumentationStarters: {
                issueCount: 2,
              },
              helpWanted: {
                issueCount: 22,
              },
              openPullRequests: {
                issueCount: 0,
              },
              externalOpenIssues: {
                issueCount: 1,
              },
              externalOpenPullRequests: {
                issueCount: 0,
              },
              externalRecentIssues: {
                issueCount: 22,
              },
              externalRecentPullRequests: {
                issueCount: 5,
              },
            },
          }
        }

        if (!query.includes('CommunityGrowthDiscussions')) {
          throw new Error(`unexpected gh GraphQL query ${query}`)
        }

        return {
          data: {
            repository: {
              discussions: {
                totalCount: 4,
                nodes: [
                  {
                    number: 213,
                    title: 'Five Node workbook automation examples',
                    url: 'https://github.com/proompteng/bilig/discussions/213',
                    category: {
                      name: 'Show and tell',
                    },
                    createdAt: '2026-05-12T21:46:51Z',
                    updatedAt: '2026-05-12T22:10:00Z',
                    comments: {
                      totalCount: 1,
                    },
                  },
                ],
              },
            },
          },
        }
      }

      throw new Error(`unexpected gh api ${endpoint}`)
    }

    const snapshot = await collectCommunityGrowthSnapshot({
      fetchImpl: fetchImpl as typeof fetch,
      githubToken: '',
      githubCliApiJson,
      now: new Date('2026-05-12T22:14:21.495Z'),
    })

    expect(calls).not.toContain('https://api.github.com/repos/proompteng/bilig/traffic/views')
    expect(calls).not.toContain('https://api.github.com/graphql')
    expect(cliCalls).toEqual([
      {
        endpoint: 'graphql',
        fields: [
          'query',
          'goodFirst',
          'firstTimers',
          'documentationStarters',
          'nonDocumentationStarters',
          'helpWanted',
          'openPullRequests',
          'externalOpenIssues',
          'externalOpenPullRequests',
          'externalRecentIssues',
          'externalRecentPullRequests',
        ],
      },
      {
        endpoint: 'graphql',
        fields: ['query', 'owner', 'repo'],
      },
      {
        endpoint: 'repos/proompteng/bilig/traffic/views',
        fields: [],
      },
      {
        endpoint: 'repos/proompteng/bilig/traffic/clones',
        fields: [],
      },
      {
        endpoint: 'repos/proompteng/bilig/traffic/popular/referrers',
        fields: [],
      },
      {
        endpoint: 'repos/proompteng/bilig/traffic/popular/paths',
        fields: [],
      },
    ])
    expect(snapshot.contributorFunnel).toEqual({
      openGoodFirstIssueCount: 22,
      openFirstTimersOnlyIssueCount: 22,
      openDocumentationStarterIssueCount: 20,
      openNonDocumentationStarterIssueCount: 2,
      openHelpWantedIssueCount: 22,
      openPullRequestCount: 0,
      externalOpenIssueCount: 1,
      externalOpenPullRequestCount: 0,
      externalIssuesOpenedLastSevenDays: 22,
      externalPullRequestsOpenedLastSevenDays: 5,
    })
    expect(snapshot.discussionActivity).toEqual({
      available: true,
      totalCount: 4,
      recent: [
        {
          number: 213,
          title: 'Five Node workbook automation examples',
          url: 'https://github.com/proompteng/bilig/discussions/213',
          category: 'Show and tell',
          createdAt: '2026-05-12T21:46:51Z',
          updatedAt: '2026-05-12T22:10:00Z',
          commentCount: 1,
        },
      ],
    })
    expect(snapshot.traffic).toEqual({
      available: true,
      views: {
        count: 393,
        uniques: 159,
      },
      clones: {
        count: 18287,
        uniques: 1907,
      },
      referrers: [
        {
          referrer: 'news.ycombinator.com',
          count: 51,
          uniques: 36,
        },
      ],
      paths: [
        {
          path: '/proompteng/bilig',
          title: 'bilig',
          count: 199,
          uniques: 145,
        },
      ],
    })
  })

  it('collects GitHub discussion and traffic metrics when a token is provided', async () => {
    const fetchImpl = async (input: string | URL | Request, init?: RequestInit): Promise<Response> => {
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
          topics: ['spreadsheet', 'formula-engine'],
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

      if (url === 'https://api.github.com/graphql') {
        const query = requestBodyQuery(init)

        if (query.includes('CommunityGrowthContributorFunnel')) {
          return responseJson({
            data: {
              goodFirst: {
                issueCount: 8,
              },
              firstTimers: {
                issueCount: 5,
              },
              documentationStarters: {
                issueCount: 4,
              },
              nonDocumentationStarters: {
                issueCount: 1,
              },
              helpWanted: {
                issueCount: 8,
              },
              openPullRequests: {
                issueCount: 2,
              },
              externalOpenIssues: {
                issueCount: 3,
              },
              externalOpenPullRequests: {
                issueCount: 1,
              },
              externalRecentIssues: {
                issueCount: 2,
              },
              externalRecentPullRequests: {
                issueCount: 1,
              },
            },
          })
        }

        if (!query.includes('CommunityGrowthDiscussions')) {
          throw new Error(`unexpected GraphQL query ${query}`)
        }

        return responseJson({
          data: {
            repository: {
              discussions: {
                totalCount: 2,
                nodes: [
                  {
                    number: 157,
                    title: 'Which Node workbook automation workflow should @bilig/headless prove next?',
                    url: 'https://github.com/proompteng/bilig/discussions/157',
                    category: {
                      name: 'Ideas',
                    },
                    createdAt: '2026-05-08T21:46:51Z',
                    updatedAt: '2026-05-12T21:10:00Z',
                    comments: {
                      totalCount: 2,
                    },
                  },
                ],
              },
            },
          },
        })
      }

      throw new Error(`unexpected fetch ${url}`)
    }

    const snapshot = await collectCommunityGrowthSnapshot({
      fetchImpl: fetchImpl as typeof fetch,
      githubToken: 'test-token',
      now: new Date('2026-05-08T06:00:00.000Z'),
    })

    expect(snapshot.contributorFunnel).toEqual({
      openGoodFirstIssueCount: 8,
      openFirstTimersOnlyIssueCount: 5,
      openDocumentationStarterIssueCount: 4,
      openNonDocumentationStarterIssueCount: 1,
      openHelpWantedIssueCount: 8,
      openPullRequestCount: 2,
      externalOpenIssueCount: 3,
      externalOpenPullRequestCount: 1,
      externalIssuesOpenedLastSevenDays: 2,
      externalPullRequestsOpenedLastSevenDays: 1,
    })
    expect(snapshot.discussionActivity).toEqual({
      available: true,
      totalCount: 2,
      recent: [
        {
          number: 157,
          title: 'Which Node workbook automation workflow should @bilig/headless prove next?',
          url: 'https://github.com/proompteng/bilig/discussions/157',
          category: 'Ideas',
          createdAt: '2026-05-08T21:46:51Z',
          updatedAt: '2026-05-12T21:10:00Z',
          commentCount: 2,
        },
      ],
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

  it('renders a markdown growth snapshot for weekly tracking', () => {
    const markdown = renderCommunityGrowthSnapshotMarkdown({
      capturedAt: '2026-05-12T22:14:21.495Z',
      github: {
        fullName: 'proompteng/bilig',
        htmlUrl: 'https://github.com/proompteng/bilig',
        description: 'Headless spreadsheet engine',
        stargazerCount: 24,
        forkCount: 4,
        watcherCount: 1,
        openIssueCount: 29,
        defaultBranch: 'main',
        topics: ['headless-spreadsheet', 'typescript'],
      },
      npm: {
        name: '@bilig/headless',
        version: '0.11.24',
        description: 'Headless spreadsheet engine',
        license: 'MIT',
        modifiedAt: '2026-05-12T05:53:15.971Z',
        downloads: {
          lastWeek: {
            downloads: 13427,
            start: '2026-05-05',
            end: '2026-05-11',
          },
          lastMonth: {
            downloads: 24931,
            start: '2026-04-12',
            end: '2026-05-11',
          },
        },
      },
      contributorFunnel: {
        openGoodFirstIssueCount: 22,
        openFirstTimersOnlyIssueCount: 22,
        openDocumentationStarterIssueCount: 20,
        openNonDocumentationStarterIssueCount: 2,
        openHelpWantedIssueCount: 22,
        openPullRequestCount: 0,
        externalOpenIssueCount: 1,
        externalOpenPullRequestCount: 0,
        externalIssuesOpenedLastSevenDays: 22,
        externalPullRequestsOpenedLastSevenDays: 5,
      },
      discussionActivity: {
        available: true,
        totalCount: 4,
        recent: [
          {
            number: 213,
            title: 'Five Node workbook automation examples',
            url: 'https://github.com/proompteng/bilig/discussions/213',
            category: 'Show and tell',
            createdAt: '2026-05-12T21:46:51Z',
            updatedAt: '2026-05-12T22:10:00Z',
            commentCount: 1,
          },
        ],
      },
      traffic: {
        available: true,
        views: {
          count: 393,
          uniques: 159,
        },
        clones: {
          count: 18287,
          uniques: 1907,
        },
        referrers: [
          {
            referrer: 'news.ycombinator.com',
            count: 51,
            uniques: 36,
          },
        ],
        paths: [
          {
            path: '/proompteng/bilig',
            title: 'bilig',
            count: 199,
            uniques: 145,
          },
        ],
      },
    })

    expect(markdown).toContain('# Community Growth Snapshot')
    expect(markdown).toContain('- Stars: 24 / 1,000 (976 remaining)')
    expect(markdown).toContain('- Topics: `headless-spreadsheet`, `typescript`')
    expect(markdown).toContain('- Downloads last week: 13,427 (2026-05-05 to 2026-05-11)')
    expect(markdown).toContain('- Open good first issues: 22')
    expect(markdown).toContain('- Non-documentation starter issues: 2')
    expect(markdown).toContain(
      '- #213 [Five Node workbook automation examples](https://github.com/proompteng/bilig/discussions/213) (Show and tell, 1 comment)',
    )
    expect(markdown).toContain('- Clones: 18,287 from 1,907 unique cloners')
    expect(markdown).toContain('news.ycombinator.com (51/36)')
    expect(markdown).toContain('## External Discovery')
    expect(markdown).toContain('https://www.libhunt.com/topic/headless-spreadsheet')
    expect(markdown).toContain('May 7 Show HN discovery path')
    expect(markdown).toContain('## Conversion Pressure')
    expect(markdown).toContain('- Last-week npm downloads per current star: 559')
    expect(markdown).toContain('- Last-month npm downloads per current star: 1,039')
    expect(markdown).toContain('- Fourteen-day unique GitHub visitors per current star: 7')
    expect(markdown).toContain('- Fourteen-day unique cloners per current star: 79')
    expect(markdown).toContain('these are pressure ratios, not attribution')
    expect(markdown).toContain('## Spike Read')
    expect(markdown).toContain('The strongest current external referrer is news.ycombinator.com with 51 views from 36 unique visitors.')
    expect(markdown).toContain('do not repost the same launch')
  })
})
