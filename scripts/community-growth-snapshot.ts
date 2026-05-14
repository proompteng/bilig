import { execFile } from 'node:child_process'
import { writeFile } from 'node:fs/promises'
import { pathToFileURL } from 'node:url'
import { promisify } from 'node:util'

export interface CommunityGrowthSnapshot {
  readonly capturedAt: string
  readonly github: GitHubRepoGrowthMetrics
  readonly npm: NpmPackageGrowthMetrics
  readonly contributorFunnel: ContributorFunnelMetrics
  readonly discussionActivity: GitHubDiscussionActivitySnapshot
  readonly traffic: GitHubTrafficSnapshot
}

export interface GitHubRepoGrowthMetrics {
  readonly fullName: string
  readonly htmlUrl: string
  readonly description: string
  readonly stargazerCount: number
  readonly forkCount: number
  readonly watcherCount: number
  readonly openIssueCount: number
  readonly defaultBranch: string
}

export interface NpmPackageGrowthMetrics {
  readonly name: string
  readonly version: string
  readonly description: string
  readonly license: string
  readonly modifiedAt: string
  readonly downloads: {
    readonly lastWeek: NpmDownloadWindow
    readonly lastMonth: NpmDownloadWindow
  }
}

export interface NpmDownloadWindow {
  readonly downloads: number
  readonly start: string
  readonly end: string
}

export interface ContributorFunnelMetrics {
  readonly openGoodFirstIssueCount: number
  readonly openFirstTimersOnlyIssueCount: number
  readonly openDocumentationStarterIssueCount: number
  readonly openNonDocumentationStarterIssueCount: number
  readonly openHelpWantedIssueCount: number
  readonly openPullRequestCount: number
  readonly externalOpenIssueCount: number
  readonly externalOpenPullRequestCount: number
  readonly externalIssuesOpenedLastSevenDays: number
  readonly externalPullRequestsOpenedLastSevenDays: number
}

export type GitHubDiscussionActivitySnapshot =
  | {
      readonly available: false
      readonly reason: string
    }
  | {
      readonly available: true
      readonly totalCount: number
      readonly recent: readonly GitHubDiscussionSummary[]
    }

export interface GitHubDiscussionSummary {
  readonly number: number
  readonly title: string
  readonly url: string
  readonly category: string
  readonly createdAt: string
  readonly updatedAt: string
  readonly commentCount: number
}

export type GitHubTrafficSnapshot =
  | {
      readonly available: false
      readonly reason: string
    }
  | {
      readonly available: true
      readonly views: GitHubTrafficCount
      readonly clones: GitHubTrafficCount
      readonly referrers: readonly GitHubTrafficReferrer[]
      readonly paths: readonly GitHubTrafficPath[]
    }

export interface GitHubTrafficCount {
  readonly count: number
  readonly uniques: number
}

export interface GitHubTrafficReferrer {
  readonly referrer: string
  readonly count: number
  readonly uniques: number
}

export interface GitHubTrafficPath {
  readonly path: string
  readonly title: string
  readonly count: number
  readonly uniques: number
}

export interface CommunityGrowthSnapshotOptions {
  readonly owner?: string
  readonly repo?: string
  readonly packageName?: string
  readonly maintainerLogin?: string
  readonly githubToken?: string
  readonly githubCliApiJson?: GitHubCliApiJson | false
  readonly now?: Date
  readonly fetchImpl?: typeof fetch
}

export interface GitHubCliApiOptions {
  readonly fields?: readonly GitHubCliApiField[]
}

export interface GitHubCliApiField {
  readonly name: string
  readonly value: string
}

export type GitHubCliApiJson = (endpoint: string, options?: GitHubCliApiOptions) => Promise<unknown>

interface CliOptions {
  readonly owner: string
  readonly repo: string
  readonly packageName: string
  readonly maintainerLogin: string
  readonly outputPath: string | undefined
  readonly format: 'json' | 'markdown'
}

const defaultOwner = 'proompteng'
const defaultRepo = 'bilig'
const defaultPackageName = '@bilig/headless'
const defaultMaintainerLogin = 'gregkonush'
const starGoal = 1000
const execFileAsync = promisify(execFile)
const githubCliMaxBufferBytes = 4 * 1024 * 1024

async function runGitHubCliApiJson(endpoint: string, options: GitHubCliApiOptions = {}): Promise<unknown> {
  const args = ['api', endpoint]

  for (const field of options.fields ?? []) {
    args.push('-f', `${field.name}=${field.value}`)
  }

  let stdout = ''
  try {
    const result = await execFileAsync('gh', args, {
      encoding: 'utf8',
      maxBuffer: githubCliMaxBufferBytes,
    })
    stdout = String(result.stdout)
  } catch {
    return undefined
  }

  const trimmed = stdout.trim()
  if (trimmed === '') {
    return undefined
  }

  try {
    return JSON.parse(trimmed) as unknown
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    throw new Error(`gh api ${endpoint} returned invalid JSON: ${message}`, { cause: error })
  }
}

function asRecord(value: unknown, context: string): Record<string, unknown> {
  if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
    const record: Record<string, unknown> = {}
    for (const key of Object.keys(value)) {
      record[key] = Reflect.get(value, key)
    }
    return record
  }
  throw new Error(`${context} response was not an object`)
}

function asRecordArray(value: unknown, context: string): readonly Record<string, unknown>[] {
  if (Array.isArray(value)) {
    return value.map((item, index) => asRecord(item, `${context}[${String(index)}]`))
  }
  throw new Error(`${context} response was not an object array`)
}

function stringField(record: Record<string, unknown>, key: string, context: string): string {
  const value = record[key]
  if (typeof value === 'string') {
    return value
  }
  throw new Error(`${context}.${key} was not a string`)
}

function optionalStringField(record: Record<string, unknown>, key: string): string {
  const value = record[key]
  return typeof value === 'string' ? value : ''
}

function numberField(record: Record<string, unknown>, key: string, context: string): number {
  const value = record[key]
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }
  throw new Error(`${context}.${key} was not a finite number`)
}

function nestedRecord(record: Record<string, unknown>, key: string, context: string): Record<string, unknown> {
  return asRecord(record[key], `${context}.${key}`)
}

function githubHeaders(githubToken: string | undefined): Record<string, string> {
  const headers: Record<string, string> = {
    accept: 'application/vnd.github+json',
    'x-github-api-version': '2022-11-28',
  }

  if (githubToken !== undefined && githubToken.trim() !== '') {
    headers.authorization = `Bearer ${githubToken}`
  }

  return headers
}

async function fetchJson(fetchImpl: typeof fetch, url: string, init?: RequestInit): Promise<unknown> {
  const response = await fetchImpl(url, init)
  if (!response.ok) {
    throw new Error(`GET ${url} failed with HTTP ${String(response.status)}`)
  }
  return await response.json()
}

async function fetchGitHubGraphqlJson(
  fetchImpl: typeof fetch,
  githubToken: string,
  query: string,
  variables: Record<string, unknown>,
): Promise<unknown> {
  const response = await fetchImpl('https://api.github.com/graphql', {
    method: 'POST',
    headers: {
      ...githubHeaders(githubToken),
      'content-type': 'application/json',
    },
    body: JSON.stringify({
      query,
      variables,
    }),
  })

  if (response.status === 401 || response.status === 403) {
    return undefined
  }

  if (!response.ok) {
    throw new Error(`POST https://api.github.com/graphql failed with HTTP ${String(response.status)}`)
  }

  return await response.json()
}

async function fetchTrafficJson(fetchImpl: typeof fetch, url: string, githubToken: string): Promise<unknown> {
  const response = await fetchImpl(url, {
    headers: githubHeaders(githubToken),
  })

  if (response.status === 401 || response.status === 403 || response.status === 404) {
    return undefined
  }

  if (!response.ok) {
    throw new Error(`GET ${url} failed with HTTP ${String(response.status)}`)
  }

  return await response.json()
}

function githubGraphqlCliFields(query: string, variables: Record<string, unknown>): readonly GitHubCliApiField[] {
  const fields: GitHubCliApiField[] = [
    {
      name: 'query',
      value: query,
    },
  ]

  for (const [name, value] of Object.entries(variables)) {
    if (typeof value !== 'string') {
      throw new Error(`GitHub GraphQL CLI variable ${name} was not a string`)
    }
    fields.push({
      name,
      value,
    })
  }

  return fields
}

async function collectGitHubGraphqlJson(
  fetchImpl: typeof fetch,
  githubToken: string | undefined,
  githubCliApiJson: GitHubCliApiJson | undefined,
  query: string,
  variables: Record<string, unknown>,
): Promise<unknown> {
  if (githubToken !== undefined && githubToken.trim() !== '') {
    const result = await fetchGitHubGraphqlJson(fetchImpl, githubToken, query, variables)
    if (result !== undefined) {
      return result
    }
  }

  if (githubCliApiJson === undefined) {
    return undefined
  }

  return await githubCliApiJson('graphql', {
    fields: githubGraphqlCliFields(query, variables),
  })
}

function parseGitHubRepoMetrics(value: unknown): GitHubRepoGrowthMetrics {
  const repo = asRecord(value, 'GitHub repository')

  return {
    fullName: stringField(repo, 'full_name', 'GitHub repository'),
    htmlUrl: stringField(repo, 'html_url', 'GitHub repository'),
    description: optionalStringField(repo, 'description'),
    stargazerCount: numberField(repo, 'stargazers_count', 'GitHub repository'),
    forkCount: numberField(repo, 'forks_count', 'GitHub repository'),
    watcherCount: numberField(repo, 'subscribers_count', 'GitHub repository'),
    openIssueCount: numberField(repo, 'open_issues_count', 'GitHub repository'),
    defaultBranch: stringField(repo, 'default_branch', 'GitHub repository'),
  }
}

function parseNpmPackageMetrics(value: unknown, downloads: NpmPackageGrowthMetrics['downloads']): NpmPackageGrowthMetrics {
  const metadata = asRecord(value, 'npm metadata')
  const distTags = nestedRecord(metadata, 'dist-tags', 'npm metadata')
  const time = nestedRecord(metadata, 'time', 'npm metadata')

  return {
    name: stringField(metadata, 'name', 'npm metadata'),
    version: stringField(distTags, 'latest', 'npm metadata.dist-tags'),
    description: optionalStringField(metadata, 'description'),
    license: optionalStringField(metadata, 'license'),
    modifiedAt: stringField(time, 'modified', 'npm metadata.time'),
    downloads,
  }
}

function parseDownloadWindow(value: unknown, context: string): NpmDownloadWindow {
  const downloads = asRecord(value, context)

  return {
    downloads: numberField(downloads, 'downloads', context),
    start: stringField(downloads, 'start', context),
    end: stringField(downloads, 'end', context),
  }
}

function parseTrafficCount(value: unknown, context: string): GitHubTrafficCount {
  const traffic = asRecord(value, context)

  return {
    count: numberField(traffic, 'count', context),
    uniques: numberField(traffic, 'uniques', context),
  }
}

function parseTrafficReferrers(value: unknown): readonly GitHubTrafficReferrer[] {
  return asRecordArray(value, 'GitHub traffic referrers').map((referrer) => ({
    referrer: stringField(referrer, 'referrer', 'GitHub traffic referrer'),
    count: numberField(referrer, 'count', 'GitHub traffic referrer'),
    uniques: numberField(referrer, 'uniques', 'GitHub traffic referrer'),
  }))
}

function parseTrafficPaths(value: unknown): readonly GitHubTrafficPath[] {
  return asRecordArray(value, 'GitHub traffic paths').map((path) => ({
    path: stringField(path, 'path', 'GitHub traffic path'),
    title: stringField(path, 'title', 'GitHub traffic path'),
    count: numberField(path, 'count', 'GitHub traffic path'),
    uniques: numberField(path, 'uniques', 'GitHub traffic path'),
  }))
}

function parseSearchCount(value: unknown, context: string): number {
  return numberField(asRecord(value, context), 'total_count', context)
}

function parseDiscussionSummary(value: Record<string, unknown>): GitHubDiscussionSummary {
  const category = nestedRecord(value, 'category', 'GitHub discussion')
  const comments = nestedRecord(value, 'comments', 'GitHub discussion')

  return {
    number: numberField(value, 'number', 'GitHub discussion'),
    title: stringField(value, 'title', 'GitHub discussion'),
    url: stringField(value, 'url', 'GitHub discussion'),
    category: stringField(category, 'name', 'GitHub discussion category'),
    createdAt: stringField(value, 'createdAt', 'GitHub discussion'),
    updatedAt: stringField(value, 'updatedAt', 'GitHub discussion'),
    commentCount: numberField(comments, 'totalCount', 'GitHub discussion comments'),
  }
}

function parseDiscussionActivity(value: unknown): GitHubDiscussionActivitySnapshot {
  const payload = asRecord(value, 'GitHub GraphQL response')
  const data = nestedRecord(payload, 'data', 'GitHub GraphQL response')
  const repository = nestedRecord(data, 'repository', 'GitHub GraphQL data')
  const discussions = nestedRecord(repository, 'discussions', 'GitHub repository')
  const nodes = asRecordArray(discussions.nodes, 'GitHub discussions')

  return {
    available: true,
    totalCount: numberField(discussions, 'totalCount', 'GitHub discussions'),
    recent: nodes.map(parseDiscussionSummary),
  }
}

function graphqlSearchCount(data: Record<string, unknown>, key: string): number {
  const result = nestedRecord(data, key, `GitHub GraphQL search ${key}`)
  return numberField(result, 'issueCount', `GitHub GraphQL search ${key}`)
}

function parseContributorFunnelGraphql(value: unknown): ContributorFunnelMetrics {
  const payload = asRecord(value, 'GitHub GraphQL response')
  const data = nestedRecord(payload, 'data', 'GitHub GraphQL response')

  return {
    openGoodFirstIssueCount: graphqlSearchCount(data, 'goodFirst'),
    openFirstTimersOnlyIssueCount: graphqlSearchCount(data, 'firstTimers'),
    openDocumentationStarterIssueCount: graphqlSearchCount(data, 'documentationStarters'),
    openNonDocumentationStarterIssueCount: graphqlSearchCount(data, 'nonDocumentationStarters'),
    openHelpWantedIssueCount: graphqlSearchCount(data, 'helpWanted'),
    openPullRequestCount: graphqlSearchCount(data, 'openPullRequests'),
    externalOpenIssueCount: graphqlSearchCount(data, 'externalOpenIssues'),
    externalOpenPullRequestCount: graphqlSearchCount(data, 'externalOpenPullRequests'),
    externalIssuesOpenedLastSevenDays: graphqlSearchCount(data, 'externalRecentIssues'),
    externalPullRequestsOpenedLastSevenDays: graphqlSearchCount(data, 'externalRecentPullRequests'),
  }
}

function dateOnly(date: Date): string {
  return date.toISOString().slice(0, 10)
}

function daysBefore(date: Date, days: number): Date {
  return new Date(date.getTime() - days * 24 * 60 * 60 * 1000)
}

async function fetchIssueSearchCount(fetchImpl: typeof fetch, githubToken: string | undefined, query: string): Promise<number> {
  const searchUrl = new URL('https://api.github.com/search/issues')
  searchUrl.searchParams.set('q', query)
  searchUrl.searchParams.set('per_page', '1')

  const result = await fetchJson(fetchImpl, searchUrl.href, {
    headers: githubHeaders(githubToken),
  })

  return parseSearchCount(result, `GitHub issue search ${query}`)
}

async function collectContributorFunnel(
  fetchImpl: typeof fetch,
  owner: string,
  repo: string,
  maintainerLogin: string,
  githubToken: string | undefined,
  githubCliApiJson: GitHubCliApiJson | undefined,
  now: Date,
): Promise<ContributorFunnelMetrics> {
  const repoQualifier = `repo:${owner}/${repo}`
  const externalQualifier = `-author:${maintainerLogin}`
  const since = dateOnly(daysBefore(now, 7))

  if ((githubToken !== undefined && githubToken.trim() !== '') || githubCliApiJson !== undefined) {
    const query = `
      query CommunityGrowthContributorFunnel(
        $goodFirst: String!
        $firstTimers: String!
        $documentationStarters: String!
        $nonDocumentationStarters: String!
        $helpWanted: String!
        $openPullRequests: String!
        $externalOpenIssues: String!
        $externalOpenPullRequests: String!
        $externalRecentIssues: String!
        $externalRecentPullRequests: String!
      ) {
        goodFirst: search(type: ISSUE, query: $goodFirst, first: 0) {
          issueCount
        }
        firstTimers: search(type: ISSUE, query: $firstTimers, first: 0) {
          issueCount
        }
        documentationStarters: search(type: ISSUE, query: $documentationStarters, first: 0) {
          issueCount
        }
        nonDocumentationStarters: search(type: ISSUE, query: $nonDocumentationStarters, first: 0) {
          issueCount
        }
        helpWanted: search(type: ISSUE, query: $helpWanted, first: 0) {
          issueCount
        }
        openPullRequests: search(type: ISSUE, query: $openPullRequests, first: 0) {
          issueCount
        }
        externalOpenIssues: search(type: ISSUE, query: $externalOpenIssues, first: 0) {
          issueCount
        }
        externalOpenPullRequests: search(type: ISSUE, query: $externalOpenPullRequests, first: 0) {
          issueCount
        }
        externalRecentIssues: search(type: ISSUE, query: $externalRecentIssues, first: 0) {
          issueCount
        }
        externalRecentPullRequests: search(type: ISSUE, query: $externalRecentPullRequests, first: 0) {
          issueCount
        }
      }
    `
    const result = await collectGitHubGraphqlJson(fetchImpl, githubToken, githubCliApiJson, query, {
      goodFirst: `${repoQualifier} is:issue is:open label:"good first issue"`,
      firstTimers: `${repoQualifier} is:issue is:open label:first-timers-only`,
      documentationStarters: `${repoQualifier} is:issue is:open label:first-timers-only label:documentation`,
      nonDocumentationStarters: `${repoQualifier} is:issue is:open label:first-timers-only -label:documentation`,
      helpWanted: `${repoQualifier} is:issue is:open label:"help wanted"`,
      openPullRequests: `${repoQualifier} is:pr is:open`,
      externalOpenIssues: `${repoQualifier} is:issue is:open ${externalQualifier}`,
      externalOpenPullRequests: `${repoQualifier} is:pr is:open ${externalQualifier}`,
      externalRecentIssues: `${repoQualifier} is:issue created:>=${since} ${externalQualifier}`,
      externalRecentPullRequests: `${repoQualifier} is:pr created:>=${since} ${externalQualifier}`,
    })

    if (result !== undefined) {
      return parseContributorFunnelGraphql(result)
    }
  }

  const [
    openGoodFirstIssueCount,
    openFirstTimersOnlyIssueCount,
    openDocumentationStarterIssueCount,
    openNonDocumentationStarterIssueCount,
    openHelpWantedIssueCount,
    openPullRequestCount,
    externalOpenIssueCount,
    externalOpenPullRequestCount,
    externalIssuesOpenedLastSevenDays,
    externalPullRequestsOpenedLastSevenDays,
  ] = await Promise.all([
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue is:open label:"good first issue"`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue is:open label:first-timers-only`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue is:open label:first-timers-only label:documentation`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue is:open label:first-timers-only -label:documentation`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue is:open label:"help wanted"`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:pr is:open`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue is:open ${externalQualifier}`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:pr is:open ${externalQualifier}`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue created:>=${since} ${externalQualifier}`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:pr created:>=${since} ${externalQualifier}`),
  ])

  return {
    openGoodFirstIssueCount,
    openFirstTimersOnlyIssueCount,
    openDocumentationStarterIssueCount,
    openNonDocumentationStarterIssueCount,
    openHelpWantedIssueCount,
    openPullRequestCount,
    externalOpenIssueCount,
    externalOpenPullRequestCount,
    externalIssuesOpenedLastSevenDays,
    externalPullRequestsOpenedLastSevenDays,
  }
}

async function collectGitHubTraffic(
  fetchImpl: typeof fetch,
  owner: string,
  repo: string,
  githubToken: string | undefined,
  githubCliApiJson: GitHubCliApiJson | undefined,
): Promise<GitHubTrafficSnapshot> {
  const baseUrl = `https://api.github.com/repos/${owner}/${repo}/traffic`
  const baseEndpoint = `repos/${owner}/${repo}/traffic`
  const collectTrafficJson = async (path: string): Promise<unknown> => {
    if (githubToken !== undefined && githubToken.trim() !== '') {
      const result = await fetchTrafficJson(fetchImpl, `${baseUrl}/${path}`, githubToken)
      if (result !== undefined) {
        return result
      }
    }

    if (githubCliApiJson === undefined) {
      return undefined
    }

    return await githubCliApiJson(`${baseEndpoint}/${path}`)
  }

  const [views, clones, referrers, paths] = await Promise.all([
    collectTrafficJson('views'),
    collectTrafficJson('clones'),
    collectTrafficJson('popular/referrers'),
    collectTrafficJson('popular/paths'),
  ])

  if (views === undefined || clones === undefined || referrers === undefined || paths === undefined) {
    return {
      available: false,
      reason:
        'Set GITHUB_TOKEN or GH_TOKEN with repository traffic access, or authenticate gh CLI as a repository collaborator, to collect views, clones, referrers, and paths.',
    }
  }

  return {
    available: true,
    views: parseTrafficCount(views, 'GitHub traffic views'),
    clones: parseTrafficCount(clones, 'GitHub traffic clones'),
    referrers: parseTrafficReferrers(referrers),
    paths: parseTrafficPaths(paths),
  }
}

async function collectGitHubDiscussionActivity(
  fetchImpl: typeof fetch,
  owner: string,
  repo: string,
  githubToken: string | undefined,
  githubCliApiJson: GitHubCliApiJson | undefined,
): Promise<GitHubDiscussionActivitySnapshot> {
  const query = `
    query CommunityGrowthDiscussions($owner: String!, $repo: String!) {
      repository(owner: $owner, name: $repo) {
        discussions(first: 20, orderBy: { field: UPDATED_AT, direction: DESC }) {
          totalCount
          nodes {
            number
            title
            url
            createdAt
            updatedAt
            category {
              name
            }
            comments {
              totalCount
            }
          }
        }
      }
    }
  `
  const result = await collectGitHubGraphqlJson(fetchImpl, githubToken, githubCliApiJson, query, {
    owner,
    repo,
  })

  if (result === undefined) {
    return {
      available: false,
      reason: 'Set GITHUB_TOKEN or GH_TOKEN, or authenticate gh CLI, to collect recent GitHub discussion activity.',
    }
  }

  return parseDiscussionActivity(result)
}

export async function collectCommunityGrowthSnapshot(options: CommunityGrowthSnapshotOptions = {}): Promise<CommunityGrowthSnapshot> {
  const owner = options.owner ?? defaultOwner
  const repo = options.repo ?? defaultRepo
  const packageName = options.packageName ?? defaultPackageName
  const maintainerLogin = options.maintainerLogin ?? defaultMaintainerLogin
  const fetchImpl = options.fetchImpl ?? fetch
  const githubToken = options.githubToken ?? process.env.GITHUB_TOKEN ?? process.env.GH_TOKEN
  const githubCliApiJson = options.githubCliApiJson === false ? undefined : (options.githubCliApiJson ?? runGitHubCliApiJson)
  const now = options.now ?? new Date()
  const encodedPackageName = encodeURIComponent(packageName)

  const [github, npmMetadata, lastWeekDownloads, lastMonthDownloads, contributorFunnel, discussionActivity, traffic] = await Promise.all([
    fetchJson(fetchImpl, `https://api.github.com/repos/${owner}/${repo}`, {
      headers: githubHeaders(githubToken),
    }).then(parseGitHubRepoMetrics),
    fetchJson(fetchImpl, `https://registry.npmjs.org/${encodedPackageName}`),
    fetchJson(fetchImpl, `https://api.npmjs.org/downloads/point/last-week/${encodedPackageName}`).then((value) =>
      parseDownloadWindow(value, 'npm last-week downloads'),
    ),
    fetchJson(fetchImpl, `https://api.npmjs.org/downloads/point/last-month/${encodedPackageName}`).then((value) =>
      parseDownloadWindow(value, 'npm last-month downloads'),
    ),
    collectContributorFunnel(fetchImpl, owner, repo, maintainerLogin, githubToken, githubCliApiJson, now),
    collectGitHubDiscussionActivity(fetchImpl, owner, repo, githubToken, githubCliApiJson),
    collectGitHubTraffic(fetchImpl, owner, repo, githubToken, githubCliApiJson),
  ])
  const issueOnlyGitHubMetrics = {
    ...github,
    // GitHub REST's open_issues_count includes pull requests.
    openIssueCount: Math.max(0, github.openIssueCount - contributorFunnel.openPullRequestCount),
  }

  return {
    capturedAt: now.toISOString(),
    github: issueOnlyGitHubMetrics,
    npm: parseNpmPackageMetrics(npmMetadata, {
      lastWeek: lastWeekDownloads,
      lastMonth: lastMonthDownloads,
    }),
    contributorFunnel,
    discussionActivity,
    traffic,
  }
}

function formatCount(value: number): string {
  return String(value).replace(/\B(?=(\d{3})+(?!\d))/g, ',')
}

function markdownLink(label: string, url: string): string {
  return `[${label}](${url})`
}

function formatCommentCount(value: number): string {
  return `${formatCount(value)} ${value === 1 ? 'comment' : 'comments'}`
}

function renderDiscussionActivityMarkdown(discussionActivity: GitHubDiscussionActivitySnapshot): readonly string[] {
  if (!discussionActivity.available) {
    return [`- Discussion activity: unavailable. ${discussionActivity.reason}`]
  }

  const lines = [`- Total discussions: ${formatCount(discussionActivity.totalCount)}`]
  for (const discussion of discussionActivity.recent.slice(0, 5)) {
    lines.push(
      `- #${String(discussion.number)} ${markdownLink(discussion.title, discussion.url)} (${discussion.category}, ${formatCommentCount(discussion.commentCount)})`,
    )
  }
  return lines
}

function renderTrafficMarkdown(traffic: GitHubTrafficSnapshot): readonly string[] {
  if (!traffic.available) {
    return [`- Traffic: unavailable. ${traffic.reason}`]
  }

  const lines = [
    `- Views: ${formatCount(traffic.views.count)} from ${formatCount(traffic.views.uniques)} unique visitors`,
    `- Clones: ${formatCount(traffic.clones.count)} from ${formatCount(traffic.clones.uniques)} unique cloners`,
  ]

  if (traffic.referrers.length > 0) {
    lines.push(
      `- Top referrers: ${traffic.referrers
        .slice(0, 5)
        .map((referrer) => `${referrer.referrer} (${formatCount(referrer.count)}/${formatCount(referrer.uniques)})`)
        .join(', ')}`,
    )
  }

  if (traffic.paths.length > 0) {
    lines.push(
      `- Top paths: ${traffic.paths
        .slice(0, 5)
        .map((path) => `${path.path} (${formatCount(path.count)}/${formatCount(path.uniques)})`)
        .join(', ')}`,
    )
  }

  return lines
}

export function renderCommunityGrowthSnapshotMarkdown(snapshot: CommunityGrowthSnapshot): string {
  const starRemaining = Math.max(0, starGoal - snapshot.github.stargazerCount)
  const lines = [
    '# Community Growth Snapshot',
    '',
    `Captured at: \`${snapshot.capturedAt}\``,
    '',
    'This snapshot tracks the public signals for the `@bilig/headless` growth loop: GitHub conversion, npm demand, contributor on-ramp health, discussion activity, and traffic quality.',
    '',
    '## GitHub',
    '',
    `- Repository: ${markdownLink(snapshot.github.fullName, snapshot.github.htmlUrl)}`,
    `- Stars: ${formatCount(snapshot.github.stargazerCount)} / ${formatCount(starGoal)} (${formatCount(starRemaining)} remaining)`,
    `- Forks: ${formatCount(snapshot.github.forkCount)}`,
    `- Watchers: ${formatCount(snapshot.github.watcherCount)}`,
    `- Open issues: ${formatCount(snapshot.github.openIssueCount)}`,
    `- Default branch: \`${snapshot.github.defaultBranch}\``,
    '',
    '## npm',
    '',
    `- Package: \`${snapshot.npm.name}@${snapshot.npm.version}\``,
    `- License: \`${snapshot.npm.license || 'unknown'}\``,
    `- Modified: \`${snapshot.npm.modifiedAt}\``,
    `- Downloads last week: ${formatCount(snapshot.npm.downloads.lastWeek.downloads)} (${snapshot.npm.downloads.lastWeek.start} to ${snapshot.npm.downloads.lastWeek.end})`,
    `- Downloads last month: ${formatCount(snapshot.npm.downloads.lastMonth.downloads)} (${snapshot.npm.downloads.lastMonth.start} to ${snapshot.npm.downloads.lastMonth.end})`,
    '',
    '## Contributor Funnel',
    '',
    `- Open good first issues: ${formatCount(snapshot.contributorFunnel.openGoodFirstIssueCount)}`,
    `- Open first-timers-only issues: ${formatCount(snapshot.contributorFunnel.openFirstTimersOnlyIssueCount)}`,
    `- Documentation starter issues: ${formatCount(snapshot.contributorFunnel.openDocumentationStarterIssueCount)}`,
    `- Non-documentation starter issues: ${formatCount(snapshot.contributorFunnel.openNonDocumentationStarterIssueCount)}`,
    `- Open help wanted issues: ${formatCount(snapshot.contributorFunnel.openHelpWantedIssueCount)}`,
    `- Open pull requests: ${formatCount(snapshot.contributorFunnel.openPullRequestCount)}`,
    `- External open issues: ${formatCount(snapshot.contributorFunnel.externalOpenIssueCount)}`,
    `- External open pull requests: ${formatCount(snapshot.contributorFunnel.externalOpenPullRequestCount)}`,
    `- External issues opened in the last 7 days: ${formatCount(snapshot.contributorFunnel.externalIssuesOpenedLastSevenDays)}`,
    `- External pull requests opened in the last 7 days: ${formatCount(snapshot.contributorFunnel.externalPullRequestsOpenedLastSevenDays)}`,
    '',
    '## Discussions',
    '',
    ...renderDiscussionActivityMarkdown(snapshot.discussionActivity),
    '',
    '## Traffic',
    '',
    ...renderTrafficMarkdown(snapshot.traffic),
    '',
    '## Read This Snapshot',
    '',
    '- Stars are the primary lagging goal; npm downloads, clone traffic, and external issues are leading signals.',
    '- If downloads or clones rise without stars, improve README and npm star/bookmark conversion after proof blocks.',
    '- If traffic comes from a channel but discussions stay quiet, switch from launch copy to a specific workflow-feedback ask.',
    '- If the starter queue drops below three current issues, open scoped example tasks before the next distribution push.',
  ]

  return `${lines.join('\n')}\n`
}

function parseCliOptions(argv: readonly string[]): CliOptions {
  let owner = defaultOwner
  let repo = defaultRepo
  let packageName = defaultPackageName
  let maintainerLogin = defaultMaintainerLogin
  let outputPath: string | undefined
  let format: 'json' | 'markdown' = 'json'

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    const next = argv[index + 1]

    if (arg === '--owner' && next !== undefined) {
      owner = next
      index += 1
    } else if (arg === '--repo' && next !== undefined) {
      repo = next
      index += 1
    } else if (arg === '--package' && next !== undefined) {
      packageName = next
      index += 1
    } else if (arg === '--maintainer' && next !== undefined) {
      maintainerLogin = next
      index += 1
    } else if (arg === '--output' && next !== undefined) {
      outputPath = next
      index += 1
    } else if (arg === '--format' && next !== undefined) {
      if (next !== 'json' && next !== 'markdown') {
        throw new Error(`Unsupported snapshot format: ${next}`)
      }
      format = next
      index += 1
    } else {
      throw new Error(`Unknown or incomplete argument: ${arg ?? ''}`)
    }
  }

  return {
    owner,
    repo,
    packageName,
    maintainerLogin,
    outputPath,
    format,
  }
}

async function runCli(): Promise<void> {
  const options = parseCliOptions(process.argv.slice(2))
  const snapshot = await collectCommunityGrowthSnapshot(options)
  const output = options.format === 'markdown' ? renderCommunityGrowthSnapshotMarkdown(snapshot) : `${JSON.stringify(snapshot, null, 2)}\n`

  if (options.outputPath !== undefined) {
    await writeFile(options.outputPath, output)
    console.log(`wrote ${options.outputPath}`)
    return
  }

  process.stdout.write(output)
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  await runCli()
}
