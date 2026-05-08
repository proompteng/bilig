import { writeFile } from 'node:fs/promises'
import { pathToFileURL } from 'node:url'

export interface CommunityGrowthSnapshot {
  readonly capturedAt: string
  readonly github: GitHubRepoGrowthMetrics
  readonly npm: NpmPackageGrowthMetrics
  readonly contributorFunnel: ContributorFunnelMetrics
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
  readonly openHelpWantedIssueCount: number
  readonly openPullRequestCount: number
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
  readonly githubToken?: string
  readonly now?: Date
  readonly fetchImpl?: typeof fetch
}

interface CliOptions {
  readonly owner: string
  readonly repo: string
  readonly packageName: string
  readonly outputPath: string | undefined
}

const defaultOwner = 'proompteng'
const defaultRepo = 'bilig'
const defaultPackageName = '@bilig/headless'

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

function githubHeaders(githubToken: string | undefined): HeadersInit {
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
  githubToken: string | undefined,
): Promise<ContributorFunnelMetrics> {
  const repoQualifier = `repo:${owner}/${repo}`

  const [openGoodFirstIssueCount, openFirstTimersOnlyIssueCount, openHelpWantedIssueCount, openPullRequestCount] = await Promise.all([
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue is:open label:"good first issue"`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue is:open label:first-timers-only`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:issue is:open label:"help wanted"`),
    fetchIssueSearchCount(fetchImpl, githubToken, `${repoQualifier} is:pr is:open`),
  ])

  return {
    openGoodFirstIssueCount,
    openFirstTimersOnlyIssueCount,
    openHelpWantedIssueCount,
    openPullRequestCount,
  }
}

async function collectGitHubTraffic(
  fetchImpl: typeof fetch,
  owner: string,
  repo: string,
  githubToken: string | undefined,
): Promise<GitHubTrafficSnapshot> {
  if (githubToken === undefined || githubToken.trim() === '') {
    return {
      available: false,
      reason: 'Set GITHUB_TOKEN or GH_TOKEN with repository traffic access to collect views, clones, referrers, and paths.',
    }
  }

  const baseUrl = `https://api.github.com/repos/${owner}/${repo}/traffic`
  const [views, clones, referrers, paths] = await Promise.all([
    fetchTrafficJson(fetchImpl, `${baseUrl}/views`, githubToken),
    fetchTrafficJson(fetchImpl, `${baseUrl}/clones`, githubToken),
    fetchTrafficJson(fetchImpl, `${baseUrl}/popular/referrers`, githubToken),
    fetchTrafficJson(fetchImpl, `${baseUrl}/popular/paths`, githubToken),
  ])

  if (views === undefined || clones === undefined || referrers === undefined || paths === undefined) {
    return {
      available: false,
      reason: 'GitHub traffic API was unavailable for this token or repository.',
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

export async function collectCommunityGrowthSnapshot(options: CommunityGrowthSnapshotOptions = {}): Promise<CommunityGrowthSnapshot> {
  const owner = options.owner ?? defaultOwner
  const repo = options.repo ?? defaultRepo
  const packageName = options.packageName ?? defaultPackageName
  const fetchImpl = options.fetchImpl ?? fetch
  const githubToken = options.githubToken ?? process.env.GITHUB_TOKEN ?? process.env.GH_TOKEN
  const encodedPackageName = encodeURIComponent(packageName)

  const [github, npmMetadata, lastWeekDownloads, lastMonthDownloads, contributorFunnel, traffic] = await Promise.all([
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
    collectContributorFunnel(fetchImpl, owner, repo, githubToken),
    collectGitHubTraffic(fetchImpl, owner, repo, githubToken),
  ])

  return {
    capturedAt: (options.now ?? new Date()).toISOString(),
    github,
    npm: parseNpmPackageMetrics(npmMetadata, {
      lastWeek: lastWeekDownloads,
      lastMonth: lastMonthDownloads,
    }),
    contributorFunnel,
    traffic,
  }
}

function parseCliOptions(argv: readonly string[]): CliOptions {
  let owner = defaultOwner
  let repo = defaultRepo
  let packageName = defaultPackageName
  let outputPath: string | undefined

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
    } else if (arg === '--output' && next !== undefined) {
      outputPath = next
      index += 1
    } else {
      throw new Error(`Unknown or incomplete argument: ${arg ?? ''}`)
    }
  }

  return {
    owner,
    repo,
    packageName,
    outputPath,
  }
}

async function runCli(): Promise<void> {
  const options = parseCliOptions(process.argv.slice(2))
  const snapshot = await collectCommunityGrowthSnapshot(options)
  const json = `${JSON.stringify(snapshot, null, 2)}\n`

  if (options.outputPath !== undefined) {
    await writeFile(options.outputPath, json)
    console.log(`wrote ${options.outputPath}`)
    return
  }

  process.stdout.write(json)
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  await runCli()
}
