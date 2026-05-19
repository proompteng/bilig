import {
  asRecord,
  asRecordOrNull,
  hasUsableLicenseEvidence,
  isSpreadsheetFileName,
  readArray,
  readString,
  validatePublicWorkbookManifest,
} from './public-workbook-corpus-json.ts'
import { defaultDownloadTimeoutMs } from './public-workbook-corpus-fetch.ts'
import { fetchJsonWithTimeout } from './public-workbook-corpus-http.ts'
import { recentWorkbookDateEvidenceForFields } from './public-workbook-corpus-topics.ts'
import { sha256HexSync } from './public-workbook-corpus-workbook.ts'
import type { PublicWorkbookLicenseEvidence, PublicWorkbookManifest, PublicWorkbookSource } from './public-workbook-corpus-types.ts'

export const defaultRecentComplexGithubQueries = [
  '2026 model extension:xlsx',
  '2026 template extension:xlsx',
  '2026 forecast extension:xlsx',
  '2026 budget extension:xlsx',
  '2026 financial extension:xlsx',
  '2026 planning extension:xlsx',
  '2026 calculator extension:xlsx',
  '2026 model extension:xlsm',
  '2026 template extension:xlsm',
  '2026 forecast extension:xlsm',
  '2026 budget extension:xlsm',
  '2025 model extension:xlsx',
  '2025 template extension:xlsx',
  '2025 forecast extension:xlsx',
  '2025 budget extension:xlsx',
  '2025 financial extension:xlsx',
  '2025 planning extension:xlsx',
  '2025 calculator extension:xlsx',
  '2025 model extension:xlsm',
  '2025 template extension:xlsm',
  '2025 forecast extension:xlsm',
  '2025 budget extension:xlsm',
] as const

export const defaultRecentComplexGithubRepositoryQueries = [
  'excel financial modeling license:mit',
  'financial model excel license:mit',
  'project finance excel license:mit',
  'financial forecast excel license:mit',
  'dcf excel model license:mit',
  'budget model excel license:mit',
  'valuation model excel license:mit',
  '3 statement model excel license:mit',
  'bond valuation excel license:mit',
  'corporate finance model excel license:mit',
  'excel dashboard 2025 license:mit',
  'excel dashboard 2026 license:mit',
  'excel dashboard license:mit',
  'financial analysis excel 2025 license:mit',
  'financial analysis excel 2026 license:mit',
  'financial analysis excel license:mit',
  'financial model 2025 excel license:mit',
  'financial model 2026 excel license:mit',
  'financial modeling course excel license:mit',
  'investment banking excel model license:mit',
  'm&a valuation excel license:mit',
  'portfolio optimization excel license:mit',
  'real estate financial model excel license:mit',
  'saas financial model excel license:mit',
  'startup valuation excel license:mit',
  'actuarial excel model license:mit',
  'excel financial modeling license:apache-2.0',
  'excel dashboard license:apache-2.0',
  'financial analysis excel license:apache-2.0',
  'financial model excel license:apache-2.0',
  'project finance excel license:apache-2.0',
  'dcf excel model license:apache-2.0',
  '3 statement model excel license:apache-2.0',
  'bond valuation excel license:apache-2.0',
  'corporate finance model excel license:apache-2.0',
  'financial modeling course excel license:apache-2.0',
  'investment banking excel model license:apache-2.0',
  'm&a valuation excel license:apache-2.0',
  'real estate financial model excel license:apache-2.0',
  'saas financial model excel license:apache-2.0',
  'startup valuation excel license:apache-2.0',
  '2026 xlsx license:mit',
  '2026 excel license:mit',
  '2026 workbook license:mit',
  '2026 xlsm license:mit',
  '2025 xlsx license:mit',
  '2025 excel license:mit',
  '2025 workbook license:mit',
  '2025 xlsm license:mit',
  '2026 xlsx license:apache-2.0',
  '2026 excel license:apache-2.0',
  '2025 xlsx license:apache-2.0',
  '2025 excel license:apache-2.0',
] as const

const maxGithubWorkbookPathsPerRepository = 50

interface DiscoverGithubWorkbookSourcesArgs {
  readonly manifest: PublicWorkbookManifest
  readonly queries: readonly string[]
  readonly repositoryQueries?: readonly string[]
  readonly limit: number
  readonly perPage: number
  readonly maxPagesPerQuery: number
  readonly maxRepositoryPagesPerQuery?: number
  readonly maxRepositoriesPerQuery?: number
  readonly githubToken?: string | null
  readonly discoveredAt?: string
  readonly onQueryDiscovered?: (manifest: PublicWorkbookManifest) => void
}

interface GithubCodeSearchPage {
  readonly query: string
  readonly items: readonly Record<string, unknown>[]
}

interface GithubWorkbookPathCandidate {
  readonly path: string
  readonly fileName: string
}

export async function discoverRecentComplexGithubQueries(args: DiscoverGithubWorkbookSourcesArgs): Promise<PublicWorkbookManifest> {
  const codeDiscoveredManifest = await discoverGithubCodeQueriesSequential(args, args.manifest, args.queries)
  return discoverGithubRepositoryQueriesSequential(
    args,
    codeDiscoveredManifest,
    args.repositoryQueries ?? defaultRecentComplexGithubRepositoryQueries,
  )
}

export async function discoverGithubWorkbookSources(args: DiscoverGithubWorkbookSourcesArgs): Promise<PublicWorkbookManifest> {
  validatePublicWorkbookManifest(args.manifest)
  const discoveredAt = args.discoveredAt ?? new Date().toISOString()
  const existingKeys = new Set(args.manifest.sources.map((source) => `${source.downloadUrl}|${source.license.evidenceUrl ?? ''}`))
  const existingSourceIds = new Set(args.manifest.sources.map((source) => source.id))
  const sources: PublicWorkbookSource[] = [...args.manifest.sources]
  const licenseCache = new Map<string, PublicWorkbookLicenseEvidence | null>()
  const pageRequests = args.queries.flatMap((query) =>
    Array.from({ length: args.maxPagesPerQuery }, (_, index) => ({
      query,
      page: index + 1,
      perPage: args.perPage,
      githubToken: args.githubToken,
    })),
  )
  const searchPages = await Promise.all(pageRequests.map(fetchGithubCodeSearchPage))
  const sourceResults = await mapWithConcurrency(
    searchPages.flatMap((page) => page.items.map((item) => ({ item, query: page.query }))),
    4,
    (entry) =>
      readGithubWorkbookSource({
        item: entry.item,
        query: entry.query,
        discoveredAt,
        githubToken: args.githubToken,
        licenseCache,
      }),
  )
  for (const source of sourceResults) {
    if (!source || sources.length >= args.limit) {
      continue
    }
    const key = `${source.downloadUrl}|${source.license.evidenceUrl ?? ''}`
    if (existingKeys.has(key)) {
      continue
    }
    existingKeys.add(key)
    if (existingSourceIds.has(source.id)) {
      continue
    }
    existingSourceIds.add(source.id)
    sources.push(source)
  }

  return {
    ...args.manifest,
    generatedAt: discoveredAt,
    sources,
  }
}

async function discoverGithubCodeQueriesSequential(
  args: DiscoverGithubWorkbookSourcesArgs,
  manifest: PublicWorkbookManifest,
  queries: readonly string[],
): Promise<PublicWorkbookManifest> {
  const [query, ...remainingQueries] = queries
  if (!query || manifest.sources.length >= args.limit) {
    return manifest
  }
  const discoveredManifest = await discoverGithubWorkbookSources({
    ...args,
    manifest,
    queries: [query],
    onQueryDiscovered: undefined,
  })
  console.log(`Discovered ${String(discoveredManifest.sources.length)} recent GitHub workbook sources after query "${query}"`)
  args.onQueryDiscovered?.(discoveredManifest)
  return discoverGithubCodeQueriesSequential(args, discoveredManifest, remainingQueries)
}

async function discoverGithubRepositoryQueriesSequential(
  args: DiscoverGithubWorkbookSourcesArgs,
  manifest: PublicWorkbookManifest,
  queries: readonly string[],
): Promise<PublicWorkbookManifest> {
  const [query, ...remainingQueries] = queries
  if (!query || manifest.sources.length >= args.limit) {
    return manifest
  }
  const discoveredManifest = await discoverGithubRepositoryWorkbookSources({
    ...args,
    manifest,
    query,
  })
  console.log(`Discovered ${String(discoveredManifest.sources.length)} recent GitHub workbook sources after repository query "${query}"`)
  args.onQueryDiscovered?.(discoveredManifest)
  return discoverGithubRepositoryQueriesSequential(args, discoveredManifest, remainingQueries)
}

async function discoverGithubRepositoryWorkbookSources(
  args: Omit<DiscoverGithubWorkbookSourcesArgs, 'queries'> & { readonly query: string },
): Promise<PublicWorkbookManifest> {
  validatePublicWorkbookManifest(args.manifest)
  const discoveredAt = args.discoveredAt ?? new Date().toISOString()
  const existingKeys = new Set(args.manifest.sources.map((source) => `${source.downloadUrl}|${source.license.evidenceUrl ?? ''}`))
  const existingSourceIds = new Set(args.manifest.sources.map((source) => source.id))
  const sources: PublicWorkbookSource[] = [...args.manifest.sources]
  const licenseCache = new Map<string, PublicWorkbookLicenseEvidence | null>()
  const repositoryPages = await Promise.all(
    Array.from({ length: Math.max(1, args.maxRepositoryPagesPerQuery ?? 1) }, (_, index) =>
      fetchGithubRepositorySearchPage({
        query: args.query,
        page: index + 1,
        perPage: args.maxRepositoriesPerQuery ?? args.perPage,
        githubToken: args.githubToken,
      }),
    ),
  )
  const repositories = repositoryPages.flat()
  const repositorySourceBatches = await mapWithConcurrency(repositories, 2, (repository) =>
    readGithubRepositoryWorkbookSources({
      repository,
      query: args.query,
      discoveredAt,
      githubToken: args.githubToken,
      licenseCache,
      remainingSourceSlots: args.limit - sources.length,
    }),
  )

  for (const repositorySources of repositorySourceBatches) {
    for (const source of repositorySources) {
      const key = `${source.downloadUrl}|${source.license.evidenceUrl ?? ''}`
      if (existingKeys.has(key)) {
        continue
      }
      existingKeys.add(key)
      if (existingSourceIds.has(source.id)) {
        continue
      }
      existingSourceIds.add(source.id)
      sources.push(source)
      if (sources.length >= args.limit) {
        break
      }
    }
  }

  return {
    ...args.manifest,
    generatedAt: discoveredAt,
    sources,
  }
}

async function fetchGithubCodeSearchPage(args: {
  readonly query: string
  readonly page: number
  readonly perPage: number
  readonly githubToken?: string | null
}): Promise<GithubCodeSearchPage> {
  const url = new URL('https://api.github.com/search/code')
  url.searchParams.set('q', args.query)
  url.searchParams.set('page', String(args.page))
  url.searchParams.set('per_page', String(Math.max(1, Math.min(100, args.perPage))))
  try {
    const response = await fetchGithubJson(url, args.githubToken)
    const root = asRecord(response)
    return {
      query: args.query,
      items: readArray(root, 'items').flatMap((item) => {
        const record = asRecordOrNull(item)
        return record ? [record] : []
      }),
    }
  } catch {
    return {
      query: args.query,
      items: [],
    }
  }
}

async function fetchGithubRepositorySearchPage(args: {
  readonly query: string
  readonly page: number
  readonly perPage: number
  readonly githubToken?: string | null
}): Promise<Record<string, unknown>[]> {
  const url = new URL('https://api.github.com/search/repositories')
  url.searchParams.set('q', args.query)
  url.searchParams.set('page', String(Math.max(1, args.page)))
  url.searchParams.set('per_page', String(Math.max(1, Math.min(100, args.perPage))))
  try {
    const response = await fetchGithubJson(url, args.githubToken)
    const root = asRecord(response)
    return readArray(root, 'items').flatMap((item) => {
      const record = asRecordOrNull(item)
      return record ? [record] : []
    })
  } catch {
    return []
  }
}

async function readGithubWorkbookSource(args: {
  readonly item: Record<string, unknown>
  readonly query: string
  readonly discoveredAt: string
  readonly githubToken?: string | null
  readonly licenseCache: Map<string, PublicWorkbookLicenseEvidence | null>
}): Promise<PublicWorkbookSource | null> {
  const fileName = readString(args.item, 'name') ?? fileNameFromPath(readString(args.item, 'path') ?? '')
  if (!fileName || !isSpreadsheetFileName(fileName)) {
    return null
  }
  const repository = asRecordOrNull(args.item['repository'])
  const repositoryFullName = repository ? readString(repository, 'full_name') : null
  if (!repositoryFullName) {
    return null
  }
  const sourceUrl = readString(args.item, 'html_url')
  const contentsUrl = readString(args.item, 'url')
  const path = readString(args.item, 'path') ?? fileName
  if (!sourceUrl || !contentsUrl) {
    return null
  }
  const downloadUrl = await readGithubContentDownloadUrl(contentsUrl, fileName, args.githubToken)
  if (!downloadUrl) {
    return null
  }
  const topicEvidence = [
    ...recentWorkbookDateEvidenceForFields([
      { name: 'github.path', value: path },
      { name: 'sourceUrl', value: sourceUrl },
      { name: 'downloadUrl', value: downloadUrl },
      { name: 'fileName', value: fileName },
    ]),
    `github-query:${stableId(args.query)}`,
  ]
  if (!topicEvidence.some((evidence) => evidence.startsWith('recent-2025:') || evidence.startsWith('recent-2026:'))) {
    return null
  }
  const license = await readGithubRepositoryLicense(repositoryFullName, args.githubToken, args.licenseCache)
  if (!license || !hasUsableLicenseEvidence(license)) {
    return null
  }
  return {
    id: `github-${stableId(`${repositoryFullName}:${path}:${downloadUrl}`)}`,
    kind: 'github-contents',
    sourceUrl,
    downloadUrl,
    fileName,
    discoveredAt: args.discoveredAt,
    license,
    topicEvidence,
  }
}

async function readGithubRepositoryWorkbookSources(args: {
  readonly repository: Record<string, unknown>
  readonly query: string
  readonly discoveredAt: string
  readonly githubToken?: string | null
  readonly licenseCache: Map<string, PublicWorkbookLicenseEvidence | null>
  readonly remainingSourceSlots: number
}): Promise<PublicWorkbookSource[]> {
  const repositoryFullName = readString(args.repository, 'full_name')
  const repositoryHtmlUrl = readString(args.repository, 'html_url')
  const defaultBranch = readString(args.repository, 'default_branch') ?? 'main'
  if (!repositoryFullName || !repositoryHtmlUrl) {
    return []
  }
  const license = await readGithubRepositoryLicense(repositoryFullName, args.githubToken, args.licenseCache)
  if (!license || !hasUsableLicenseEvidence(license)) {
    return []
  }
  const repositoryDateFields = githubRepositoryDateFields(args.repository)
  const treeEntries = await readGithubRepositoryTree(repositoryFullName, defaultBranch, args.githubToken)
  const candidates = treeEntries
    .flatMap((entry) => {
      const path = readString(entry, 'path')
      const type = readString(entry, 'type')
      const fileName = path ? fileNameFromPath(path) : null
      if (type !== 'blob' || !path || !fileName || !isSpreadsheetFileName(fileName)) {
        return []
      }
      return [{ path, fileName }]
    })
    .toSorted((left, right) => rankGithubWorkbookPath(right.path) - rankGithubWorkbookPath(left.path))
    .slice(0, Math.min(args.remainingSourceSlots, maxGithubWorkbookPathsPerRepository))
  const sources = await mapWithConcurrency(candidates, 4, (candidate) =>
    readGithubRepositoryWorkbookSource({
      ...args,
      repositoryFullName,
      repositoryHtmlUrl,
      defaultBranch,
      license,
      candidate,
      repositoryDateFields,
    }),
  )
  return sources.flatMap((source) => (source ? [source] : []))
}

async function readGithubRepositoryWorkbookSource(args: {
  readonly repositoryFullName: string
  readonly repositoryHtmlUrl: string
  readonly defaultBranch: string
  readonly license: PublicWorkbookLicenseEvidence
  readonly candidate: GithubWorkbookPathCandidate
  readonly query: string
  readonly discoveredAt: string
  readonly githubToken?: string | null
  readonly repositoryDateFields: readonly { readonly name: string; readonly value: string }[]
}): Promise<PublicWorkbookSource | null> {
  const { path, fileName } = args.candidate
  const encodedPath = encodeGithubPath(path)
  const downloadUrl = `https://raw.githubusercontent.com/${args.repositoryFullName}/${encodeURIComponent(args.defaultBranch)}/${encodedPath}`
  const sourceUrl = `${args.repositoryHtmlUrl}/blob/${encodeURIComponent(args.defaultBranch)}/${encodedPath}`
  const dateFields = [
    { name: 'github.path', value: path },
    { name: 'sourceUrl', value: sourceUrl },
    { name: 'downloadUrl', value: downloadUrl },
    { name: 'fileName', value: fileName },
  ]
  const pathDateEvidence = recentWorkbookDateEvidenceForFields([...args.repositoryDateFields, ...dateFields])
  const commitDate =
    pathDateEvidence.length > 0
      ? null
      : await readGithubPathLatestCommitDate(args.repositoryFullName, args.defaultBranch, path, args.githubToken)
  const topicEvidence = [
    ...recentWorkbookDateEvidenceForFields(
      commitDate
        ? [...args.repositoryDateFields, ...dateFields, { name: 'github.commitDate', value: commitDate }]
        : [...args.repositoryDateFields, ...dateFields],
    ),
    `github-repo-query:${stableId(args.query)}`,
  ]
  if (!topicEvidence.some((evidence) => evidence.startsWith('recent-2025:') || evidence.startsWith('recent-2026:'))) {
    return null
  }
  return {
    id: `github-${stableId(`${args.repositoryFullName}:${path}:${downloadUrl}`)}`,
    kind: 'github-contents',
    sourceUrl,
    downloadUrl,
    fileName,
    discoveredAt: args.discoveredAt,
    license: args.license,
    topicEvidence,
  }
}

function githubRepositoryDateFields(repository: Record<string, unknown>): { readonly name: string; readonly value: string }[] {
  return [
    { name: 'github.repositoryFullName', value: readString(repository, 'full_name') },
    { name: 'github.repositoryName', value: readString(repository, 'name') },
    { name: 'github.repositoryDescription', value: readString(repository, 'description') },
    { name: 'github.repositoryUrl', value: readString(repository, 'html_url') },
  ].flatMap((field) => (field.value ? [{ name: field.name, value: field.value }] : []))
}

async function readGithubRepositoryTree(
  repositoryFullName: string,
  defaultBranch: string,
  githubToken?: string | null,
): Promise<Record<string, unknown>[]> {
  try {
    const url = new URL(`https://api.github.com/repos/${repositoryFullName}/git/trees/${encodeURIComponent(defaultBranch)}`)
    url.searchParams.set('recursive', '1')
    const root = asRecord(await fetchGithubJson(url, githubToken))
    return readArray(root, 'tree').flatMap((entry) => {
      const record = asRecordOrNull(entry)
      return record ? [record] : []
    })
  } catch {
    return []
  }
}

async function readGithubPathLatestCommitDate(
  repositoryFullName: string,
  defaultBranch: string,
  path: string,
  githubToken?: string | null,
): Promise<string | null> {
  try {
    const url = new URL(`https://api.github.com/repos/${repositoryFullName}/commits`)
    url.searchParams.set('path', path)
    url.searchParams.set('per_page', '1')
    url.searchParams.set('sha', defaultBranch)
    const commits = await fetchGithubJson(url, githubToken)
    if (!Array.isArray(commits)) {
      return null
    }
    const firstCommit = asRecordOrNull(commits[0])
    const commit = firstCommit ? asRecordOrNull(firstCommit['commit']) : null
    const committer = commit ? asRecordOrNull(commit['committer']) : null
    const author = commit ? asRecordOrNull(commit['author']) : null
    return (committer ? readString(committer, 'date') : null) ?? (author ? readString(author, 'date') : null)
  } catch {
    return null
  }
}

async function readGithubContentDownloadUrl(contentsUrl: string, fileName: string, githubToken?: string | null): Promise<string | null> {
  try {
    const content = asRecord(await fetchGithubJson(new URL(contentsUrl), githubToken))
    if (!githubInlineContentMatchesSpreadsheetContainer(fileName, content)) {
      return null
    }
    return readString(content, 'download_url')
  } catch {
    return null
  }
}

function githubInlineContentMatchesSpreadsheetContainer(fileName: string, content: Record<string, unknown>): boolean {
  if (!requiresZipContainerMagic(fileName)) {
    return true
  }
  if (readString(content, 'encoding') !== 'base64') {
    return true
  }
  const encodedContent = readString(content, 'content')
  if (!encodedContent) {
    return true
  }
  const decoded = Buffer.from(encodedContent.replace(/\s+/gu, ''), 'base64')
  return decoded.length < 2 || (decoded[0] === 0x50 && decoded[1] === 0x4b)
}

function requiresZipContainerMagic(fileName: string): boolean {
  return /\.(?:xlsx|xlsm|xltx|xltm|ods)$/iu.test(fileName)
}

async function readGithubRepositoryLicense(
  repositoryFullName: string,
  githubToken: string | null | undefined,
  cache: Map<string, PublicWorkbookLicenseEvidence | null>,
): Promise<PublicWorkbookLicenseEvidence | null> {
  const cached = cache.get(repositoryFullName)
  if (cached !== undefined) {
    return cached
  }
  try {
    const url = new URL(`https://api.github.com/repos/${repositoryFullName}/license`)
    const root = asRecord(await fetchGithubJson(url, githubToken))
    const licenseRecord = asRecordOrNull(root['license'])
    const spdxId = licenseRecord ? readString(licenseRecord, 'spdx_id') : null
    const title = licenseRecord ? (readString(licenseRecord, 'name') ?? readString(licenseRecord, 'key') ?? spdxId ?? '') : ''
    const license: PublicWorkbookLicenseEvidence = {
      spdxId: spdxId && spdxId !== 'NOASSERTION' ? spdxId : null,
      title,
      evidenceUrl: readString(root, 'html_url'),
    }
    const usableLicense = hasUsableLicenseEvidence(license) ? license : null
    cache.set(repositoryFullName, usableLicense)
    return usableLicense
  } catch {
    cache.set(repositoryFullName, null)
    return null
  }
}

async function fetchGithubJson(url: URL, githubToken?: string | null): Promise<unknown> {
  return fetchJsonWithTimeout(
    url,
    {
      headers: {
        'user-agent': 'bilig-public-workbook-corpus/1.0',
        accept: 'application/vnd.github+json',
        'x-github-api-version': '2022-11-28',
        ...(githubToken ? { authorization: `Bearer ${githubToken}` } : {}),
      },
    },
    {
      timeoutMs: defaultDownloadTimeoutMs,
      maxBytes: 20 * 1024 * 1024,
      maxBytesLabel: 'GitHub API response',
      validateResponse: (response) => {
        if (!response.ok) {
          throw new Error(`Unable to fetch ${url.href}: HTTP ${String(response.status)}`)
        }
      },
    },
  )
}

function fileNameFromPath(path: string): string | null {
  const fileName = path.split('/').at(-1)
  return fileName && fileName.trim().length > 0 ? fileName : null
}

function rankGithubWorkbookPath(path: string): number {
  const normalized = path.toLowerCase()
  let score = 0
  if (/\b202[56]\b/u.test(normalized)) {
    score += 40
  }
  if (/\b(financial|finance|model|modelling|modeling|forecast|projection|budget|valuation|dcf|proforma|pro-forma)\b/u.test(normalized)) {
    score += 30
  }
  if (/\b(lambda|formula|calculator|scenario|sensitivity|cohort|cash[-_ ]?flow|irr|npv)\b/u.test(normalized)) {
    score += 20
  }
  if (/\b(template|sample|example|case|chapter|exercise)\b/u.test(normalized)) {
    score += 10
  }
  if (/\b(raw|export|output|results|dataset|data)\b/u.test(normalized)) {
    score -= 10
  }
  return score
}

function encodeGithubPath(path: string): string {
  return path.split('/').map(encodeURIComponent).join('/')
}

function stableId(value: string): string {
  return sha256HexSync(Buffer.from(value)).slice(0, 16)
}

async function mapWithConcurrency<T, R>(
  items: readonly T[],
  concurrency: number,
  mapper: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results: R[] = []
  let nextIndex = 0
  const workerCount = Math.min(Math.max(1, concurrency), items.length)
  const runNext = async (): Promise<void> => {
    const index = nextIndex
    nextIndex += 1
    if (index >= items.length) {
      return
    }
    results[index] = await mapper(items[index], index)
    await runNext()
  }
  await Promise.all(Array.from({ length: workerCount }, () => runNext()))
  return results
}
