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
import { fetchBodyBytesWithTimeout } from './public-workbook-corpus-http.ts'
import { recentWorkbookDateEvidenceForFields } from './public-workbook-corpus-topics.ts'
import { sha256HexSync } from './public-workbook-corpus-workbook.ts'
import type { PublicWorkbookLicenseEvidence, PublicWorkbookManifest, PublicWorkbookSource } from './public-workbook-corpus-types.ts'

export const defaultRecentComplexFigshareQueries = [
  'analysis xlsx 2026',
  'analysis xlsx 2025',
  'results xlsx 2026',
  'results xlsx 2025',
  'model xlsx 2026',
  'model xlsx 2025',
  'financial xlsx 2026',
  'financial xlsx 2025',
  'forecast xlsx 2026',
  'forecast xlsx 2025',
  'budget xlsx 2026',
  'budget xlsx 2025',
  'workbook 2026',
  'workbook 2025',
  'spreadsheet 2026',
  'spreadsheet 2025',
  'xlsx 2026',
  'xlsx 2025',
] as const

interface DiscoverFigshareWorkbookSourcesArgs {
  readonly manifest: PublicWorkbookManifest
  readonly queries: readonly string[]
  readonly limit: number
  readonly pageSize: number
  readonly maxPagesPerQuery: number
  readonly discoveredAt?: string
  readonly onQueryDiscovered?: (manifest: PublicWorkbookManifest) => void
}

interface FigshareDiscoveryState {
  readonly existingKeys: Set<string>
  readonly existingSourceIds: Set<string>
  readonly sources: PublicWorkbookSource[]
}

export async function discoverRecentComplexFigshareQueries(args: DiscoverFigshareWorkbookSourcesArgs): Promise<PublicWorkbookManifest> {
  const [query, ...remainingQueries] = args.queries
  if (!query || args.manifest.sources.length >= args.limit) {
    return args.manifest
  }
  const discoveredManifest = await discoverFigshareWorkbookSources({
    ...args,
    manifest: args.manifest,
    queries: [query],
    onQueryDiscovered: undefined,
  })
  console.log(`Discovered ${String(discoveredManifest.sources.length)} recent Figshare workbook sources after query "${query}"`)
  args.onQueryDiscovered?.(discoveredManifest)
  return discoverRecentComplexFigshareQueries({ ...args, manifest: discoveredManifest, queries: remainingQueries })
}

export async function discoverFigshareWorkbookSources(args: DiscoverFigshareWorkbookSourcesArgs): Promise<PublicWorkbookManifest> {
  validatePublicWorkbookManifest(args.manifest)
  const discoveredAt = args.discoveredAt ?? new Date().toISOString()
  const state: FigshareDiscoveryState = {
    existingKeys: new Set(args.manifest.sources.map((source) => `${source.downloadUrl}|${source.license.evidenceUrl ?? ''}`)),
    existingSourceIds: new Set(args.manifest.sources.map((source) => source.id)),
    sources: [...args.manifest.sources],
  }
  await appendFigshareQuerySources({
    ...args,
    discoveredAt,
    queries: args.queries,
    state,
  })

  return {
    ...args.manifest,
    generatedAt: discoveredAt,
    sources: state.sources,
  }
}

async function appendFigshareQuerySources(
  args: DiscoverFigshareWorkbookSourcesArgs & {
    readonly discoveredAt: string
    readonly state: FigshareDiscoveryState
  },
): Promise<void> {
  const [query, ...remainingQueries] = args.queries
  if (!query || args.state.sources.length >= args.limit) {
    return
  }
  await appendFigshareQueryPageSources({
    ...args,
    query,
    pageNumber: 1,
  })
  await appendFigshareQuerySources({
    ...args,
    queries: remainingQueries,
  })
}

async function appendFigshareQueryPageSources(
  args: DiscoverFigshareWorkbookSourcesArgs & {
    readonly discoveredAt: string
    readonly state: FigshareDiscoveryState
    readonly query: string
    readonly pageNumber: number
  },
): Promise<void> {
  if (args.pageNumber > Math.max(1, args.maxPagesPerQuery) || args.state.sources.length >= args.limit) {
    return
  }
  const page = await fetchFigshareSearchPage({
    query: args.query,
    page: args.pageNumber,
    pageSize: args.pageSize,
  })
  await appendFigshareArticleSources({
    ...args,
    query: page.query,
    articles: page.articles,
  })
  if (page.articles.length === 0) {
    return
  }
  await appendFigshareQueryPageSources({
    ...args,
    pageNumber: args.pageNumber + 1,
  })
}

async function appendFigshareArticleSources(args: {
  readonly discoveredAt: string
  readonly limit: number
  readonly query: string
  readonly articles: readonly Record<string, unknown>[]
  readonly state: FigshareDiscoveryState
}): Promise<void> {
  const articles = await Promise.all(
    args.articles.flatMap((searchArticle) => {
      const articleUrl = readString(searchArticle, 'url_public_api') ?? readString(searchArticle, 'url')
      return articleUrl ? [fetchFigshareArticle(articleUrl)] : []
    }),
  )
  for (const article of articles) {
    if (args.state.sources.length >= args.limit) {
      return
    }
    for (const source of readFigshareWorkbookSources({ article, query: args.query, discoveredAt: args.discoveredAt })) {
      if (args.state.sources.length >= args.limit) {
        return
      }
      const key = `${source.downloadUrl}|${source.license.evidenceUrl ?? ''}`
      if (args.state.existingKeys.has(key)) {
        continue
      }
      args.state.existingKeys.add(key)
      if (args.state.existingSourceIds.has(source.id)) {
        continue
      }
      args.state.existingSourceIds.add(source.id)
      args.state.sources.push(source)
    }
  }
}

async function fetchFigshareSearchPage(args: {
  readonly query: string
  readonly page: number
  readonly pageSize: number
}): Promise<{ readonly query: string; readonly articles: readonly Record<string, unknown>[] }> {
  const url = new URL('https://api.figshare.com/v2/articles/search')
  try {
    const root = await fetchFigshareJsonWithRetry(url, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        search_for: args.query,
        page: Math.max(1, args.page),
        page_size: Math.max(1, Math.min(100, args.pageSize)),
        order: 'published_date',
        order_direction: 'desc',
      }),
    })
    return {
      query: args.query,
      articles: readFigshareRecords(root),
    }
  } catch {
    return {
      query: args.query,
      articles: [],
    }
  }
}

async function fetchFigshareArticle(articleUrl: string): Promise<Record<string, unknown>> {
  try {
    return asRecord(
      await fetchFigshareJsonWithRetry(new URL(articleUrl), {
        headers: {
          accept: 'application/json',
        },
      }),
    )
  } catch {
    return {}
  }
}

async function fetchFigshareJsonWithRetry(url: URL, init: RequestInit): Promise<unknown> {
  const maxAttempts = 2
  return fetchFigshareJsonAttempt(url, init, 1, maxAttempts)
}

async function fetchFigshareJsonAttempt(url: URL, init: RequestInit, attempt: number, maxAttempts: number): Promise<unknown> {
  const { bytes, response } = await fetchBodyBytesWithTimeout(
    url,
    {
      ...init,
      headers: figshareRequestHeaders(init.headers),
    },
    {
      timeoutMs: defaultDownloadTimeoutMs,
      maxBytes: 20 * 1024 * 1024,
      maxBytesLabel: 'Figshare API response',
    },
  )
  if (response.status === 429 && attempt < maxAttempts) {
    await sleep(readRetryAfterMs(response))
    return fetchFigshareJsonAttempt(url, init, attempt + 1, maxAttempts)
  }
  if (!response.ok) {
    throw new Error(`Unable to fetch ${url.href}: HTTP ${String(response.status)}`)
  }
  return JSON.parse(new TextDecoder().decode(bytes)) as unknown
}

function figshareRequestHeaders(initHeaders: HeadersInit | undefined): Record<string, string> {
  const headers = new Headers(initHeaders)
  headers.set('user-agent', 'bilig-public-workbook-corpus/1.0')
  if (!headers.has('accept')) {
    headers.set('accept', 'application/json')
  }
  return Object.fromEntries(headers.entries())
}

function readFigshareRecords(value: unknown): readonly Record<string, unknown>[] {
  if (!Array.isArray(value)) {
    return []
  }
  return value.flatMap((entry) => {
    const record = asRecordOrNull(entry)
    return record ? [record] : []
  })
}

function readFigshareWorkbookSources(args: {
  readonly article: Record<string, unknown>
  readonly query: string
  readonly discoveredAt: string
}): PublicWorkbookSource[] {
  const articleId = readFigshareArticleId(args.article)
  const sourceUrl =
    readString(args.article, 'url_public_html') ?? (articleId ? `https://figshare.com/articles/${encodeURIComponent(articleId)}` : null)
  if (!articleId || !sourceUrl) {
    return []
  }
  const license = readFigshareLicense(args.article, sourceUrl)
  if (!license || !hasUsableLicenseEvidence(license)) {
    return []
  }
  const articleFields = readFigshareArticleDateFields(args.article, articleId, sourceUrl)
  return readArray(args.article, 'files').flatMap((file) => {
    const fileRecord = asRecordOrNull(file)
    const fileName = fileRecord ? readString(fileRecord, 'name') : null
    const downloadUrl = fileRecord ? readString(fileRecord, 'download_url') : null
    const isLinkOnly = fileRecord ? fileRecord['is_link_only'] === true : false
    if (!fileName || !downloadUrl || isLinkOnly || !isSpreadsheetFileName(fileName)) {
      return []
    }
    const topicEvidence = [
      ...recentWorkbookDateEvidenceForFields([
        ...articleFields,
        { name: 'figshare.fileName', value: fileName },
        { name: 'downloadUrl', value: downloadUrl },
      ]),
      `figshare-query:${stableId(args.query)}`,
    ]
    if (!topicEvidence.some((evidence) => evidence.startsWith('recent-2025:') || evidence.startsWith('recent-2026:'))) {
      return []
    }
    return [
      {
        id: `figshare-${stableId(`${articleId}:${fileName}:${downloadUrl}`)}`,
        kind: 'direct-url' as const,
        sourceUrl,
        downloadUrl,
        fileName,
        discoveredAt: args.discoveredAt,
        license,
        topicEvidence,
      },
    ]
  })
}

function readFigshareArticleDateFields(
  article: Record<string, unknown>,
  articleId: string,
  sourceUrl: string,
): { readonly name: string; readonly value: string }[] {
  const timeline = asRecordOrNull(article['timeline'])
  return [
    { name: 'figshare.title', value: readString(article, 'title') },
    { name: 'figshare.publishedDate', value: readString(article, 'published_date') },
    { name: 'figshare.createdDate', value: readString(article, 'created_date') },
    { name: 'figshare.modifiedDate', value: readString(article, 'modified_date') },
    { name: 'figshare.timelinePosted', value: timeline ? readString(timeline, 'posted') : null },
    { name: 'figshare.timelineFirstOnline', value: timeline ? readString(timeline, 'firstOnline') : null },
    { name: 'figshare.articleId', value: articleId },
    { name: 'sourceUrl', value: sourceUrl },
  ].flatMap((field) => (field.value ? [{ name: field.name, value: field.value }] : []))
}

function readFigshareArticleId(article: Record<string, unknown>): string | null {
  const raw = article['id']
  if (typeof raw === 'string') {
    return raw
  }
  if (typeof raw === 'number' && Number.isSafeInteger(raw)) {
    return String(raw)
  }
  return null
}

function readFigshareLicense(article: Record<string, unknown>, sourceUrl: string): PublicWorkbookLicenseEvidence | null {
  const license = asRecordOrNull(article['license'])
  const title = license ? (readString(license, 'name') ?? readString(license, 'title')) : null
  if (!title) {
    return null
  }
  return {
    spdxId: normalizeFigshareSpdx(title),
    title,
    evidenceUrl: readString(license, 'url') ?? sourceUrl,
  }
}

function normalizeFigshareSpdx(title: string): string | null {
  const normalized = title.trim().toLowerCase()
  if (normalized === 'cc by 4.0' || normalized === 'cc-by-4.0') {
    return 'CC-BY-4.0'
  }
  if (normalized === 'cc0' || normalized === 'cc0 1.0') {
    return 'CC0-1.0'
  }
  return null
}

function stableId(value: string): string {
  return sha256HexSync(Buffer.from(value)).slice(0, 16)
}

function readRetryAfterMs(response: Response): number {
  const raw = response.headers.get('retry-after')
  const seconds = raw ? Number(raw) : NaN
  if (Number.isFinite(seconds) && seconds > 0) {
    return Math.min(65_000, Math.ceil(seconds * 1_000))
  }
  return 30_000
}

async function sleep(ms: number): Promise<void> {
  await new Promise<void>((resolve) => {
    setTimeout(resolve, ms)
  })
}
