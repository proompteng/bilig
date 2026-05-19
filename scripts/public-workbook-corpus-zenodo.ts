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

export const defaultRecentComplexZenodoQueries = [
  'filetype:xlsx publication_date:2026',
  'filetype:xlsx publication_date:2025',
  'filetype:xlsx 2026 workbook',
  'filetype:xlsx 2026 spreadsheet',
  'filetype:xlsx 2026 financial',
  'filetype:xlsx 2026 forecast',
  'filetype:xlsx 2026 budget',
  'filetype:xlsx 2025 workbook',
  'filetype:xlsx 2025 spreadsheet',
  'filetype:xlsx 2025 financial',
  'filetype:xlsx 2025 forecast',
  'filetype:xlsx 2025 budget',
] as const

interface DiscoverZenodoWorkbookSourcesArgs {
  readonly manifest: PublicWorkbookManifest
  readonly queries: readonly string[]
  readonly limit: number
  readonly perPage: number
  readonly maxPagesPerQuery: number
  readonly discoveredAt?: string
  readonly onQueryDiscovered?: (manifest: PublicWorkbookManifest) => void
}

export async function discoverRecentComplexZenodoQueries(args: DiscoverZenodoWorkbookSourcesArgs): Promise<PublicWorkbookManifest> {
  const [query, ...remainingQueries] = args.queries
  if (!query || args.manifest.sources.length >= args.limit) {
    return args.manifest
  }
  const discoveredManifest = await discoverZenodoWorkbookSources({
    ...args,
    manifest: args.manifest,
    queries: [query],
    onQueryDiscovered: undefined,
  })
  console.log(`Discovered ${String(discoveredManifest.sources.length)} recent Zenodo workbook sources after query "${query}"`)
  args.onQueryDiscovered?.(discoveredManifest)
  return discoverRecentComplexZenodoQueries({ ...args, manifest: discoveredManifest, queries: remainingQueries })
}

export async function discoverZenodoWorkbookSources(args: DiscoverZenodoWorkbookSourcesArgs): Promise<PublicWorkbookManifest> {
  validatePublicWorkbookManifest(args.manifest)
  const discoveredAt = args.discoveredAt ?? new Date().toISOString()
  const state: ZenodoDiscoveryState = {
    existingKeys: new Set(args.manifest.sources.map((source) => `${source.downloadUrl}|${source.license.evidenceUrl ?? ''}`)),
    existingSourceIds: new Set(args.manifest.sources.map((source) => source.id)),
    sources: [...args.manifest.sources],
  }
  await appendZenodoQuerySources({
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

interface ZenodoDiscoveryState {
  readonly existingKeys: Set<string>
  readonly existingSourceIds: Set<string>
  readonly sources: PublicWorkbookSource[]
}

async function appendZenodoQuerySources(
  args: DiscoverZenodoWorkbookSourcesArgs & {
    readonly discoveredAt: string
    readonly state: ZenodoDiscoveryState
  },
): Promise<void> {
  const [query, ...remainingQueries] = args.queries
  if (!query || args.state.sources.length >= args.limit) {
    return
  }
  await appendZenodoQueryPageSources({
    ...args,
    query,
    pageNumber: 1,
  })
  await appendZenodoQuerySources({
    ...args,
    queries: remainingQueries,
  })
}

async function appendZenodoQueryPageSources(
  args: DiscoverZenodoWorkbookSourcesArgs & {
    readonly discoveredAt: string
    readonly state: ZenodoDiscoveryState
    readonly query: string
    readonly pageNumber: number
  },
): Promise<void> {
  if (args.pageNumber > Math.max(1, args.maxPagesPerQuery) || args.state.sources.length >= args.limit) {
    return
  }
  const page = await fetchZenodoSearchPage({
    query: args.query,
    page: args.pageNumber,
    perPage: args.perPage,
  })
  appendZenodoRecordSources({
    ...args,
    query: page.query,
    records: page.records,
  })
  await appendZenodoQueryPageSources({
    ...args,
    pageNumber: args.pageNumber + 1,
  })
}

function appendZenodoRecordSources(args: {
  readonly discoveredAt: string
  readonly limit: number
  readonly query: string
  readonly records: readonly Record<string, unknown>[]
  readonly state: ZenodoDiscoveryState
}): void {
  for (const record of args.records) {
    const recordSources = readZenodoWorkbookSources({
      record,
      query: args.query,
      discoveredAt: args.discoveredAt,
    })
    for (const source of recordSources) {
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

async function fetchZenodoSearchPage(args: {
  readonly query: string
  readonly page: number
  readonly perPage: number
}): Promise<{ readonly query: string; readonly records: readonly Record<string, unknown>[] }> {
  const url = new URL('https://zenodo.org/api/records')
  url.searchParams.set('q', args.query)
  url.searchParams.set('sort', 'mostrecent')
  url.searchParams.set('page', String(Math.max(1, args.page)))
  url.searchParams.set('size', String(Math.max(1, Math.min(100, args.perPage))))
  try {
    const root = asRecord(await fetchZenodoJsonWithRetry(url))
    const hits = asRecordOrNull(root['hits'])
    return {
      query: args.query,
      records: hits
        ? readArray(hits, 'hits').flatMap((record) => {
            const parsed = asRecordOrNull(record)
            return parsed ? [parsed] : []
          })
        : [],
    }
  } catch {
    return {
      query: args.query,
      records: [],
    }
  }
}

async function fetchZenodoJsonWithRetry(url: URL): Promise<unknown> {
  const maxAttempts = 2
  return fetchZenodoJsonAttempt(url, 1, maxAttempts)
}

async function fetchZenodoJsonAttempt(url: URL, attempt: number, maxAttempts: number): Promise<unknown> {
  const { bytes, response } = await fetchBodyBytesWithTimeout(
    url,
    {
      headers: {
        'user-agent': 'bilig-public-workbook-corpus/1.0',
        accept: 'application/json',
      },
    },
    {
      timeoutMs: defaultDownloadTimeoutMs,
      maxBytes: 20 * 1024 * 1024,
      maxBytesLabel: 'Zenodo API response',
    },
  )
  if (response.status === 429 && attempt < maxAttempts) {
    await sleep(readRetryAfterMs(response))
    return fetchZenodoJsonAttempt(url, attempt + 1, maxAttempts)
  }
  if (!response.ok) {
    throw new Error(`Unable to fetch ${url.href}: HTTP ${String(response.status)}`)
  }
  return JSON.parse(new TextDecoder().decode(bytes)) as unknown
}

function readZenodoWorkbookSources(args: {
  readonly record: Record<string, unknown>
  readonly query: string
  readonly discoveredAt: string
}): PublicWorkbookSource[] {
  const metadata = asRecordOrNull(args.record['metadata'])
  const links = asRecordOrNull(args.record['links'])
  const recordId = readZenodoRecordId(args.record)
  const sourceUrl =
    readString(links ?? {}, 'html') ??
    readString(links ?? {}, 'self_html') ??
    (recordId ? `https://zenodo.org/records/${encodeURIComponent(recordId)}` : null)
  if (!metadata || !sourceUrl || !recordId) {
    return []
  }
  const license = readZenodoLicense(metadata, sourceUrl)
  if (!license || !hasUsableLicenseEvidence(license)) {
    return []
  }
  const recordFields = [
    { name: 'zenodo.title', value: readString(metadata, 'title') },
    { name: 'zenodo.publicationDate', value: readString(metadata, 'publication_date') },
    { name: 'zenodo.recordId', value: recordId },
    { name: 'sourceUrl', value: sourceUrl },
  ].flatMap((field) => (field.value ? [{ name: field.name, value: field.value }] : []))
  return readArray(args.record, 'files').flatMap((file) => {
    const fileRecord = asRecordOrNull(file)
    const fileLinks = fileRecord ? asRecordOrNull(fileRecord['links']) : null
    const fileName = fileRecord ? readString(fileRecord, 'key') : null
    const downloadUrl = fileLinks ? readString(fileLinks, 'self') : null
    if (!fileName || !downloadUrl || !isSpreadsheetFileName(fileName)) {
      return []
    }
    const topicEvidence = [
      ...recentWorkbookDateEvidenceForFields([
        ...recordFields,
        { name: 'zenodo.fileName', value: fileName },
        { name: 'downloadUrl', value: downloadUrl },
      ]),
      `zenodo-query:${stableId(args.query)}`,
    ]
    if (!topicEvidence.some((evidence) => evidence.startsWith('recent-2025:') || evidence.startsWith('recent-2026:'))) {
      return []
    }
    return [
      {
        id: `zenodo-${stableId(`${recordId}:${fileName}:${downloadUrl}`)}`,
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

function readZenodoRecordId(record: Record<string, unknown>): string | null {
  const raw = record['id']
  if (typeof raw === 'string') {
    return raw
  }
  if (typeof raw === 'number' && Number.isSafeInteger(raw)) {
    return String(raw)
  }
  return null
}

function readZenodoLicense(metadata: Record<string, unknown>, sourceUrl: string): PublicWorkbookLicenseEvidence | null {
  const license = asRecordOrNull(metadata['license'])
  const id = license ? readString(license, 'id') : null
  const title = license ? (readString(license, 'title') ?? readString(license, 'name') ?? id) : null
  if (!title) {
    return null
  }
  return {
    spdxId: id ? id.toUpperCase() : null,
    title,
    evidenceUrl: sourceUrl,
  }
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
