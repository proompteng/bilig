import {
  asRecord,
  asRecordOrNull,
  hasUsableLicenseEvidence,
  isSpreadsheetFileName,
  isSpreadsheetUrl,
  readArray,
  readString,
  validatePublicWorkbookManifest,
} from './public-workbook-corpus-json.ts'
import { defaultDownloadTimeoutMs } from './public-workbook-corpus-fetch.ts'
import { fetchJsonWithTimeout } from './public-workbook-corpus-http.ts'
import {
  financialWorkbookTopicEvidence,
  recentWorkbookDateEvidence,
  type PublicWorkbookRequiredTopic,
} from './public-workbook-corpus-topics.ts'
import { sha256HexSync } from './public-workbook-corpus-workbook.ts'
import type {
  CkanPageRequest,
  CkanPageResult,
  DiscoverCkanArgs,
  PublicWorkbookLicenseEvidence,
  PublicWorkbookManifest,
  PublicWorkbookSource,
} from './public-workbook-corpus-types.ts'

export const defaultCkanPortalBases = [
  'https://data.gov.au/data/api/3/action',
  'https://open.canada.ca/data/api/3/action',
  'https://data.gov.uk/api/3/action',
  'https://data.gov.ie/api/3/action',
  'https://data.qld.gov.au/api/3/action',
  'https://data.nsw.gov.au/data/api/3/action',
  'https://data.sa.gov.au/data/api/3/action',
  'https://data.nt.gov.au/api/3/action',
]

export const defaultRecentComplexCkanPortalBases = [
  ...defaultCkanPortalBases,
  'https://data.humdata.org/api/3/action',
  'https://data.ontario.ca/api/3/action',
  'https://open.alberta.ca/api/3/action',
  'https://catalogue.data.gov.bc.ca/api/3/action',
] as const

export async function discoverCkanWorkbookSources(args: DiscoverCkanArgs): Promise<PublicWorkbookManifest> {
  validatePublicWorkbookManifest(args.manifest)
  const discoveredAt = args.discoveredAt ?? new Date().toISOString()
  const existingKeys = new Set(args.manifest.sources.map((source) => `${source.downloadUrl}|${source.license.evidenceUrl ?? ''}`))
  const existingSourceIds = new Set(args.manifest.sources.map((source) => source.id))
  const sources: PublicWorkbookSource[] = [...args.manifest.sources]
  const remainingSourceSlots = Math.max(0, args.limit - sources.length)
  if (remainingSourceSlots === 0) {
    return {
      ...args.manifest,
      generatedAt: discoveredAt,
      sources,
    }
  }
  const pagesPerPortal = Math.max(1, Math.ceil(remainingSourceSlots / args.rowsPerRequest) * 3)
  const pageRequests = args.portalBases.flatMap((portalBase) =>
    Array.from({ length: pagesPerPortal }, (_, pageIndex) => ({
      portalBase,
      query: args.query,
      rowsPerRequest: args.rowsPerRequest,
      start: pageIndex * args.rowsPerRequest,
    })),
  )
  const pages = await mapWithConcurrency(pageRequests, 8, fetchCkanPackagePage)
  for (const page of pages) {
    for (const dataset of page.packages) {
      const license = readCkanLicense(dataset)
      if (!hasUsableLicenseEvidence(license)) {
        continue
      }
      const resources = readArray(dataset, 'resources')
      for (const resource of resources) {
        const resourceRecord = asRecordOrNull(resource)
        if (!resourceRecord) {
          continue
        }
        const rawDownloadUrl = readString(resourceRecord, 'url')
        const downloadUrl = rawDownloadUrl ? resolveCkanResourceUrl(page.portalBase, rawDownloadUrl) : null
        const fileName = readResourceFileName(resourceRecord, downloadUrl)
        if (!downloadUrl || !fileName || (!isSpreadsheetUrl(downloadUrl) && !isSpreadsheetFileName(fileName))) {
          continue
        }
        const sourceUrl = readCkanDatasetUrl(page.portalBase, dataset)
        const topicEvidence = workbookTopicEvidence(args.requiredTopic, {
          dataset,
          resource: resourceRecord,
          sourceUrl,
          downloadUrl,
          fileName,
        })
        if (args.requiredTopic && topicEvidence.length === 0) {
          continue
        }
        const key = `${downloadUrl}|${license.evidenceUrl ?? ''}`
        if (existingKeys.has(key)) {
          continue
        }
        existingKeys.add(key)
        const sourceId = `ckan-${stableId(
          `${page.portalBase}:${readString(dataset, 'id') ?? ''}:${readString(resourceRecord, 'id') ?? ''}:${downloadUrl}`,
        )}`
        if (existingSourceIds.has(sourceId)) {
          continue
        }
        existingSourceIds.add(sourceId)
        sources.push({
          id: sourceId,
          kind: 'ckan-resource',
          portal: page.portalBase,
          datasetId: readString(dataset, 'id') ?? undefined,
          resourceId: readString(resourceRecord, 'id') ?? undefined,
          sourceUrl,
          downloadUrl,
          fileName,
          discoveredAt,
          license,
          ...(topicEvidence.length > 0 ? { topicEvidence } : {}),
        })
        if (sources.length >= args.limit) {
          break
        }
      }
      if (sources.length >= args.limit) {
        break
      }
    }
    if (sources.length >= args.limit) {
      break
    }
  }
  return {
    ...args.manifest,
    generatedAt: discoveredAt,
    sources,
  }
}

function workbookTopicEvidence(
  requiredTopic: PublicWorkbookRequiredTopic | undefined,
  candidate: {
    readonly dataset: Record<string, unknown>
    readonly resource: Record<string, unknown>
    readonly sourceUrl: string
    readonly downloadUrl: string
    readonly fileName: string
  },
): string[] {
  if (requiredTopic === 'financial-workpapers') {
    return financialWorkbookTopicEvidence(candidate)
  }
  if (requiredTopic === 'recent-2025-2026-workbooks') {
    return recentWorkbookDateEvidence(candidate)
  }
  return []
}

export async function discoverFinancialCkanQueries(args: {
  readonly manifest: PublicWorkbookManifest
  readonly portalBases: readonly string[]
  readonly queries: readonly string[]
  readonly limit: number
  readonly rowsPerRequest: number
  readonly onQueryDiscovered?: (manifest: PublicWorkbookManifest) => void
}): Promise<PublicWorkbookManifest> {
  const [query, ...remainingQueries] = args.queries
  if (!query || args.manifest.sources.length >= args.limit) {
    return args.manifest
  }
  const discoveredManifest = await discoverCkanWorkbookSources({
    manifest: args.manifest,
    portalBases: args.portalBases,
    query,
    limit: args.limit,
    rowsPerRequest: args.rowsPerRequest,
    requiredTopic: 'financial-workpapers',
  })
  console.log(`Discovered ${String(discoveredManifest.sources.length)} financial workbook sources after query "${query}"`)
  args.onQueryDiscovered?.(discoveredManifest)
  return discoverFinancialCkanQueries({
    ...args,
    manifest: discoveredManifest,
    queries: remainingQueries,
  })
}

export async function discoverRecentComplexCkanQueries(args: {
  readonly manifest: PublicWorkbookManifest
  readonly portalBases: readonly string[]
  readonly queries: readonly string[]
  readonly limit: number
  readonly rowsPerRequest: number
  readonly onQueryDiscovered?: (manifest: PublicWorkbookManifest) => void
}): Promise<PublicWorkbookManifest> {
  const [query, ...remainingQueries] = args.queries
  if (!query || args.manifest.sources.length >= args.limit) {
    return args.manifest
  }
  const discoveredManifest = await discoverCkanWorkbookSources({
    manifest: args.manifest,
    portalBases: args.portalBases,
    query,
    limit: args.limit,
    rowsPerRequest: args.rowsPerRequest,
    requiredTopic: 'recent-2025-2026-workbooks',
  })
  console.log(`Discovered ${String(discoveredManifest.sources.length)} recent workbook sources after query "${query}"`)
  args.onQueryDiscovered?.(discoveredManifest)
  return discoverRecentComplexCkanQueries({
    ...args,
    manifest: discoveredManifest,
    queries: remainingQueries,
  })
}

async function fetchCkanPackagePage(request: CkanPageRequest): Promise<CkanPageResult> {
  try {
    const url = new URL(`${request.portalBase.replace(/\/$/u, '')}/package_search`)
    url.searchParams.set('q', request.query)
    url.searchParams.set('rows', String(request.rowsPerRequest))
    url.searchParams.set('start', String(request.start))
    const response = await fetchJson(url)
    return {
      portalBase: request.portalBase,
      packages: readCkanPackages(response),
    }
  } catch {
    return {
      portalBase: request.portalBase,
      packages: [],
    }
  }
}

async function fetchJson(url: URL): Promise<unknown> {
  return fetchJsonWithTimeout(
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
      maxBytesLabel: 'CKAN response',
      validateResponse: (response) => {
        if (!response.ok) {
          throw new Error(`Unable to fetch ${url.href}: HTTP ${String(response.status)}`)
        }
      },
    },
  )
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

function readCkanPackages(value: unknown): Record<string, unknown>[] {
  const root = asRecord(value)
  const result = asRecord(root['result'])
  const results = readArray(result, 'results').length > 0 ? readArray(result, 'results') : readArray(result, 'result')
  return results.flatMap((entry) => {
    const record = asRecordOrNull(entry)
    return record ? [record] : []
  })
}

function readCkanLicense(dataset: Record<string, unknown>): PublicWorkbookLicenseEvidence {
  const licenseId = readString(dataset, 'license_id')
  const licenseTitle = readString(dataset, 'license_title') ?? licenseId ?? ''
  return {
    spdxId: licenseId ? normalizeSpdxLikeLicense(licenseId) : null,
    title: licenseTitle,
    evidenceUrl: readString(dataset, 'license_url') ?? readString(dataset, 'url') ?? null,
  }
}

function readCkanDatasetUrl(portalBase: string, dataset: Record<string, unknown>): string {
  const explicitUrl = readString(dataset, 'url')
  if (explicitUrl) {
    return explicitUrl
  }
  const name = readString(dataset, 'name') ?? readString(dataset, 'id') ?? 'dataset'
  return `${portalBase.replace(/\/api\/3\/action\/?$/u, '').replace(/\/$/u, '')}/dataset/${encodeURIComponent(name)}`
}

function resolveCkanResourceUrl(portalBase: string, downloadUrl: string): string | null {
  const portalWebBase = portalBase.replace(/\/api\/3\/action\/?$/u, '/')
  try {
    return new URL(downloadUrl, portalWebBase).href
  } catch {
    return null
  }
}

function readResourceFileName(resource: Record<string, unknown>, downloadUrl: string | null): string | null {
  const name = readString(resource, 'name') ?? readString(resource, 'filename')
  if (name && isSpreadsheetFileName(name)) {
    return name
  }
  if (!downloadUrl) {
    return null
  }
  const parsedUrl = parseUrlOrNull(downloadUrl)
  if (!parsedUrl) {
    return null
  }
  const pathName = parsedUrl.pathname.split('/').at(-1)
  return pathName && pathName.trim().length > 0 ? decodeUriComponentOrNull(pathName) : null
}

function parseUrlOrNull(value: string): URL | null {
  try {
    return new URL(value)
  } catch {
    return null
  }
}

function decodeUriComponentOrNull(value: string): string | null {
  try {
    return decodeURIComponent(value)
  } catch {
    return null
  }
}

function normalizeSpdxLikeLicense(value: string): string {
  return value.trim().toUpperCase()
}

function stableId(value: string): string {
  return sha256HexSync(Buffer.from(value)).slice(0, 16)
}
