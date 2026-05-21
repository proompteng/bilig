#!/usr/bin/env bun

import { spawn } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

import type { ExternalXlsxStressWorkerSummary } from './external-xlsx-memory-stress-worker.ts'
import { readNumberArg, readStringArg } from './public-workbook-corpus-cli.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const workerScriptPath = fileURLToPath(new URL('./external-xlsx-memory-stress-worker.ts', import.meta.url))
const mib = 1024 * 1024
const defaultCacheDir = join(rootDir, '.cache', 'external-xlsx-stress')
const defaultMaxRssBytes = 512 * mib
const defaultFetchTimeoutMs = 180_000
const defaultWorkerTimeoutMs = 180_000
const defaultMaxDownloadBytes = 350 * mib
const rssCheckIntervalMs = 10

export interface ExternalXlsxStressWorkbook {
  readonly id: string
  readonly fileName: string
  readonly expectedMinBytes: number
  readonly sourcePageUrl: string
  readonly downloadUrl: string
  readonly licenseTitle: string
  readonly archiveEntryPath?: string
}

export interface ExternalXlsxStressSource {
  readonly id: string
  readonly sourcePageUrl: string
  readonly downloadUrl: string
  readonly fileName: string
  readonly licenseTitle: string
  readonly workbooks: readonly Omit<ExternalXlsxStressWorkbook, 'sourcePageUrl' | 'downloadUrl' | 'licenseTitle'>[]
}

export interface ExternalXlsxStressPlan {
  readonly schemaVersion: 1
  readonly mode: 'external-xlsx-memory-stress-plan'
  readonly cacheDir: string
  readonly maxRssBytes: number
  readonly sourceCount: number
  readonly workbookCount: number
  readonly giantWorkbookCount: number
  readonly sources: readonly {
    readonly id: string
    readonly sourcePageUrl: string
    readonly downloadUrl: string
    readonly fileName: string
    readonly workbookCount: number
  }[]
  readonly workbooks: readonly ExternalXlsxStressWorkbook[]
  readonly commands: {
    readonly plan: string
    readonly run: string
  }
}

export interface ExternalXlsxStressRunSummary {
  readonly schemaVersion: 1
  readonly mode: 'external-xlsx-memory-stress-run'
  readonly cacheDir: string
  readonly maxRssBytes: number
  readonly results: readonly ExternalXlsxStressResult[]
}

export interface ExternalXlsxStressResult {
  readonly id: string
  readonly fileName: string
  readonly filePath: string
  readonly byteSize: number
  readonly sha256: string
  readonly peakRssBytes: number | null
  readonly maxRssBytes: number
  readonly status: 'passed' | 'failed'
  readonly reason?: string
  readonly importMode?: ExternalXlsxStressWorkerSummary['importMode']
  readonly sheets?: number
  readonly cells?: number
  readonly formulas?: number
  readonly warnings?: number
  readonly workbookMetadataKeys?: readonly string[]
  readonly sheetMetadataKeys?: readonly string[]
}

interface ResolvedWorkbook {
  readonly fixture: ExternalXlsxStressWorkbook
  readonly path: string
  readonly byteSize: number
  readonly sha256: string
}

const powerBiSamplesRepositoryUrl = 'https://github.com/microsoft/powerbi-desktop-samples'
const powerBiSamplesRawBaseUrl = 'https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main'

function powerBiSampleXlsxSource(args: {
  readonly id: string
  readonly workbookId: string
  readonly path: string
  readonly expectedMinBytes: number
}): ExternalXlsxStressSource {
  const encodedPath = args.path.split('/').map(encodeURIComponent).join('/')
  const fileName = basename(args.path)
  return {
    id: args.id,
    sourcePageUrl: `${powerBiSamplesRepositoryUrl}/blob/main/${encodedPath}`,
    downloadUrl: `${powerBiSamplesRawBaseUrl}/${encodedPath}`,
    fileName,
    licenseTitle: 'MIT',
    workbooks: [
      {
        id: args.workbookId,
        fileName,
        expectedMinBytes: args.expectedMinBytes,
      },
    ],
  }
}

export const externalXlsxStressSources: readonly ExternalXlsxStressSource[] = [
  {
    id: 'microsoft-powerpivot-excel-2013',
    sourcePageUrl: 'https://www.microsoft.com/en-us/download/details.aspx?id=102',
    downloadUrl: 'https://download.microsoft.com/download/0/2/c/02c5d169-11fe-4d7a-9ade-ebdd469e249b/PowerPivotExamplesExcel2013.zip',
    fileName: 'PowerPivotExamplesExcel2013.zip',
    licenseTitle: 'Microsoft Download Center sample terms',
    workbooks: [
      {
        id: 'powerpivot-tutorial-sample',
        fileName: 'PowerPivotTutorialSample.xlsx',
        archiveEntryPath: 'PowerPivotTutorialSample.xlsx',
        expectedMinBytes: 100 * mib,
      },
      {
        id: 'powerpivot-healthcare-audit',
        fileName: 'PowerPivot Healthcare Audit.xlsx',
        archiveEntryPath: 'PowerPivot Healthcare Audit.xlsx',
        expectedMinBytes: 4 * mib,
      },
      {
        id: 'powerpivot-financial-report-usage',
        fileName: 'LCA BI - Financial Report Usage.xlsx',
        archiveEntryPath: 'LCA BI - Financial Report Usage.xlsx',
        expectedMinBytes: 512 * 1024,
      },
    ],
  },
  {
    id: 'microsoft-contoso-dax-formulas',
    sourcePageUrl: 'https://www.microsoft.com/en-nz/download/details.aspx?id=28572',
    downloadUrl: 'https://download.microsoft.com/download/1/3/0/130544af-21f2-44fd-9c05-158b8316c2d0/Contoso%20DAX%20Formula%20Samples.zip',
    fileName: 'Contoso DAX Formula Samples.zip',
    licenseTitle: 'Microsoft Download Center sample terms',
    workbooks: [
      {
        id: 'contoso-sample-dax-formulas',
        fileName: 'Contoso Sample DAX Formulas.xlsx',
        archiveEntryPath: 'Contoso Sample DAX Formulas.xlsx',
        expectedMinBytes: 200 * mib,
      },
    ],
  },
  {
    id: 'microsoft-contoso-pnl-powerpivot',
    sourcePageUrl: 'https://www.microsoft.com/en-us/download/details.aspx?id=38838',
    downloadUrl: 'https://download.microsoft.com/download/b/e/c/becf5873-6b88-4920-9096-2c10ba98de60/ContosoPnL_Excel2013.zip',
    fileName: 'ContosoPnL_Excel2013.zip',
    licenseTitle: 'Microsoft Download Center sample terms',
    workbooks: [
      {
        id: 'contoso-pnl-excel-2013',
        fileName: 'ContosoPnL_Excel2013.xlsx',
        archiveEntryPath: 'ContosoPnL_Excel2013/ContosoPnL_Excel2013.xlsx',
        expectedMinBytes: 5 * mib,
      },
    ],
  },
  powerBiSampleXlsxSource({
    id: 'powerbi-adventureworks-sales-xlsx',
    workbookId: 'powerbi-adventureworks-sales',
    path: 'AdventureWorks Sales Sample/AdventureWorks Sales.xlsx',
    expectedMinBytes: 13 * mib,
  }),
  powerBiSampleXlsxSource({
    id: 'powerbi-customer-feedback-xlsx',
    workbookId: 'powerbi-customer-feedback',
    path: 'Monthly Desktop Blog Samples/2019/customerfeedback.xlsx',
    expectedMinBytes: 7 * mib,
  }),
  powerBiSampleXlsxSource({
    id: 'powerbi-customer-profitability-xlsx',
    workbookId: 'powerbi-customer-profitability',
    path: 'powerbi-service-samples/Customer Profitability Sample-no-PV.xlsx',
    expectedMinBytes: 2 * mib,
  }),
  powerBiSampleXlsxSource({
    id: 'powerbi-human-resources-xlsx',
    workbookId: 'powerbi-human-resources',
    path: 'powerbi-service-samples/Human Resources Sample-no-PV.xlsx',
    expectedMinBytes: 9 * mib,
  }),
  powerBiSampleXlsxSource({
    id: 'powerbi-it-spend-analysis-xlsx',
    workbookId: 'powerbi-it-spend-analysis',
    path: 'powerbi-service-samples/IT Spend Analysis Sample-no-PV.xlsx',
    expectedMinBytes: 1 * mib,
  }),
  powerBiSampleXlsxSource({
    id: 'powerbi-opportunity-tracking-xlsx',
    workbookId: 'powerbi-opportunity-tracking',
    path: 'powerbi-service-samples/Opportunity Tracking Sample no PV.xlsx',
    expectedMinBytes: 640 * 1024,
  }),
  powerBiSampleXlsxSource({
    id: 'powerbi-procurement-analysis-xlsx',
    workbookId: 'powerbi-procurement-analysis',
    path: 'powerbi-service-samples/Procurement Analysis Sample-no-PV.xlsx',
    expectedMinBytes: 14 * mib,
  }),
  powerBiSampleXlsxSource({
    id: 'powerbi-retail-analysis-xlsx',
    workbookId: 'powerbi-retail-analysis',
    path: 'powerbi-service-samples/Retail Analysis Sample-no-PV.xlsx',
    expectedMinBytes: 10 * mib,
  }),
  powerBiSampleXlsxSource({
    id: 'powerbi-sales-marketing-xlsx',
    workbookId: 'powerbi-sales-marketing',
    path: 'powerbi-service-samples/Sales and Marketing Sample-no-PV.xlsx',
    expectedMinBytes: 8 * mib,
  }),
  powerBiSampleXlsxSource({
    id: 'powerbi-supplier-quality-xlsx',
    workbookId: 'powerbi-supplier-quality',
    path: 'powerbi-service-samples/Supplier Quality Analysis Sample-no-PV.xlsx',
    expectedMinBytes: 700 * 1024,
  }),
]

export function buildExternalXlsxStressPlan(args: {
  readonly cacheDir: string
  readonly maxRssBytes?: number
  readonly sources?: readonly ExternalXlsxStressSource[]
}): ExternalXlsxStressPlan {
  const sources = args.sources ?? externalXlsxStressSources
  const workbooks = sources.flatMap((source) =>
    source.workbooks.map((workbook) => ({
      ...workbook,
      sourcePageUrl: source.sourcePageUrl,
      downloadUrl: source.downloadUrl,
      licenseTitle: source.licenseTitle,
    })),
  )
  return {
    schemaVersion: 1,
    mode: 'external-xlsx-memory-stress-plan',
    cacheDir: args.cacheDir,
    maxRssBytes: args.maxRssBytes ?? defaultMaxRssBytes,
    sourceCount: sources.length,
    workbookCount: workbooks.length,
    giantWorkbookCount: workbooks.filter((workbook) => workbook.expectedMinBytes >= 100 * mib).length,
    sources: sources.map((source) => ({
      id: source.id,
      sourcePageUrl: source.sourcePageUrl,
      downloadUrl: source.downloadUrl,
      fileName: source.fileName,
      workbookCount: source.workbooks.length,
    })),
    workbooks,
    commands: {
      plan: 'pnpm external-xlsx-memory-stress:plan',
      run: 'pnpm external-xlsx-memory-stress',
    },
  }
}

export function validateExternalXlsxStressPlan(plan: ExternalXlsxStressPlan): string[] {
  const findings: string[] = []
  if (plan.schemaVersion !== 1 || plan.mode !== 'external-xlsx-memory-stress-plan') {
    findings.push('plan has an invalid schema or mode')
  }
  if (plan.sourceCount !== plan.sources.length) {
    findings.push('source count does not match sources length')
  }
  if (plan.workbookCount !== plan.workbooks.length) {
    findings.push('workbook count does not match workbooks length')
  }
  if (plan.giantWorkbookCount < 2) {
    findings.push('plan must include at least two 100 MiB+ workbook stress targets')
  }
  if (
    !plan.sources.some((source) => source.sourcePageUrl.includes('microsoft.com') && source.downloadUrl.includes('download.microsoft.com'))
  ) {
    findings.push('plan must include Microsoft Download Center PowerPivot or DAX workbook sources')
  }
  for (const workbook of plan.workbooks) {
    if (!workbook.id || !workbook.fileName || !workbook.downloadUrl || !workbook.sourcePageUrl) {
      findings.push(`workbook is missing required source fields: ${workbook.id || workbook.fileName || 'unknown'}`)
    }
    if (workbook.expectedMinBytes <= 0) {
      findings.push(`workbook expected minimum size must be positive: ${workbook.id}`)
    }
  }
  return findings
}

async function runExternalXlsxStress(args: {
  readonly cacheDir: string
  readonly maxRssBytes: number
  readonly fetchTimeoutMs: number
  readonly workerTimeoutMs: number
  readonly maxDownloadBytes: number
  readonly limit?: number
}): Promise<ExternalXlsxStressRunSummary> {
  const selectedSources = limitSources(externalXlsxStressSources, args.limit)
  const resolvedWorkbooks: ResolvedWorkbook[] = []
  for (const source of selectedSources) {
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Sequential downloads keep local memory and network pressure bounded.
    const sourceWorkbooks = await ensureExternalXlsxStressSource(source, {
      cacheDir: args.cacheDir,
      fetchTimeoutMs: args.fetchTimeoutMs,
      maxDownloadBytes: args.maxDownloadBytes,
    })
    resolvedWorkbooks.push(...sourceWorkbooks)
  }

  const results: ExternalXlsxStressResult[] = []
  for (const workbook of resolvedWorkbooks) {
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Sequential import workers isolate peak RSS per workbook.
    results.push(await runStressWorker(workbook, args.maxRssBytes, args.workerTimeoutMs))
  }
  return {
    schemaVersion: 1,
    mode: 'external-xlsx-memory-stress-run',
    cacheDir: args.cacheDir,
    maxRssBytes: args.maxRssBytes,
    results,
  }
}

function limitSources(sources: readonly ExternalXlsxStressSource[], limit: number | undefined): readonly ExternalXlsxStressSource[] {
  if (limit === undefined || limit <= 0) {
    return sources
  }
  const selected: ExternalXlsxStressSource[] = []
  let remaining = Math.trunc(limit)
  for (const source of sources) {
    if (remaining <= 0) {
      break
    }
    const workbooks = source.workbooks.slice(0, remaining)
    if (workbooks.length > 0) {
      selected.push({ ...source, workbooks })
      remaining -= workbooks.length
    }
  }
  return selected
}

async function ensureExternalXlsxStressSource(
  source: ExternalXlsxStressSource,
  args: {
    readonly cacheDir: string
    readonly fetchTimeoutMs: number
    readonly maxDownloadBytes: number
  },
): Promise<ResolvedWorkbook[]> {
  mkdirSync(args.cacheDir, { recursive: true })
  const sourceCachePath = join(args.cacheDir, source.fileName)
  const sourceBytes = await readOrFetchSourceBytes(source, sourceCachePath, args)
  if (source.fileName.toLowerCase().endsWith('.zip')) {
    return await extractWorkbookEntries(source, sourceBytes, args.cacheDir)
  }
  const workbook = source.workbooks[0]
  if (!workbook) {
    return []
  }
  return [
    assertResolvedWorkbook({
      fixture: {
        ...workbook,
        sourcePageUrl: source.sourcePageUrl,
        downloadUrl: source.downloadUrl,
        licenseTitle: source.licenseTitle,
      },
      path: sourceCachePath,
    }),
  ]
}

async function readOrFetchSourceBytes(
  source: ExternalXlsxStressSource,
  sourceCachePath: string,
  args: {
    readonly fetchTimeoutMs: number
    readonly maxDownloadBytes: number
  },
): Promise<Uint8Array> {
  if (existsSync(sourceCachePath)) {
    return readFileSync(sourceCachePath)
  }
  const { fetchBodyBytesWithTimeout } = await import('./public-workbook-corpus-http.ts')
  const { bytes } = await fetchBodyBytesWithTimeout(
    source.downloadUrl,
    {},
    {
      timeoutMs: args.fetchTimeoutMs,
      maxBytes: args.maxDownloadBytes,
      maxBytesLabel: source.fileName,
      validateResponse: (response) => {
        if (!response.ok) {
          throw new Error(`Failed to fetch ${source.fileName}: HTTP ${String(response.status)}`)
        }
      },
    },
  )
  mkdirSync(dirname(sourceCachePath), { recursive: true })
  writeFileSync(sourceCachePath, bytes)
  return bytes
}

async function extractWorkbookEntries(
  source: ExternalXlsxStressSource,
  sourceBytes: Uint8Array,
  cacheDir: string,
): Promise<ResolvedWorkbook[]> {
  const { unzipSync } = await import('fflate')
  const unzipped = unzipSync(sourceBytes)
  const sourceDir = join(cacheDir, source.id)
  mkdirSync(sourceDir, { recursive: true })
  return source.workbooks.map((workbook) => {
    const entryPath = workbook.archiveEntryPath ?? workbook.fileName
    const bytes = unzipped[entryPath]
    if (!bytes) {
      throw new Error(`Archive ${source.fileName} is missing workbook entry ${entryPath}`)
    }
    const outputPath = join(sourceDir, workbook.fileName)
    if (!existsSync(outputPath) || readFileSync(outputPath).byteLength !== bytes.byteLength) {
      writeFileSync(outputPath, bytes)
    }
    return assertResolvedWorkbook({
      fixture: {
        ...workbook,
        sourcePageUrl: source.sourcePageUrl,
        downloadUrl: source.downloadUrl,
        licenseTitle: source.licenseTitle,
      },
      path: outputPath,
    })
  })
}

function assertResolvedWorkbook(input: { readonly fixture: ExternalXlsxStressWorkbook; readonly path: string }): ResolvedWorkbook {
  const bytes = readFileSync(input.path)
  if (bytes.byteLength < input.fixture.expectedMinBytes) {
    throw new Error(
      `${input.fixture.fileName} is smaller than expected: ${formatByteSize(bytes.byteLength)} < ${formatByteSize(
        input.fixture.expectedMinBytes,
      )}`,
    )
  }
  return {
    fixture: input.fixture,
    path: input.path,
    byteSize: bytes.byteLength,
    sha256: createHash('sha256').update(bytes).digest('hex'),
  }
}

async function runStressWorker(workbook: ResolvedWorkbook, maxRssBytes: number, timeoutMs: number): Promise<ExternalXlsxStressResult> {
  const { startChildRssWatchdog, terminateChildProcess } = await import('./public-workbook-corpus-process.ts')
  return new Promise((resolvePromise) => {
    const child = spawn(process.execPath, [workerScriptPath, '--file', workbook.path, '--file-name', workbook.fixture.fileName], {
      detached: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let peakRssBytes = 0
    let settled = false
    const finish = (result: ExternalXlsxStressResult): void => {
      if (settled) {
        return
      }
      settled = true
      clearTimeout(timer)
      stopRssWatchdog()
      // oxlint-disable-next-line eslint-plugin-promise(no-multiple-resolved) -- `settled` gates close/error/watchdog races before resolving.
      resolvePromise(result)
    }
    const fail = (reason: string): void => {
      finish({
        id: workbook.fixture.id,
        fileName: workbook.fixture.fileName,
        filePath: workbook.path,
        byteSize: workbook.byteSize,
        sha256: workbook.sha256,
        peakRssBytes: peakRssBytes || null,
        maxRssBytes,
        status: 'failed',
        reason,
      })
    }
    const stopRssWatchdog = startChildRssWatchdog(child, {
      maxRssBytes,
      intervalMs: rssCheckIntervalMs,
      onSample: (rssBytes) => {
        peakRssBytes = Math.max(peakRssBytes, rssBytes)
      },
      onLimitExceeded: (rssBytes) => {
        peakRssBytes = Math.max(peakRssBytes, rssBytes)
        terminateChildProcess(child, 'SIGTERM', { processGroup: true })
        fail(`peak RSS ${formatByteSize(rssBytes)} exceeded ${formatByteSize(maxRssBytes)}`)
      },
    })
    const timer = setTimeout(() => {
      terminateChildProcess(child, 'SIGTERM', { processGroup: true })
      fail(`worker timed out after ${String(timeoutMs)}ms`)
    }, timeoutMs)
    timer.unref()
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.stderr.on('data', (chunk: string) => {
      stderr += chunk
    })
    child.on('error', (error) => fail(`worker failed to start: ${error.message}`))
    child.on('close', (code, signal) => {
      if (settled) {
        return
      }
      if (code !== 0) {
        fail(`worker exited with ${signal ? `signal ${signal}` : `code ${String(code)}`}: ${compactWorkerOutput(stderr || stdout)}`)
        return
      }
      try {
        const parsedWorkerSummary = parseWorkerSummaryJson(stdout)
        finish({
          id: workbook.fixture.id,
          fileName: workbook.fixture.fileName,
          filePath: workbook.path,
          byteSize: workbook.byteSize,
          sha256: workbook.sha256,
          peakRssBytes: peakRssBytes || null,
          maxRssBytes,
          status: 'passed',
          importMode: parsedWorkerSummary.importMode,
          sheets: parsedWorkerSummary.sheets,
          cells: parsedWorkerSummary.cells,
          formulas: parsedWorkerSummary.formulas,
          warnings: parsedWorkerSummary.warnings,
          workbookMetadataKeys: parsedWorkerSummary.workbookMetadataKeys,
          sheetMetadataKeys: parsedWorkerSummary.sheetMetadataKeys,
        })
      } catch (error) {
        fail(`worker returned invalid JSON: ${error instanceof Error ? error.message : String(error)}`)
      }
    })
  })
}

function parseWorkerSummaryJson(stdout: string): ExternalXlsxStressWorkerSummary {
  const value: unknown = JSON.parse(stdout)
  if (!isRecord(value)) {
    throw new Error('Expected worker summary object')
  }
  return {
    importMode: readWorkerImportMode(value),
    sheets: readWorkerNumber(value, 'sheets'),
    cells: readWorkerNumber(value, 'cells'),
    formulas: readWorkerNumber(value, 'formulas'),
    warnings: readWorkerNumber(value, 'warnings'),
    workbookMetadataKeys: readWorkerStringArray(value, 'workbookMetadataKeys'),
    sheetMetadataKeys: readWorkerStringArray(value, 'sheetMetadataKeys'),
  }
}

function readWorkerImportMode(record: Readonly<Record<string, unknown>>): ExternalXlsxStressWorkerSummary['importMode'] {
  const value = record['importMode']
  if (value === 'headless-inspect' || value === 'public-snapshot') {
    return value
  }
  throw new Error('Expected worker summary import mode')
}

function readWorkerNumber(record: Readonly<Record<string, unknown>>, key: string): number {
  const value = record[key]
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`Expected worker summary numeric field ${key}`)
  }
  return value
}

function readWorkerStringArray(record: Readonly<Record<string, unknown>>, key: string): readonly string[] {
  const value = record[key]
  if (!Array.isArray(value) || !value.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected worker summary string array field ${key}`)
  }
  return value
}

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function compactWorkerOutput(value: string): string {
  return value.replace(/\s+/gu, ' ').trim().slice(0, 1_000)
}

function formatByteSize(bytes: number): string {
  const gib = bytes / 1024 / 1024 / 1024
  if (gib >= 1) {
    return `${gib.toFixed(2)} GiB`
  }
  return `${(bytes / mib).toFixed(1)} MiB`
}

async function main(): Promise<void> {
  const command = process.argv[2] ?? 'plan'
  if (command === 'worker') {
    const { runExternalXlsxStressWorker } = await import('./external-xlsx-memory-stress-worker.ts')
    runExternalXlsxStressWorker()
    return
  }
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const maxRssBytes = Math.max(1, Math.trunc(readNumberArg('--max-rss-mb', defaultMaxRssBytes / mib))) * mib
  if (command === 'plan' || command === 'check') {
    const plan = buildExternalXlsxStressPlan({ cacheDir, maxRssBytes })
    process.stdout.write(`${JSON.stringify(plan, null, 2)}\n`)
    const findings = validateExternalXlsxStressPlan(plan)
    if (command === 'check' && findings.length > 0) {
      process.stderr.write(`External XLSX memory stress plan failed: ${findings.join('; ')}\n`)
      process.exitCode = 1
    }
    return
  }
  if (command === 'run') {
    const summary = await runExternalXlsxStress({
      cacheDir,
      maxRssBytes,
      fetchTimeoutMs: readNumberArg('--fetch-timeout-ms', defaultFetchTimeoutMs),
      workerTimeoutMs: readNumberArg('--worker-timeout-ms', defaultWorkerTimeoutMs),
      maxDownloadBytes: readNumberArg('--max-download-bytes', defaultMaxDownloadBytes),
      limit: readOptionalPositiveIntegerArg('--limit'),
    })
    process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`)
    if (summary.results.some((result) => result.status === 'failed')) {
      process.exitCode = 1
    }
    return
  }
  throw new Error(`Unknown external XLSX memory stress command: ${command}`)
}

function readOptionalPositiveIntegerArg(name: string): number | undefined {
  const raw = readStringArg(name, '')
  if (!raw) {
    return undefined
  }
  const parsed = Number(raw)
  if (!/^\d+$/u.test(raw) || parsed <= 0 || !Number.isSafeInteger(parsed)) {
    throw new Error(`Expected ${name} to be a positive integer`)
  }
  return parsed
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    await main()
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`)
    process.exit(1)
  }
}
