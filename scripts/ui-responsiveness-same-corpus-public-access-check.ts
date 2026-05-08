#!/usr/bin/env bun

import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { buildWorkbookBenchmarkCorpus, isWorkbookBenchmarkCorpusId } from '../packages/benchmarks/src/workbook-corpus.js'
import type { WorkbookBenchmarkCorpusId } from '../packages/benchmarks/src/workbook-corpus.js'
import type {
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureVerifiedCell,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { formatJsonForRepo } from './scorecard-format.ts'
import { verifyXlsxCorpusFingerprint } from './ui-responsiveness-same-corpus-fingerprint.ts'

type SameCorpusPublicAccessProduct = Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>

export interface SameCorpusPublicAccessCheck {
  readonly schemaVersion: 1
  readonly suite: 'ui-responsiveness-same-corpus-public-access-check'
  readonly generatedAt: string
  readonly corpusCaseId: WorkbookBenchmarkCorpusId
  readonly materializedCells: number
  readonly requestedProductCount: number
  readonly verifiedProductCount: number
  readonly allRequestedProductsVerified: boolean
  readonly products: readonly SameCorpusPublicAccessProductResult[]
  readonly limitations: readonly string[]
}

export interface SameCorpusPublicAccessProductResult {
  readonly product: SameCorpusPublicAccessProduct
  readonly source: string
  readonly resolvedXlsxUrl: string
  readonly byteSize: number
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: readonly string[]
}

export type SameCorpusPublicAccessFetch = (url: string, label: string) => Promise<Uint8Array>

interface SameCorpusPublicAccessArgs {
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly generatedAt: string
  readonly googleSheetsUrl: string | null
  readonly microsoftExcelWebUrl: string | null
  readonly outputPath: string | null
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCorpusId: WorkbookBenchmarkCorpusId = 'wide-mixed-250k'

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2))
  const check = await buildSameCorpusPublicAccessCheck({
    corpusId: args.corpusId,
    fetchXlsxBytes: fetchSameCorpusXlsxBytes,
    generatedAt: args.generatedAt,
    googleSheetsUrl: args.googleSheetsUrl,
    microsoftExcelWebUrl: args.microsoftExcelWebUrl,
  })
  validateSameCorpusPublicAccessCheck(check)
  const serializedJson = `${JSON.stringify(check, null, 2)}\n`
  if (args.outputPath) {
    mkdirSync(dirname(args.outputPath), { recursive: true })
    writeFileSync(
      args.outputPath,
      formatJsonForRepo({
        rootDir,
        serializedJson,
        tempPrefix: 'ui-responsiveness-same-corpus-public-access',
      }),
    )
  }
  process.stdout.write(serializedJson)
}

export async function buildSameCorpusPublicAccessCheck(args: {
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly fetchXlsxBytes: SameCorpusPublicAccessFetch
  readonly generatedAt: string
  readonly googleSheetsUrl: string | null
  readonly microsoftExcelWebUrl: string | null
}): Promise<SameCorpusPublicAccessCheck> {
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const productSpecs = [
    ...(args.googleSheetsUrl
      ? [
          {
            product: 'google-sheets' as const,
            source: args.googleSheetsUrl,
            resolvedXlsxUrl: googleSheetsXlsxExportUrl(args.googleSheetsUrl),
          },
        ]
      : []),
    ...(args.microsoftExcelWebUrl
      ? [
          {
            product: 'microsoft-excel-web' as const,
            source: args.microsoftExcelWebUrl,
            resolvedXlsxUrl: microsoftExcelWebSourceXlsxUrl(args.microsoftExcelWebUrl),
          },
        ]
      : []),
  ]
  if (productSpecs.length === 0) {
    throw new Error('Same-corpus public access check requires --google-sheets-url, --microsoft-excel-web-url, or both.')
  }
  const products = await Promise.all(
    productSpecs.map(async (spec) => {
      const bytes = await args.fetchXlsxBytes(spec.resolvedXlsxUrl, `${spec.product} same-corpus XLSX`)
      return {
        product: spec.product,
        source: spec.source,
        resolvedXlsxUrl: spec.resolvedXlsxUrl,
        byteSize: bytes.length,
        corpusVerification: verifyXlsxCorpusFingerprint(bytes, corpus, verificationMethodForProduct(spec.product)),
        limitations: limitationsForProduct(spec.product),
      }
    }),
  )
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-public-access-check',
    generatedAt: args.generatedAt,
    corpusCaseId: corpus.id,
    materializedCells: corpus.materializedCellCount,
    requestedProductCount: productSpecs.length,
    verifiedProductCount: products.filter((entry) => entry.corpusVerification.verified).length,
    allRequestedProductsVerified: products.every((entry) => entry.corpusVerification.verified),
    products,
    limitations: [
      'This check proves URL reachability and same-corpus workbook identity through XLSX export bytes.',
      'It is not browser timing evidence and does not satisfy the same-corpus 10x UI responsiveness requirement by itself.',
    ],
  }
}

export function validateSameCorpusPublicAccessCheck(check: SameCorpusPublicAccessCheck): void {
  if (check.schemaVersion !== 1 || check.suite !== 'ui-responsiveness-same-corpus-public-access-check') {
    throw new Error('Unexpected same-corpus public access check header')
  }
  if (check.requestedProductCount < 1 || check.products.length !== check.requestedProductCount) {
    throw new Error('Same-corpus public access check product count is stale')
  }
  const verifiedProductCount = check.products.filter((entry) => entry.corpusVerification.verified).length
  if (check.verifiedProductCount !== verifiedProductCount) {
    throw new Error('Same-corpus public access check verified count is stale')
  }
  if (check.allRequestedProductsVerified !== check.products.every((entry) => entry.corpusVerification.verified)) {
    throw new Error('Same-corpus public access check pass flag is stale')
  }
  for (const product of check.products) {
    if (product.byteSize <= 0) {
      throw new Error(`Same-corpus public access check has empty XLSX bytes for ${product.product}`)
    }
    if (product.corpusVerification.materializedCells !== check.materializedCells) {
      throw new Error(`Same-corpus public access check materialized cell count is stale for ${product.product}`)
    }
    if (product.corpusVerification.checkedCells.length < 3) {
      throw new Error(`Same-corpus public access check has too few verified cells for ${product.product}`)
    }
  }
}

export function parseSameCorpusPublicAccessCheckJson(value: unknown): SameCorpusPublicAccessCheck {
  if (!isSameCorpusPublicAccessCheck(value)) {
    throw new Error('Unexpected same-corpus public access check JSON')
  }
  validateSameCorpusPublicAccessCheck(value)
  return value
}

export async function fetchSameCorpusXlsxBytes(url: string, label: string): Promise<Uint8Array> {
  const response = await fetch(url, { redirect: 'follow' })
  if (!response.ok) {
    const bodySnippet = (await response.text().catch(() => '')).replace(/\s+/g, ' ').trim().slice(0, 300)
    throw new Error(`${label} returned HTTP ${String(response.status)}: ${bodySnippet}`)
  }
  const bytes = new Uint8Array(await response.arrayBuffer())
  if (bytes.length === 0) {
    throw new Error(`${label} returned an empty XLSX payload`)
  }
  if (looksLikeHtml(bytes)) {
    throw new Error(`${label} returned HTML instead of XLSX bytes`)
  }
  return bytes
}

export function googleSheetsXlsxExportUrl(sourceUrl: string): string {
  const spreadsheetId = /\/spreadsheets\/d\/([^/?#]+)/u.exec(sourceUrl)?.[1]
  if (!spreadsheetId) {
    throw new Error(`Unable to extract Google Sheets spreadsheet ID from URL: ${sourceUrl}`)
  }
  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/export?format=xlsx`
}

export function microsoftExcelWebSourceXlsxUrl(sourceUrl: string): string {
  const parsed = new URL(sourceUrl)
  if (!parsed.hostname.includes('view.officeapps.live.com')) {
    return sourceUrl
  }
  const source = parsed.searchParams.get('src')
  if (!source) {
    throw new Error(`Unable to extract Microsoft Excel Web source XLSX URL from viewer URL: ${sourceUrl}`)
  }
  return source
}

function verificationMethodForProduct(product: SameCorpusPublicAccessProduct): SameCorpusCaptureCorpusVerification['method'] {
  return product === 'google-sheets' ? 'google-sheets-xlsx-export' : 'microsoft-excel-web-source-xlsx'
}

function limitationsForProduct(product: SameCorpusPublicAccessProduct): readonly string[] {
  if (product === 'google-sheets') {
    return ['Google Sheets must be shared so anyone with the link can export the native sheet as XLSX.']
  }
  return ['Microsoft Excel Web must wrap a public HTTPS XLSX URL for the same emitted corpus workbook.']
}

function parseArgs(argv: readonly string[]): SameCorpusPublicAccessArgs {
  const googleSheetsUrl = argumentValue(argv, '--google-sheets-url')
  const microsoftExcelWebUrl = argumentValue(argv, '--microsoft-excel-web-url')
  if (!googleSheetsUrl && !microsoftExcelWebUrl) {
    throw new Error('Same-corpus public access check requires --google-sheets-url, --microsoft-excel-web-url, or both.')
  }
  return {
    corpusId: parseCorpusId(argumentValue(argv, '--corpus') ?? defaultCorpusId),
    generatedAt: argumentValue(argv, '--generated-at') ?? new Date().toISOString(),
    googleSheetsUrl,
    microsoftExcelWebUrl,
    outputPath: resolveOptionalPath(argumentValue(argv, '--output')),
  }
}

function parseCorpusId(value: string): WorkbookBenchmarkCorpusId {
  if (!isWorkbookBenchmarkCorpusId(value)) {
    throw new Error(`Unexpected workbook benchmark corpus id: ${value}`)
  }
  return value
}

function resolveOptionalPath(value: string | null): string | null {
  return value ? resolve(value) : null
}

function argumentValue(argv: readonly string[], name: string): string | null {
  const index = argv.indexOf(name)
  if (index === -1) {
    return null
  }
  const value = argv[index + 1]
  if (!value) {
    throw new Error(`Missing value after ${name}`)
  }
  return value
}

function isSameCorpusPublicAccessCheck(value: unknown): value is SameCorpusPublicAccessCheck {
  if (!isJsonRecord(value)) {
    return false
  }
  return (
    value.schemaVersion === 1 &&
    value.suite === 'ui-responsiveness-same-corpus-public-access-check' &&
    typeof value.generatedAt === 'string' &&
    isWorkbookBenchmarkCorpusId(value.corpusCaseId) &&
    typeof value.materializedCells === 'number' &&
    typeof value.requestedProductCount === 'number' &&
    typeof value.verifiedProductCount === 'number' &&
    typeof value.allRequestedProductsVerified === 'boolean' &&
    Array.isArray(value.products) &&
    value.products.every(isSameCorpusPublicAccessProductResult) &&
    Array.isArray(value.limitations) &&
    value.limitations.every(isString)
  )
}

function isSameCorpusPublicAccessProductResult(value: unknown): value is SameCorpusPublicAccessProductResult {
  if (!isJsonRecord(value)) {
    return false
  }
  return (
    (value.product === 'google-sheets' || value.product === 'microsoft-excel-web') &&
    typeof value.source === 'string' &&
    typeof value.resolvedXlsxUrl === 'string' &&
    typeof value.byteSize === 'number' &&
    isSameCorpusCaptureCorpusVerification(value.corpusVerification) &&
    Array.isArray(value.limitations) &&
    value.limitations.every(isString)
  )
}

function isSameCorpusCaptureCorpusVerification(value: unknown): value is SameCorpusCaptureCorpusVerification {
  if (!isJsonRecord(value)) {
    return false
  }
  return (
    typeof value.verified === 'boolean' &&
    (value.method === 'bilig-benchmark-state' ||
      value.method === 'google-sheets-xlsx-export' ||
      value.method === 'microsoft-excel-web-source-xlsx') &&
    typeof value.sheetName === 'string' &&
    typeof value.materializedCells === 'number' &&
    Array.isArray(value.checkedCells) &&
    value.checkedCells.every(isSameCorpusCaptureVerifiedCell)
  )
}

function isSameCorpusCaptureVerifiedCell(value: unknown): value is SameCorpusCaptureVerifiedCell {
  return isJsonRecord(value) && typeof value.address === 'string' && typeof value.expected === 'string' && typeof value.actual === 'string'
}

function isJsonRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isString(value: unknown): value is string {
  return typeof value === 'string'
}

function looksLikeHtml(bytes: Uint8Array): boolean {
  const prefix = new TextDecoder()
    .decode(bytes.slice(0, Math.min(bytes.length, 256)))
    .trimStart()
    .toLowerCase()
  return prefix.startsWith('<!doctype html') || prefix.startsWith('<html')
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    await main()
  } catch (error: unknown) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exitCode = 1
  }
}
