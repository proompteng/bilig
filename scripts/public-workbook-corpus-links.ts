import {
  hasUsableLicenseEvidence,
  isSpreadsheetFileName,
  isSpreadsheetUrl,
  validatePublicWorkbookManifest,
} from './public-workbook-corpus-json.ts'
import { sha256HexSync } from './public-workbook-corpus-workbook.ts'
import type { PublicWorkbookLicenseEvidence, PublicWorkbookManifest, PublicWorkbookSource } from './public-workbook-corpus-types.ts'

export interface AddPublicWorkbookLinkSourceResult {
  readonly added: boolean
  readonly manifest: PublicWorkbookManifest
  readonly source: PublicWorkbookSource
}

export function addPublicWorkbookLinkSource(args: {
  readonly manifest: PublicWorkbookManifest
  readonly sourceUrl: string
  readonly downloadUrl?: string
  readonly fileName?: string
  readonly licenseTitle: string
  readonly licenseUrl: string
  readonly licenseSpdxId?: string | null
  readonly discoveredAt?: string
  readonly topicEvidence?: readonly string[]
}): AddPublicWorkbookLinkSourceResult {
  validatePublicWorkbookManifest(args.manifest)
  const normalized = normalizePublicWorkbookLink({
    downloadUrl: args.downloadUrl,
    fileName: args.fileName,
    sourceUrl: args.sourceUrl,
  })
  const license = publicWorkbookLicenseEvidence({
    licenseSpdxId: args.licenseSpdxId,
    licenseTitle: args.licenseTitle,
    licenseUrl: args.licenseUrl,
  })
  const existing = args.manifest.sources.find(
    (source) =>
      normalizeUrlForDeduplication(source.downloadUrl) === normalizeUrlForDeduplication(normalized.downloadUrl) &&
      source.license.evidenceUrl === license.evidenceUrl,
  )
  if (existing) {
    return {
      added: false,
      manifest: args.manifest,
      source: existing,
    }
  }
  const discoveredAt = args.discoveredAt ?? new Date().toISOString()
  const source: PublicWorkbookSource = {
    id: `direct-${stableId(`${normalized.sourceUrl}:${normalized.downloadUrl}:${license.evidenceUrl ?? ''}`)}`,
    kind: 'direct-url',
    sourceUrl: normalized.sourceUrl,
    downloadUrl: normalized.downloadUrl,
    fileName: normalized.fileName,
    discoveredAt,
    license,
    ...(args.topicEvidence && args.topicEvidence.length > 0 ? { topicEvidence: args.topicEvidence } : {}),
  }
  const manifest: PublicWorkbookManifest = {
    ...args.manifest,
    generatedAt: discoveredAt,
    sources: [...args.manifest.sources, source],
  }
  validatePublicWorkbookManifest(manifest)
  return {
    added: true,
    manifest,
    source,
  }
}

export function normalizePublicWorkbookLink(args: {
  readonly sourceUrl: string
  readonly downloadUrl?: string
  readonly fileName?: string
}): { readonly sourceUrl: string; readonly downloadUrl: string; readonly fileName: string } {
  const sourceUrl = parseAbsoluteHttpUrl(args.sourceUrl, '--source-url')
  const explicitDownloadUrl = args.downloadUrl?.trim()
  const downloadUrl = explicitDownloadUrl
    ? parseAbsoluteHttpUrl(explicitDownloadUrl, '--download-url').href
    : downloadUrlForSharedWorkbookSource(sourceUrl, args.fileName)
  const fileName =
    normalizedWorkbookFileName(args.fileName) ?? workbookFileNameFromUrl(downloadUrl) ?? workbookFileNameForSharedSource(sourceUrl)
  if (!fileName) {
    throw new Error('Expected --file-name for shared workbook links that do not expose a spreadsheet file name')
  }
  if (!isSpreadsheetUrl(downloadUrl) && !isSpreadsheetFileName(fileName)) {
    throw new Error('Shared workbook link must resolve to an .xlsx, .xlsm, or .xls workbook candidate')
  }
  return {
    sourceUrl: sourceUrl.href,
    downloadUrl,
    fileName,
  }
}

function publicWorkbookLicenseEvidence(args: {
  readonly licenseTitle: string
  readonly licenseUrl: string
  readonly licenseSpdxId?: string | null
}): PublicWorkbookLicenseEvidence {
  const license: PublicWorkbookLicenseEvidence = {
    spdxId: args.licenseSpdxId?.trim() || null,
    title: args.licenseTitle.trim(),
    evidenceUrl: args.licenseUrl.trim() || null,
  }
  if (!hasUsableLicenseEvidence(license)) {
    throw new Error('Shared workbook corpus sources require usable public license evidence')
  }
  return license
}

function downloadUrlForSharedWorkbookSource(sourceUrl: URL, explicitFileName?: string): string {
  const googleSheetId = googleSheetDocumentId(sourceUrl)
  if (googleSheetId) {
    return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(googleSheetId)}/export?format=xlsx`
  }
  const driveFileId = googleDriveFileId(sourceUrl)
  if (driveFileId) {
    if (!normalizedWorkbookFileName(explicitFileName)) {
      throw new Error('Expected --file-name for Google Drive file links because the share URL does not expose a workbook extension')
    }
    return `https://drive.google.com/uc?export=download&id=${encodeURIComponent(driveFileId)}`
  }
  return sourceUrl.href
}

function workbookFileNameForSharedSource(sourceUrl: URL): string | null {
  const googleSheetId = googleSheetDocumentId(sourceUrl)
  if (googleSheetId) {
    return `google-sheet-${safeFileNameToken(googleSheetId).slice(0, 32)}.xlsx`
  }
  return null
}

function googleSheetDocumentId(url: URL): string | null {
  if (url.hostname !== 'docs.google.com') {
    return null
  }
  const match = /^\/spreadsheets\/d\/([^/]+)/u.exec(url.pathname)
  return match?.[1] ?? null
}

function googleDriveFileId(url: URL): string | null {
  if (url.hostname !== 'drive.google.com') {
    return null
  }
  const pathMatch = /^\/file\/d\/([^/]+)/u.exec(url.pathname)
  if (pathMatch?.[1]) {
    return pathMatch[1]
  }
  return url.pathname === '/open' ? url.searchParams.get('id') : null
}

function workbookFileNameFromUrl(value: string): string | null {
  const parsed = parseAbsoluteHttpUrl(value, '--download-url')
  const lastSegment = parsed.pathname.split('/').findLast((segment) => segment.length > 0)
  if (!lastSegment) {
    return null
  }
  const decoded = decodeURIComponent(lastSegment)
  return normalizedWorkbookFileName(decoded)
}

function normalizedWorkbookFileName(value: string | undefined): string | null {
  const trimmed = value?.trim()
  if (!trimmed) {
    return null
  }
  const baseName = trimmed.split(/[\\/]/u).at(-1) ?? trimmed
  return isSpreadsheetFileName(baseName) ? baseName : null
}

function parseAbsoluteHttpUrl(value: string, label: string): URL {
  try {
    const parsed = new URL(value.trim())
    if (parsed.protocol !== 'https:' && parsed.protocol !== 'http:') {
      throw new Error('not http')
    }
    return parsed
  } catch {
    throw new Error(`Expected ${label} to be an absolute HTTP(S) URL`)
  }
}

function normalizeUrlForDeduplication(value: string): string {
  return value.trim().toLowerCase()
}

function safeFileNameToken(value: string): string {
  return value.replace(/[^A-Za-z0-9._-]+/gu, '-').replace(/^-+|-+$/gu, '') || 'shared-workbook'
}

function stableId(value: string): string {
  return sha256HexSync(Buffer.from(value)).slice(0, 16)
}
