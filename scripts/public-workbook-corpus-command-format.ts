import { isAbsolute, relative, resolve } from 'node:path'

import { publicCorpusStopMarkerOverrideEnvVar, publicCorpusStopMarkerOverrideFlag } from './public-workbook-corpus-cli.ts'

export interface PublicWorkbookLinkInput {
  readonly sourceUrl: string
  readonly downloadUrl: string
  readonly fileName: string
  readonly licenseTitle: string
  readonly licenseUrl: string
  readonly licenseSpdxId: string | null
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)

export interface SplitPublicWorkbookCorpusCommand {
  readonly command: string | null
  readonly blockedCommand: string | null
}

export function formatPublicWorkbookCorpusDiscoverCommand(args: {
  readonly cacheDir: string
  readonly limit: number
  readonly manifestPath: string
  readonly stopMarkerActive?: boolean
}): string {
  const parts = [
    'pnpm',
    'public-workbook-corpus:discover',
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--limit',
    String(args.limit),
  ]
  if (args.stopMarkerActive !== true) {
    return parts.map(shellQuote).join(' ')
  }
  return `${publicCorpusStopMarkerOverrideEnvVar}=1 ${[...parts, publicCorpusStopMarkerOverrideFlag].map(shellQuote).join(' ')}`
}

export function splitPublicWorkbookCorpusFetchCommand(args: {
  readonly cacheDir: string
  readonly fetchBatchSize?: number | null
  readonly limit: number
  readonly manifestPath: string
  readonly scriptName?: string
  readonly stopMarkerActive: boolean
}): SplitPublicWorkbookCorpusCommand {
  const parts = [
    'pnpm',
    args.scriptName ?? 'public-workbook-corpus:fetch',
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--limit',
    String(args.limit),
    ...(args.fetchBatchSize ? ['--fetch-batch-size', String(args.fetchBatchSize)] : []),
  ]
  return splitStopMarkerCommand(parts, args.stopMarkerActive)
}

export function publicWorkbookCorpusPlanStopMarker(active: boolean): {
  readonly active: boolean
  readonly requiresExplicitResume: boolean
  readonly overrideFlag: string
  readonly overrideEnvVar: string
} {
  return {
    active,
    requiresExplicitResume: active,
    overrideFlag: publicCorpusStopMarkerOverrideFlag,
    overrideEnvVar: publicCorpusStopMarkerOverrideEnvVar,
  }
}

export function formatPublicWorkbookCorpusAddLinkCommand(args: {
  readonly linkInput: PublicWorkbookLinkInput
  readonly manifestPath: string
}): string {
  return linkCommandParts('public-workbook-corpus:add-link', args.linkInput, args.manifestPath).map(shellQuote).join(' ')
}

export function formatPublicWorkbookCorpusLinkPlanCommand(args: {
  readonly linkInput: PublicWorkbookLinkInput
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly verifyCheckpointPath: string
}): string {
  return [
    ...linkCommandParts('public-workbook-corpus:link-plan', args.linkInput, args.manifestPath),
    '--scorecard',
    formatCommandPath(args.scorecardPath),
    '--verify-checkpoint',
    formatCommandPath(args.verifyCheckpointPath),
  ]
    .map(shellQuote)
    .join(' ')
}

export function formatPublicWorkbookCorpusFetchSourceCommand(args: {
  readonly cacheDir: string
  readonly manifestPath: string
  readonly sourceId: string
  readonly stopMarkerActive: boolean
}): string {
  const split = splitPublicWorkbookCorpusFetchSourceCommand(args)
  return split.command ?? split.blockedCommand ?? ''
}

export function formatPublicWorkbookCorpusVerifyArtifactCommand(args: {
  readonly artifactId: string
  readonly cacheDir: string
  readonly manifestPath: string
  readonly stopMarkerActive?: boolean
  readonly verifyCheckpointPath: string
}): string {
  const split = splitPublicWorkbookCorpusVerifyArtifactCommand(args)
  return split.command ?? split.blockedCommand ?? ''
}

export function formatPublicWorkbookCorpusStatusCommand(args: {
  readonly cacheDir: string
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly verifyCheckpointPath: string
}): string {
  return [
    'pnpm',
    'public-workbook-corpus:status',
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--scorecard',
    formatCommandPath(args.scorecardPath),
    '--verify-checkpoint',
    formatCommandPath(args.verifyCheckpointPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
  ]
    .map(shellQuote)
    .join(' ')
}

export function formatCommandPath(path: string): string {
  const absolutePath = resolve(path)
  const relativePath = relative(rootDir, absolutePath)
  return relativePath && !relativePath.startsWith('..') && !isAbsolute(relativePath) ? relativePath : path
}

function linkCommandParts(scriptName: string, input: PublicWorkbookLinkInput, manifestPath: string): string[] {
  return [
    'pnpm',
    scriptName,
    '--',
    '--manifest',
    formatCommandPath(manifestPath),
    '--source-url',
    input.sourceUrl,
    ...(input.downloadUrl ? ['--download-url', input.downloadUrl] : []),
    ...(input.fileName ? ['--file-name', input.fileName] : []),
    '--license-title',
    input.licenseTitle,
    '--license-url',
    input.licenseUrl,
    ...(input.licenseSpdxId ? ['--license-spdx', input.licenseSpdxId] : []),
  ]
}

export function splitPublicWorkbookCorpusFetchSourceCommand(args: {
  readonly cacheDir: string
  readonly manifestPath: string
  readonly sourceId: string
  readonly stopMarkerActive: boolean
}): SplitPublicWorkbookCorpusCommand {
  return splitStopMarkerCommand(publicWorkbookCorpusFetchSourceCommandParts(args), args.stopMarkerActive)
}

function publicWorkbookCorpusFetchSourceCommandParts(args: {
  readonly cacheDir: string
  readonly manifestPath: string
  readonly sourceId: string
}): string[] {
  return [
    'pnpm',
    'public-workbook-corpus:fetch-source',
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--source-id',
    args.sourceId,
  ]
}

export function splitPublicWorkbookCorpusVerifyArtifactCommand(args: {
  readonly artifactId: string
  readonly cacheDir: string
  readonly manifestPath: string
  readonly stopMarkerActive?: boolean
  readonly verifyCheckpointPath: string
}): SplitPublicWorkbookCorpusCommand {
  return splitStopMarkerCommand(publicWorkbookCorpusVerifyArtifactCommandParts(args), args.stopMarkerActive === true)
}

function publicWorkbookCorpusVerifyArtifactCommandParts(args: {
  readonly artifactId: string
  readonly cacheDir: string
  readonly manifestPath: string
  readonly verifyCheckpointPath: string
}): string[] {
  return [
    'pnpm',
    'public-workbook-corpus:verify-artifact',
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--verify-checkpoint',
    formatCommandPath(args.verifyCheckpointPath),
    '--artifact-id',
    args.artifactId,
    '--update-verify-checkpoint',
  ]
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function splitStopMarkerCommand(parts: readonly string[], stopMarkerActive: boolean): SplitPublicWorkbookCorpusCommand {
  const command = parts.map(shellQuote).join(' ')
  if (!stopMarkerActive) {
    return { command, blockedCommand: null }
  }
  return {
    command: null,
    blockedCommand: `${publicCorpusStopMarkerOverrideEnvVar}=1 ${[...parts, publicCorpusStopMarkerOverrideFlag].map(shellQuote).join(' ')}`,
  }
}
