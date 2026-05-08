#!/usr/bin/env bun

import { existsSync } from 'node:fs'
import { isAbsolute, join, relative, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import {
  publicCorpusStopMarkerOverrideEnvVar,
  publicCorpusStopMarkerOverrideFlag,
  readNumberArg,
  readStringArg,
} from './public-workbook-corpus-cli.ts'
import { buildFeatureWitnessCoverage } from './public-workbook-corpus-completion-audit-helpers.ts'
import { readReusablePublicWorkbookCorpusCases } from './public-workbook-corpus-verify-checkpoint.ts'
import type { PublicWorkbookCorpusCase } from './public-workbook-corpus-types.ts'

export interface PublicWorkbookCorpusFeatureWitnessPlan {
  readonly schemaVersion: 1
  readonly mode: 'feature-witness-plan'
  readonly generatedAt: string
  readonly stopMarker: {
    readonly active: boolean
    readonly path: string
    readonly overrideFlag: string
    readonly overrideEnvVar: string
  }
  readonly recordedCaseCount: number
  readonly coverage: readonly {
    readonly id: string
    readonly label: string
    readonly totalCount: number
    readonly witnessCaseCount: number
    readonly needsWitness: boolean
    readonly discoveryQuery: string
    readonly commands: {
      readonly discover: string
    }
  }[]
  readonly missingWitnessCount: number
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'public-workbook-corpus-scorecard.json')
const defaultVerifyCheckpointPath = join(defaultCacheDir, 'verification-checkpoint.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')

const featureDiscoveryQueries: Readonly<Record<string, string>> = {
  'conditional formats': 'conditional formatting xlsx',
  charts: 'chart xlsx',
  formulas: 'formula xlsx',
  'merged ranges': 'merged cells xlsx',
  names: 'defined names xlsx',
  pivots: 'pivot table xlsx',
  styles: 'cell styles xlsx',
  tables: 'table xlsx',
  values: 'xlsx',
}

function main(): void {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultScorecardPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', defaultVerifyCheckpointPath))
  const stopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
  const discoveryLimit = readNumberArg('--discovery-limit', 10_000)
  const generatedAt = readStringArg('--generated-at', new Date().toISOString())
  const plan = buildPublicWorkbookCorpusFeatureWitnessPlan({
    cacheDir,
    cases: readReusablePublicWorkbookCorpusCases([scorecardPath, verifyCheckpointPath]),
    discoveryLimit,
    displayRootDir: rootDir,
    generatedAt,
    manifestPath,
    stopMarkerActive: existsSync(stopMarkerPath),
    stopMarkerPath,
  })
  if (process.argv.includes('--check')) {
    const findings = validatePublicWorkbookCorpusFeatureWitnessPlan(plan)
    if (findings.length > 0) {
      throw new Error(`Public workbook corpus feature witness plan is invalid: ${findings.join('; ')}`)
    }
    process.stdout.write(
      `${JSON.stringify(
        {
          mode: 'check',
          schemaVersion: plan.schemaVersion,
          generatedAt: plan.generatedAt,
          missingWitnessCount: plan.missingWitnessCount,
          recordedCaseCount: plan.recordedCaseCount,
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  process.stdout.write(`${JSON.stringify(plan, null, 2)}\n`)
}

export function buildPublicWorkbookCorpusFeatureWitnessPlan(args: {
  readonly cacheDir: string
  readonly cases: readonly PublicWorkbookCorpusCase[]
  readonly discoveryLimit: number
  readonly displayRootDir?: string
  readonly generatedAt: string
  readonly manifestPath: string
  readonly stopMarkerActive: boolean
  readonly stopMarkerPath: string
}): PublicWorkbookCorpusFeatureWitnessPlan {
  const coverage = buildFeatureWitnessCoverage(args.cases).map((entry) => {
    const discoveryQuery = featureDiscoveryQueries[entry.id] ?? `${entry.label} xlsx`
    return {
      id: entry.id,
      label: entry.label,
      totalCount: entry.totalCount,
      witnessCaseCount: entry.witnessCaseCount,
      needsWitness: entry.witnessCaseCount === 0,
      discoveryQuery,
      commands: {
        discover: guardedCommand(args.stopMarkerActive, [
          'pnpm',
          'public-workbook-corpus:discover',
          '--',
          '--manifest',
          commandPath(args.manifestPath, args.displayRootDir),
          '--cache-dir',
          commandPath(args.cacheDir, args.displayRootDir),
          '--query',
          discoveryQuery,
          '--limit',
          String(Math.max(0, Math.trunc(args.discoveryLimit))),
        ]),
      },
    }
  })
  return {
    schemaVersion: 1,
    mode: 'feature-witness-plan',
    generatedAt: args.generatedAt,
    stopMarker: {
      active: args.stopMarkerActive,
      path: commandPath(args.stopMarkerPath, args.displayRootDir),
      overrideFlag: publicCorpusStopMarkerOverrideFlag,
      overrideEnvVar: publicCorpusStopMarkerOverrideEnvVar,
    },
    recordedCaseCount: args.cases.length,
    coverage,
    missingWitnessCount: coverage.filter((entry) => entry.needsWitness).length,
  }
}

export function validatePublicWorkbookCorpusFeatureWitnessPlan(plan: PublicWorkbookCorpusFeatureWitnessPlan): string[] {
  const findings: string[] = []
  if (plan.schemaVersion !== 1) {
    findings.push(`unexpected schema version: ${String(plan.schemaVersion)}`)
  }
  if (!plan.generatedAt.trim()) {
    findings.push('generatedAt is empty')
  }
  if (plan.stopMarker.active && plan.stopMarker.overrideFlag !== publicCorpusStopMarkerOverrideFlag) {
    findings.push('stop-marker override flag does not match corpus CLI guard')
  }
  if (plan.stopMarker.active && plan.stopMarker.overrideEnvVar !== publicCorpusStopMarkerOverrideEnvVar) {
    findings.push('stop-marker override environment variable does not match corpus CLI guard')
  }
  if (plan.missingWitnessCount !== plan.coverage.filter((entry) => entry.needsWitness).length) {
    findings.push('missing witness count does not match coverage rows')
  }
  for (const entry of plan.coverage) {
    if (!entry.id.trim() || !entry.label.trim() || !entry.discoveryQuery.trim()) {
      findings.push(`feature witness row has empty identity or query: ${entry.id}`)
    }
    if (entry.needsWitness !== (entry.witnessCaseCount === 0)) {
      findings.push(`feature witness row has inconsistent needsWitness flag: ${entry.id}`)
    }
    if (plan.stopMarker.active && !entry.commands.discover.includes(publicCorpusStopMarkerOverrideFlag)) {
      findings.push(`feature witness discover command is missing stop-marker override: ${entry.id}`)
    }
  }
  return findings
}

function guardedCommand(stopMarkerActive: boolean, parts: readonly string[]): string {
  if (!stopMarkerActive) {
    return command(parts)
  }
  return `${publicCorpusStopMarkerOverrideEnvVar}=1 ${command([...parts, publicCorpusStopMarkerOverrideFlag])}`
}

function commandPath(path: string, displayRootDir: string | undefined): string {
  if (!displayRootDir) {
    return path
  }
  const relativePath = relative(displayRootDir, path)
  if (!relativePath || relativePath.startsWith('..') || isAbsolute(relativePath)) {
    return path
  }
  return relativePath
}

function command(parts: readonly string[]): string {
  return parts.map(shellQuote).join(' ')
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
