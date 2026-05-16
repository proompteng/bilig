#!/usr/bin/env bun

import { readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { formatJsonForRepo } from './scorecard-format.ts'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const outputPath = join(repoRoot, 'docs', 'public-evidence.json')
const checkMode = process.argv.includes('--check')

interface LaneScorecard {
  readonly lane: string
  readonly comparableCount: number
  readonly workpaperWins: number
  readonly hyperformulaWins: number
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly worstWorkpaperToHyperFormulaMeanRatio: number
  readonly worstMeanRatioWorkload: string
  readonly worstWorkpaperToHyperFormulaP95Ratio: number
  readonly worstP95RatioWorkload: string
}

interface PublicEvidence {
  readonly schemaVersion: 1
  readonly package: {
    readonly name: string
    readonly version: string
    readonly releasePleaseManifestVersion: string
    readonly mcpServerVersion: string
    readonly mcpPackageVersion: string
  }
  readonly workpaperVsHyperFormula: {
    readonly artifactPath: string
    readonly generatedAt: string
    readonly sampleCount: number
    readonly warmupCount: number
    readonly workpaperPackageVersion: string
    readonly hyperformulaVersion: string
    readonly hyperformulaCommit: string
    readonly overall: LaneScorecard
    readonly publicLane: LaneScorecard
    readonly holdout: LaneScorecard
    readonly meanAndP95WinCount: number
    readonly p95HoldoutCount: number
  }
}

function asRecord(value: unknown, context: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${context} must be an object`)
  }
  return Object.fromEntries(Object.entries(value))
}

function readString(record: Record<string, unknown>, key: string, context: string): string {
  const value = record[key]
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`${context}.${key} must be a non-empty string`)
  }
  return value
}

function readNumber(record: Record<string, unknown>, key: string, context: string): number {
  const value = record[key]
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`${context}.${key} must be a finite number`)
  }
  return value
}

function readJsonRecord(path: string, context: string): Promise<Record<string, unknown>> {
  return readFile(path, 'utf8').then((content) => asRecord(JSON.parse(content) as unknown, context))
}

function readLane(value: unknown, context: string): LaneScorecard {
  const record = asRecord(value, context)
  return {
    lane: readString(record, 'lane', context),
    comparableCount: readNumber(record, 'comparableCount', context),
    workpaperWins: readNumber(record, 'workpaperWins', context),
    hyperformulaWins: readNumber(record, 'hyperformulaWins', context),
    directionalMeanRatioGeomean: readNumber(record, 'directionalMeanRatioGeomean', context),
    directionalP95RatioGeomean: readNumber(record, 'directionalP95RatioGeomean', context),
    worstWorkpaperToHyperFormulaMeanRatio: readNumber(record, 'worstWorkpaperToHyperFormulaMeanRatio', context),
    worstMeanRatioWorkload: readString(record, 'worstMeanRatioWorkload', context),
    worstWorkpaperToHyperFormulaP95Ratio: readNumber(record, 'worstWorkpaperToHyperFormulaP95Ratio', context),
    worstP95RatioWorkload: readString(record, 'worstP95RatioWorkload', context),
  }
}

function headline(lane: Pick<LaneScorecard, 'workpaperWins' | 'comparableCount'>): string {
  return `${lane.workpaperWins.toString()}/${lane.comparableCount.toString()}`
}

function ratio3(value: number): string {
  return `${value.toFixed(3).replace(/0+$/u, '').replace(/\.$/u, '')}x`
}

function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

function requireNotIncludes(haystack: string, needle: string, context: string): void {
  if (haystack.includes(needle)) {
    throw new Error(`${context} must not include stale public evidence token ${needle}`)
  }
}

async function buildEvidence(): Promise<PublicEvidence> {
  const [packageManifest, serverManifest, releasePleaseManifest, benchmarkArtifact, leadershipScorecard] = await Promise.all([
    readJsonRecord(join(repoRoot, 'packages', 'headless', 'package.json'), 'packages/headless/package.json'),
    readJsonRecord(join(repoRoot, 'packages', 'headless', 'server.json'), 'packages/headless/server.json'),
    readJsonRecord(join(repoRoot, '.release-please-manifest.json'), '.release-please-manifest.json'),
    readJsonRecord(join(repoRoot, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-hyperformula.json'), 'workpaper benchmark artifact'),
    readJsonRecord(
      join(repoRoot, 'packages', 'benchmarks', 'baselines', 'headless-performance-leadership-scorecard.json'),
      'headless performance leadership scorecard',
    ),
  ])

  const packageName = readString(packageManifest, 'name', 'packages/headless/package.json')
  const packageVersion = readString(packageManifest, 'version', 'packages/headless/package.json')
  const releasePleaseVersion = readString(releasePleaseManifest, 'packages/headless', '.release-please-manifest.json')
  const serverVersion = readString(serverManifest, 'version', 'packages/headless/server.json')
  const serverPackages = serverManifest['packages']
  if (!Array.isArray(serverPackages)) {
    throw new Error('packages/headless/server.json.packages must be an array')
  }
  const npmServerPackage = serverPackages
    .map((entry) => asRecord(entry, 'packages/headless/server.json.packages[]'))
    .find((entry) => entry['identifier'] === packageName)
  if (!npmServerPackage) {
    throw new Error(`packages/headless/server.json is missing npm package entry for ${packageName}`)
  }
  const mcpPackageVersion = readString(npmServerPackage, 'version', 'packages/headless/server.json package entry')

  const benchmark = asRecord(benchmarkArtifact['benchmark'], 'workpaper benchmark artifact.benchmark')
  const engines = asRecord(benchmarkArtifact['engines'], 'workpaper benchmark artifact.engines')
  const workpaperEngine = asRecord(engines['workpaper'], 'workpaper benchmark artifact.engines.workpaper')
  const hyperformulaEngine = asRecord(engines['hyperformula'], 'workpaper benchmark artifact.engines.hyperformula')
  const scorecard = asRecord(benchmarkArtifact['scorecard'], 'workpaper benchmark artifact.scorecard')
  const scorecards = asRecord(scorecard['scorecards'], 'workpaper benchmark artifact.scorecard.scorecards')
  const leadershipSummary = asRecord(leadershipScorecard['summary'], 'headless performance leadership scorecard.summary')
  const p95Holdouts = leadershipSummary['p95Holdouts']
  if (!Array.isArray(p95Holdouts)) {
    throw new Error('headless performance leadership scorecard.summary.p95Holdouts must be an array')
  }

  const alignedVersionEntries = [
    ['.release-please-manifest.json', releasePleaseVersion],
    ['packages/headless/server.json', serverVersion],
    ['packages/headless/server.json package entry', mcpPackageVersion],
    [
      'packages/benchmarks/baselines/workpaper-vs-hyperformula.json engine metadata',
      readString(workpaperEngine, 'version', 'workpaper engine'),
    ],
  ] as const
  for (const [context, version] of alignedVersionEntries) {
    if (version !== packageVersion) {
      throw new Error(`${context} version ${version} must match ${packageName}@${packageVersion}`)
    }
  }

  return {
    schemaVersion: 1,
    package: {
      name: packageName,
      version: packageVersion,
      releasePleaseManifestVersion: releasePleaseVersion,
      mcpServerVersion: serverVersion,
      mcpPackageVersion,
    },
    workpaperVsHyperFormula: {
      artifactPath: 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
      generatedAt: readString(benchmarkArtifact, 'generatedAt', 'workpaper benchmark artifact'),
      sampleCount: readNumber(benchmark, 'sampleCount', 'workpaper benchmark artifact.benchmark'),
      warmupCount: readNumber(benchmark, 'warmupCount', 'workpaper benchmark artifact.benchmark'),
      workpaperPackageVersion: readString(workpaperEngine, 'version', 'workpaper engine'),
      hyperformulaVersion: readString(hyperformulaEngine, 'version', 'hyperformula engine'),
      hyperformulaCommit: readString(hyperformulaEngine, 'commit', 'hyperformula engine'),
      overall: readLane(scorecards['overall'], 'workpaper benchmark artifact.scorecard.scorecards.overall'),
      publicLane: readLane(scorecards['public'], 'workpaper benchmark artifact.scorecard.scorecards.public'),
      holdout: readLane(scorecards['holdout'], 'workpaper benchmark artifact.scorecard.scorecards.holdout'),
      meanAndP95WinCount: readNumber(leadershipSummary, 'meanAndP95WinCount', 'headless performance leadership scorecard.summary'),
      p95HoldoutCount: p95Holdouts.length,
    },
  }
}

async function assertPublicSurfaces(evidence: PublicEvidence): Promise<void> {
  const benchmark = evidence.workpaperVsHyperFormula
  const overall = benchmark.overall
  const publicLane = benchmark.publicLane
  const holdout = benchmark.holdout
  const meanHeadline = headline(overall)
  const publicHeadline = headline(publicLane)
  const holdoutHeadline = headline(holdout)
  const meanAndP95Headline = `${benchmark.meanAndP95WinCount.toString()}/${overall.comparableCount.toString()}`
  const p95Ratio = ratio3(overall.worstWorkpaperToHyperFormulaP95Ratio)
  const scannedPaths = [
    'README.md',
    'packages/headless/README.md',
    'docs/index.html',
    'docs/what-workpaper-benchmark-proves.md',
    'docs/headless-workpaper-benchmark-evidence.md',
    'docs/hyperformula-alternative-headless-workpaper.md',
    'docs/why-agents-need-workbook-apis.md',
    'docs/where-bilig-is-not-excel-compatible-yet.md',
    'docs/dev-to-workbook-apis-post.md',
    'docs/local-workpaper-benchmark-walkthrough.md',
    'docs/llms.txt',
  ] as const
  const staleTokens = [
    '46/46',
    '46 of 46',
    '37/52',
    '37 of 52',
    '41/52',
    'lookup-approximate-duplicates` at `1.043x',
    '1.043x</code>',
    '0.7489873822783492',
    '0.7354308040905896',
    '3.777197275754674',
    '2026-05-15T04:04:38.038Z',
    '29/40',
    '42/57',
    '10/17',
    '45/57',
    '33/40',
    '12/17',
    '6.493x',
    '6.4928649835338925',
    '0.7240066714283266',
    '0.7330720883107373',
    '6.152744637995318',
    '2026-05-16T03:46:32.343Z',
    '5.397x',
    '5.396915291352403',
    '0.7165647582609914',
    '0.7159317903242608',
    '5.603036418492105',
    '2026-05-16T03:34:41.623Z',
    '8.722x',
    '8.72243346007912',
    '7.981x',
    '7.981245577368439',
    '7.649x',
    '7.648801690864582',
    '7.541560588587015',
    '0.7553949494105464',
    '0.7510834854399419',
    '0.7577447189137954',
    '0.7980273811097534',
    '0.7442626408109101',
    '0.7724839680358417',
    '2026-05-16T02:12:30.841Z',
    '2026-05-16T02:38:29.935Z',
    '2026-05-16T02:45:18.556Z',
    '@bilig/headless` `0.14.23`',
    '@bilig/headless` `0.14.25`',
  ] as const

  const scannedContents = await Promise.all(scannedPaths.map(async (path) => [path, await readFile(join(repoRoot, path), 'utf8')] as const))
  for (const [path, content] of scannedContents) {
    for (const token of staleTokens) {
      requireNotIncludes(content, token, path)
    }
  }

  const [readme, headlessReadme, index, benchmarkExplainer, benchmarkEvidence, hyperformulaAlternative, svgCard] = await Promise.all([
    readFile(join(repoRoot, 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'index.html'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'what-workpaper-benchmark-proves.md'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'headless-workpaper-benchmark-evidence.md'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'hyperformula-alternative-headless-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'assets', 'workpaper-benchmark-card.svg'), 'utf8'),
  ])

  for (const [path, content] of [
    ['README.md', readme],
    ['packages/headless/README.md', headlessReadme],
  ] as const) {
    requireIncludes(content, `[\`${meanHeadline}\` comparable WorkPaper mean wins]`, path)
    requireIncludes(content, `\`${overall.worstP95RatioWorkload}\``, path)
    requireIncludes(content, `\`${p95Ratio}\``, path)
  }

  requireIncludes(index, `<strong>${meanHeadline}</strong>`, 'docs/index.html')
  requireIncludes(
    index,
    `${overall.workpaperWins.toString()} of ${overall.comparableCount.toString()} comparable mean-latency rows`,
    'docs/index.html',
  )
  requireIncludes(index, `${overall.worstP95RatioWorkload} is slower at p95: <code>${p95Ratio}</code>`, 'docs/index.html')

  for (const [path, content] of [
    ['docs/what-workpaper-benchmark-proves.md', benchmarkExplainer],
    ['docs/headless-workpaper-benchmark-evidence.md', benchmarkEvidence],
  ] as const) {
    requireIncludes(content, `\`${meanHeadline}\` mean-latency wins`, path)
    requireIncludes(
      content,
      `| Overall |                 \`${overall.comparableCount.toString()}\` |                \`${overall.workpaperWins.toString()}\` |`,
      path,
    )
    requireIncludes(
      content,
      `| Public  |                 \`${publicLane.comparableCount.toString()}\` |                \`${publicLane.workpaperWins.toString()}\` |`,
      path,
    )
    requireIncludes(
      content,
      `| Holdout |                 \`${holdout.comparableCount.toString()}\` |                 \`${holdout.workpaperWins.toString()}\` |`,
      path,
    )
    requireIncludes(content, `generated at \`${benchmark.generatedAt}\``, path)
    requireIncludes(content, `\`${overall.directionalMeanRatioGeomean.toString()}\``, path)
    requireIncludes(content, `\`${overall.directionalP95RatioGeomean.toString()}\``, path)
    requireIncludes(content, `\`${meanAndP95Headline}\` workloads winning both`, path)
    requireIncludes(content, `\`${overall.worstP95RatioWorkload}\``, path)
    requireIncludes(content, `\`${overall.worstWorkpaperToHyperFormulaP95Ratio.toString()}\``, path)
  }

  requireIncludes(hyperformulaAlternative, `\`${meanHeadline}\` mean wins`, 'docs/hyperformula-alternative-headless-workpaper.md')
  requireIncludes(
    hyperformulaAlternative,
    `\`${publicHeadline}\` public-lane mean wins`,
    'docs/hyperformula-alternative-headless-workpaper.md',
  )
  requireIncludes(
    hyperformulaAlternative,
    `\`${holdoutHeadline}\` holdout-lane mean wins`,
    'docs/hyperformula-alternative-headless-workpaper.md',
  )

  requireIncludes(svgCard, `>${meanHeadline}</text>`, 'docs/assets/workpaper-benchmark-card.svg')
  requireIncludes(svgCard, `>${publicHeadline}</text>`, 'docs/assets/workpaper-benchmark-card.svg')
  requireIncludes(svgCard, `>${holdoutHeadline}</text>`, 'docs/assets/workpaper-benchmark-card.svg')
  requireIncludes(svgCard, `${overall.worstP95RatioWorkload} p95: ${p95Ratio}`, 'docs/assets/workpaper-benchmark-card.svg')
}

const evidence = await buildEvidence()
const rendered = formatJsonForRepo(`${JSON.stringify(evidence, null, 2)}\n`)

if (checkMode) {
  const current = await readFile(outputPath, 'utf8')
  if (current !== rendered) {
    throw new Error('docs/public-evidence.json is out of date. Run: pnpm public:evidence:generate')
  }
  await assertPublicSurfaces(evidence)
  console.log(
    JSON.stringify(
      {
        ok: true,
        outputPath,
        version: evidence.package.version,
        workpaperMeanWins: headline(evidence.workpaperVsHyperFormula.overall),
        workpaperMeanAndP95Wins: `${evidence.workpaperVsHyperFormula.meanAndP95WinCount.toString()}/${evidence.workpaperVsHyperFormula.overall.comparableCount.toString()}`,
      },
      null,
      2,
    ),
  )
} else {
  await writeFile(outputPath, rendered)
  console.log(
    JSON.stringify(
      {
        mode: 'write',
        outputPath,
        version: evidence.package.version,
        workpaperMeanWins: headline(evidence.workpaperVsHyperFormula.overall),
      },
      null,
      2,
    ),
  )
}
