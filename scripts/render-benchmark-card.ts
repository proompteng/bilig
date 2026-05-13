import type { Buffer } from 'node:buffer'
import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const cardWidth = 1280
const cardHeight = 720
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const artifactPath = join(repoRoot, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-hyperformula.json')
const outputPath = join(repoRoot, 'docs', 'assets', 'workpaper-benchmark-card.png')
const svgOutputPath = join(repoRoot, 'docs', 'assets', 'workpaper-benchmark-card.svg')
const checkMode = process.argv.includes('--check')

interface LaneScorecard {
  readonly lane: string
  readonly comparableCount: number
  readonly workpaperWins: number
  readonly hyperformulaWins: number
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly worstWorkpaperToHyperFormulaP95Ratio: number
  readonly worstP95RatioWorkload: string
}

interface BenchmarkSummary {
  readonly generatedAt: string
  readonly overall: LaneScorecard
  readonly publicLane: LaneScorecard
  readonly holdout: LaneScorecard
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;')
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

function readLaneScorecard(value: unknown, context: string): LaneScorecard {
  const record = asRecord(value, context)
  return {
    lane: readString(record, 'lane', context),
    comparableCount: readNumber(record, 'comparableCount', context),
    workpaperWins: readNumber(record, 'workpaperWins', context),
    hyperformulaWins: readNumber(record, 'hyperformulaWins', context),
    directionalMeanRatioGeomean: readNumber(record, 'directionalMeanRatioGeomean', context),
    directionalP95RatioGeomean: readNumber(record, 'directionalP95RatioGeomean', context),
    worstWorkpaperToHyperFormulaP95Ratio: readNumber(record, 'worstWorkpaperToHyperFormulaP95Ratio', context),
    worstP95RatioWorkload: readString(record, 'worstP95RatioWorkload', context),
  }
}

async function readBenchmarkSummary(): Promise<BenchmarkSummary> {
  const artifact = asRecord(JSON.parse(await readFile(artifactPath, 'utf8')) as unknown, 'benchmark artifact')
  const scorecard = asRecord(artifact.scorecard, 'benchmark artifact.scorecard')
  const scorecards = asRecord(scorecard.scorecards, 'benchmark artifact.scorecard.scorecards')

  return {
    generatedAt: readString(artifact, 'generatedAt', 'benchmark artifact'),
    overall: readLaneScorecard(scorecards.overall, 'benchmark artifact.scorecard.scorecards.overall'),
    publicLane: readLaneScorecard(scorecards.public, 'benchmark artifact.scorecard.scorecards.public'),
    holdout: readLaneScorecard(scorecards.holdout, 'benchmark artifact.scorecard.scorecards.holdout'),
  }
}

function formatRatio(value: number): string {
  return `${value.toFixed(3).replace(/0+$/, '').replace(/\.$/, '')}x`
}

function formatPercentLower(ratio: number): string {
  return `${Math.round((1 - ratio) * 100).toString()}% lower`
}

function renderLane(scorecard: LaneScorecard, y: number, label: string): string {
  const width = Math.max(0, Math.min(460, (scorecard.workpaperWins / scorecard.comparableCount) * 460))
  return String.raw`<g transform="translate(88 ${y})">
  <text x="0" y="0" fill="#526273" font-family="Inter, Arial, Helvetica, sans-serif" font-size="24" font-weight="760">${escapeXml(label)}</text>
  <text x="438" y="0" text-anchor="end" fill="#111820" font-family="Inter, Arial, Helvetica, sans-serif" font-size="25" font-weight="820">${scorecard.workpaperWins.toString()}/${scorecard.comparableCount.toString()}</text>
  <rect x="0" y="16" width="460" height="22" rx="8" fill="#e6edf3"/>
  <rect x="0" y="16" width="${width.toFixed(2)}" height="22" rx="8" fill="#147a4b"/>
</g>`
}

function buildSvg(summary: BenchmarkSummary): string {
  const generatedDate = summary.generatedAt.slice(0, 10)
  const meanLower = formatPercentLower(summary.overall.directionalMeanRatioGeomean)
  const p95Lower = formatPercentLower(summary.overall.directionalP95RatioGeomean)
  const meanRatio = formatRatio(summary.overall.directionalMeanRatioGeomean)
  const p95Ratio = formatRatio(summary.overall.directionalP95RatioGeomean)
  const caveatRatio = formatRatio(summary.overall.worstWorkpaperToHyperFormulaP95Ratio)

  return String.raw`<svg xmlns="http://www.w3.org/2000/svg" width="${cardWidth}" height="${cardHeight}" viewBox="0 0 ${cardWidth} ${cardHeight}">
  <defs>
    <linearGradient id="panel" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#ffffff"/>
      <stop offset="1" stop-color="#eef4f7"/>
    </linearGradient>
    <linearGradient id="dark" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#101820"/>
      <stop offset="1" stop-color="#192838"/>
    </linearGradient>
    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="24" stdDeviation="28" flood-color="#102033" flood-opacity="0.16"/>
    </filter>
  </defs>

  <rect width="${cardWidth}" height="${cardHeight}" fill="#f4f7f9"/>
  <path d="M0 590 C190 520 340 674 520 600 C700 526 814 502 1000 560 C1135 602 1210 592 1280 542 L1280 720 L0 720 Z" fill="#e4edf2"/>
  <g opacity="0.5" stroke="#d9e3ea" stroke-width="1">
    <path d="M96 0 V720"/>
    <path d="M224 0 V720"/>
    <path d="M352 0 V720"/>
    <path d="M480 0 V720"/>
    <path d="M608 0 V720"/>
    <path d="M736 0 V720"/>
    <path d="M864 0 V720"/>
    <path d="M992 0 V720"/>
    <path d="M1120 0 V720"/>
    <path d="M1248 0 V720"/>
  </g>

  <g transform="translate(64 58)">
    <rect x="0" y="0" width="56" height="56" rx="13" fill="#101820"/>
    <rect x="13" y="13" width="12" height="12" rx="3" fill="#77c98b"/>
    <rect x="31" y="13" width="12" height="12" rx="3" fill="#77c98b"/>
    <rect x="13" y="31" width="12" height="12" rx="3" fill="#77c98b"/>
    <rect x="31" y="31" width="12" height="12" rx="3" fill="#77c98b"/>
    <text x="74" y="38" fill="#111820" font-family="Inter, Arial, Helvetica, sans-serif" font-size="34" font-weight="820">bilig</text>
  </g>

  <text x="64" y="182" fill="#147a4b" font-family="Inter, Arial, Helvetica, sans-serif" font-size="28" font-weight="820">WorkPaper benchmark evidence</text>
  <text x="64" y="252" fill="#111820" font-family="Inter, Arial, Helvetica, sans-serif" font-size="86" font-weight="830" letter-spacing="0">46/46</text>
  <text x="66" y="310" fill="#111820" font-family="Inter, Arial, Helvetica, sans-serif" font-size="48" font-weight="800" letter-spacing="0">mean wins</text>
  <text x="66" y="354" fill="#526273" font-family="Inter, Arial, Helvetica, sans-serif" font-size="26" font-weight="650">Comparable HyperFormula-style workloads</text>

  <g filter="url(#shadow)">
    <rect x="64" y="382" width="560" height="278" rx="22" fill="url(#panel)" stroke="#cbd8e2"/>
  </g>
  ${renderLane(summary.overall, 428, 'Overall')}
  ${renderLane(summary.publicLane, 508, 'Public')}
  ${renderLane(summary.holdout, 588, 'Holdout')}

  <g filter="url(#shadow)">
    <rect x="666" y="92" width="550" height="546" rx="26" fill="url(#dark)"/>
  </g>
  <text x="710" y="154" fill="#9fb2c5" font-family="Inter, Arial, Helvetica, sans-serif" font-size="22" font-weight="760">Geomean latency ratios</text>

  <g transform="translate(710 196)">
    <rect x="0" y="0" width="462" height="116" rx="18" fill="#ffffff" fill-opacity="0.08" stroke="#36506a"/>
    <text x="28" y="43" fill="#ffffff" font-family="Inter, Arial, Helvetica, sans-serif" font-size="34" font-weight="840">Mean ${meanRatio}</text>
    <text x="28" y="82" fill="#7ed894" font-family="Inter, Arial, Helvetica, sans-serif" font-size="24" font-weight="780">${meanLower} than HyperFormula</text>
  </g>

  <g transform="translate(710 340)">
    <rect x="0" y="0" width="462" height="116" rx="18" fill="#ffffff" fill-opacity="0.08" stroke="#36506a"/>
    <text x="28" y="43" fill="#ffffff" font-family="Inter, Arial, Helvetica, sans-serif" font-size="34" font-weight="840">p95 ${p95Ratio}</text>
    <text x="28" y="82" fill="#7ed894" font-family="Inter, Arial, Helvetica, sans-serif" font-size="24" font-weight="780">${p95Lower} geomean lead</text>
  </g>

  <g transform="translate(710 490)">
    <rect x="0" y="0" width="462" height="92" rx="18" fill="#f4d991" fill-opacity="0.14" stroke="#8a6a23"/>
    <text x="28" y="35" fill="#f3d98b" font-family="Inter, Arial, Helvetica, sans-serif" font-size="22" font-weight="820">Visible caveat</text>
    <text x="28" y="66" fill="#ffffff" font-family="Inter, Arial, Helvetica, sans-serif" font-size="21" font-weight="650">${escapeXml(summary.overall.worstP95RatioWorkload)} p95: ${caveatRatio}</text>
  </g>

  <text x="64" y="686" fill="#526273" font-family="SFMono-Regular, Menlo, Consolas, monospace" font-size="20">pnpm workpaper:bench:competitive:check</text>
  <text x="1216" y="686" text-anchor="end" fill="#526273" font-family="Inter, Arial, Helvetica, sans-serif" font-size="18" font-weight="650">artifact ${generatedDate}</text>
</svg>`
}

function execFileBuffer(file: string, args: readonly string[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    execFile(
      file,
      [...args],
      {
        encoding: 'buffer',
        maxBuffer: 16 * 1024 * 1024,
      },
      (error, stdout, stderr) => {
        if (error !== null) {
          const message = Buffer.isBuffer(stderr) ? stderr.toString('utf8') : String(stderr)
          reject(new Error(`${file} failed: ${message.trim() || error.message}`))
          return
        }
        resolve(Buffer.from(stdout))
      },
    )
  })
}

async function renderPng(svg: string): Promise<Buffer> {
  const tempRoot = await mkdtemp(join(tmpdir(), 'bilig-benchmark-card-'))
  const svgPath = join(tempRoot, 'benchmark-card.svg')

  try {
    await writeFile(svgPath, svg)
    return await execFileBuffer('rsvg-convert', ['--format=png', '--width', String(cardWidth), '--height', String(cardHeight), svgPath])
  } finally {
    await rm(tempRoot, { recursive: true, force: true })
  }
}

function validatePngDimensions(image: Buffer, expectedWidth: number, expectedHeight: number, context: string): void {
  const pngSignature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  if (image.length < 24 || !image.subarray(0, pngSignature.length).equals(pngSignature)) {
    throw new Error(`${context} is not a PNG file`)
  }

  const actualWidth = image.readUInt32BE(16)
  const actualHeight = image.readUInt32BE(20)
  if (actualWidth !== expectedWidth || actualHeight !== expectedHeight) {
    throw new Error(
      `${context} must be ${expectedWidth.toString()}x${expectedHeight.toString()}, got ${actualWidth.toString()}x${actualHeight.toString()}`,
    )
  }
}

const svg = buildSvg(await readBenchmarkSummary())

if (checkMode) {
  const existingSvg = await readFile(svgOutputPath, 'utf8')
  if (existingSvg !== svg) {
    throw new Error(`${svgOutputPath} is stale. Run pnpm docs:benchmark-card:generate.`)
  }
  validatePngDimensions(await readFile(outputPath), cardWidth, cardHeight, outputPath)
  console.log(`benchmark card source is current: ${svgOutputPath}`)
  console.log(`benchmark card PNG dimensions are current: ${outputPath}`)
} else {
  await mkdir(dirname(outputPath), { recursive: true })
  await writeFile(svgOutputPath, svg)
  await writeFile(outputPath, await renderPng(svg))
  console.log(`wrote ${svgOutputPath}`)
  console.log(`wrote ${outputPath}`)
}
