#!/usr/bin/env bun
import { existsSync } from 'node:fs'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { resolve } from 'node:path'

const defaultCoveragePath = fileURLToPath(new URL('../coverage/coverage-final.json', import.meta.url))

const thresholds = [
  { label: 'packages/core/src', prefix: '/packages/core/src/', lines: 91 },
  { label: 'packages/formula/src', prefix: '/packages/formula/src/', lines: 91 },
  { label: 'packages/renderer/src', prefix: '/packages/renderer/src/', lines: 91 },
]

const ignoredSuffixes = ['/index.ts', '/snapshot.ts', '/ast.ts']

function lineStatsForFile(fileCoverage: {
  s?: Record<string, number>
  statementMap?: Record<string, { start: { line: number }; end: { line: number } }>
}) {
  const totalLines = new Set<number>()
  const coveredLines = new Set<number>()

  if (!fileCoverage.statementMap || !fileCoverage.s) {
    return { total: 0, covered: 0 }
  }

  for (const [statementId, location] of Object.entries(fileCoverage.statementMap)) {
    const hits = fileCoverage.s[statementId] ?? 0
    for (let line = location.start.line; line <= location.end.line; line += 1) {
      totalLines.add(line)
      if (hits > 0) {
        coveredLines.add(line)
      }
    }
  }

  return {
    total: totalLines.size,
    covered: coveredLines.size,
  }
}

function aggregatePrefix(
  prefix: string,
  coverageData: Record<
    string,
    { s?: Record<string, number>; statementMap?: Record<string, { start: { line: number }; end: { line: number } }> }
  >,
) {
  let total = 0
  let covered = 0

  for (const [filePath, fileCoverage] of Object.entries(coverageData)) {
    if (!filePath.includes(prefix)) {
      continue
    }
    if (ignoredSuffixes.some((suffix) => filePath.endsWith(suffix))) {
      continue
    }
    const stats = lineStatsForFile(fileCoverage)
    total += stats.total
    covered += stats.covered
  }

  if (total === 0) {
    throw new Error(`No coverage entries found for ${prefix}`)
  }

  return (covered / total) * 100
}

function isDirectInvocation(): boolean {
  const entryPoint = process.argv[1]
  if (!entryPoint) {
    return false
  }
  try {
    return pathToFileURL(resolve(entryPoint)).href === import.meta.url
  } catch {
    return false
  }
}

export function resolveCoverageFilePath(): string {
  const explicitPath = process.env['BILIG_COVERAGE_FILE']
  if (explicitPath) {
    return resolve(explicitPath)
  }

  const reportsDirectory = process.env['BILIG_COVERAGE_DIR']
  if (reportsDirectory) {
    return resolve(reportsDirectory, 'coverage-final.json')
  }

  return defaultCoveragePath
}

export async function runCoverageContracts(path = resolveCoverageFilePath()): Promise<void> {
  const resolvedPath = path instanceof URL ? fileURLToPath(path) : resolve(path)
  if (!existsSync(resolvedPath)) {
    throw new Error(`Coverage file not found at ${resolvedPath}`)
  }

  const coverage = await Bun.file(resolvedPath).json()

  const results = thresholds.map((threshold) => ({
    label: threshold.label,
    linesPct: aggregatePrefix(threshold.prefix, coverage),
    requiredLinesPct: threshold.lines,
  }))

  for (const result of results) {
    if (result.linesPct < result.requiredLinesPct) {
      throw new Error(`${result.label} line coverage is below target: ${result.linesPct.toFixed(2)}% < ${result.requiredLinesPct}%`)
    }
  }

  console.log(JSON.stringify({ results }, null, 2))
}

if (isDirectInvocation()) {
  await runCoverageContracts()
}
