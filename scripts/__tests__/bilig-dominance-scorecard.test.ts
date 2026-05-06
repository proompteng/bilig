import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildBiligDominanceScorecard, type BuildScorecardInput } from '../gen-bilig-dominance-scorecard.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('bilig dominance scorecard', () => {
  it('keeps the active Sheets/Excel goal separated from narrower HyperFormula evidence', () => {
    const scorecard = buildBiligDominanceScorecard(buildFixtureInput())

    expect(scorecard.goalStatus).toBe('active-not-achieved')
    expect(scorecard.claimPolicy.blanketTenXClaimAllowed).toBe(false)
    expect(scorecard.claimPolicy.workloadSpecificTenXWins).toEqual([
      {
        workload: 'rebuild-config-toggle',
        meanRatio: 0.05,
        p95Ratio: 0.08,
        comparisonTarget: 'HyperFormula',
      },
    ])
    expect(scorecard.summary.externalGoogleSheetsEvidence).toBe('not-captured-in-repo')
    expect(scorecard.summary.externalMicrosoftExcelEvidence).toBe('not-captured-in-repo')
    expect(scorecard.summary.formulaOfficeListedBreadthPercent).toBe(90.7)
    expect(scorecard.summary.formulaTrackedBreadthPercent).toBe(90.5)
    expect(scorecard.summary.largeWorkbookSloRowsCovered).toEqual([100_000, 250_000])
    expect(scorecard.summary.largeWorkbookSloPassed).toBe(true)
    expect(scorecard.sourceArtifacts.largeWorkbookSloScorecard).toBe('packages/benchmarks/baselines/large-workbook-slo-scorecard.json')
    expect(scorecard.categories.find((category) => category.id === 'large-workbook-scale')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining(['packages/benchmarks/baselines/large-workbook-slo-scorecard.json']),
    })
  })

  it('maps every explicit objective category into the checked-in generated artifact', () => {
    const artifact = parseGeneratedScorecard(
      readFileSync(resolve(repoRoot, 'packages/benchmarks/baselines/bilig-dominance-scorecard.json'), 'utf8'),
    )
    const categoryIds = artifact.categories.map((category) => category.id)

    expect(categoryIds).toEqual([
      'calculation-correctness',
      'recalculation-speed',
      'structural-edit-performance',
      'large-workbook-scale',
      'ui-responsiveness',
      'collaboration',
      'automation-api-extensibility',
      'import-export-compatibility',
      'auditability',
      'reliability',
      'security',
      'operator-developer-workflow',
    ])
  })

  it('wires the dominance check into fast CI generated checks', () => {
    const packageJson = readFileSync(resolve(repoRoot, 'package.json'), 'utf8')
    const runCi = readFileSync(resolve(repoRoot, 'scripts/run-ci.ts'), 'utf8')

    expect(packageJson).toContain('"dominance:check": "bun scripts/gen-bilig-dominance-scorecard.ts --check"')
    expect(packageJson).toContain('"large-workbook:slo:check": "bun scripts/gen-large-workbook-slo-scorecard.ts --check"')
    expect(runCi).toContain("pnpm('bilig dominance scorecard check', 'dominance:check')")
    expect(runCi).toContain("pnpm('large workbook SLO scorecard check', 'large-workbook:slo:check')")
  })
})

function parseGeneratedScorecard(source: string): { categories: Array<{ id: unknown }> } {
  const parsed: unknown = JSON.parse(source)
  if (!isRecord(parsed) || !Array.isArray(parsed['categories'])) {
    throw new Error('Generated scorecard must include a categories array')
  }

  return {
    categories: parsed['categories'].map((category) => {
      if (!isRecord(category)) {
        throw new Error('Generated scorecard category must be an object')
      }
      return {
        id: category['id'],
      }
    }),
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function buildFixtureInput(): BuildScorecardInput {
  return {
    competitiveArtifactPath: 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
    formulaSnapshotPath: 'packages/formula/src/__tests__/fixtures/formula-dominance-snapshot.json',
    largeWorkbookSloScorecardPath: 'packages/benchmarks/baselines/large-workbook-slo-scorecard.json',
    surfaceSnapshotPath: 'packages/headless/src/__tests__/fixtures/hyperformula-surface.json',
    formulaSnapshot: {
      schemaVersion: 1,
      formulaBreadth: {
        officeListed: {
          production: 461,
          total: 508,
          percent: 90.7,
        },
        tracked: {
          production: 475,
          total: 525,
          percent: 90.5,
        },
        missingOfficeFunctions: ['ACCRINT', 'MMULT'],
      },
      canonical: {
        summary: {
          production: 300,
          total: 300,
          percent: 100,
        },
        nonProductionRows: [],
      },
    },
    surfaceSnapshot: {
      hyperFormulaVersion: '3.2.0',
      hyperFormulaCommit: 'abc123',
      classSurface: {
        staticMembers: ['version'],
        staticMethods: ['buildEmpty'],
        instanceAccessors: ['dependencyGraph'],
        instanceMethods: ['getCellValue', 'setCellContents'],
      },
      configKeys: ['licenseKey', 'useColumnIndex'],
    },
    largeWorkbookSloScorecard: {
      schemaVersion: 1,
      suite: 'large-workbook-slo',
      summary: {
        coveredLargeWorkbookRows: [100_000, 250_000],
        allSloBudgetsPassed: true,
        allGateBudgetsPassed: true,
        headedBrowserFrameP95Evidence: 'not-captured',
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
      },
      measurements: [
        sloMeasurement('load100k', 'large-workbook-scale', 100_000, 230, 1500),
        sloMeasurement('load250k', 'large-workbook-scale', 250_000, 600, 1500),
        sloMeasurement('workerWarmStart100k', 'large-workbook-scale', 100_000, 12, 500),
        sloMeasurement('workerWarmStart250k', 'large-workbook-scale', 250_000, 17, 700),
        sloMeasurement('workerVisibleEdit10k', 'ui-responsiveness', 10_000, 4, 16),
        sloMeasurement('workerReconnectCatchUp100Pending', 'collaboration', 10_000, 270, 2000),
      ],
    },
    competitiveArtifact: {
      generatedAt: '2026-05-05T19:00:09.455Z',
      engines: {
        hyperformula: {
          commit: 'abc123',
          version: '3.2.0',
        },
      },
      scorecard: {
        comparableCount: 2,
        workpaperWins: 2,
        hyperformulaWins: 0,
        directionalMeanRatioGeomean: 0.2,
        directionalP95RatioGeomean: 0.25,
        worstWorkpaperToHyperFormulaMeanRatio: 0.7,
        worstMeanRatioWorkload: 'single-formula-edit-recalc',
        worstWorkpaperToHyperFormulaP95Ratio: 0.8,
        worstP95RatioWorkload: 'single-formula-edit-recalc',
      },
      families: [
        family('structural-rows', 0.7),
        family('structural-columns', 0.8),
        family('dirty-execution', 0.7),
        family('batch-edit', 0.6),
        family('rebuild', 0.05),
        family('range-read', 0.3),
        family('lookup-exact', 0.5),
        family('lookup-text', 0.4),
      ],
      results: [
        {
          workload: 'rebuild-config-toggle',
          comparable: true,
          comparison: {
            workpaperToHyperFormulaMeanRatio: 0.05,
            workpaperToHyperFormulaP95Ratio: 0.08,
          },
        },
        {
          workload: 'single-formula-edit-recalc',
          comparable: true,
          comparison: {
            workpaperToHyperFormulaMeanRatio: 0.7,
            workpaperToHyperFormulaP95Ratio: 0.8,
          },
        },
      ],
    },
  }
}

function sloMeasurement(
  id: string,
  category: BuildScorecardInput['largeWorkbookSloScorecard']['measurements'][number]['category'],
  materializedCells: number,
  actualP95: number,
  budgetP95: number,
): BuildScorecardInput['largeWorkbookSloScorecard']['measurements'][number] {
  return {
    id,
    category,
    label: id,
    materializedCells,
    corpusCaseId: null,
    metric: 'elapsedMs.p95',
    actualP95,
    budgetP95,
    gateBudgetP95: budgetP95,
    sampleCount: 3,
    passed: true,
    gatePassed: true,
  }
}

function family(familyName: string, ratio: number): BuildScorecardInput['competitiveArtifact']['families'][number] {
  return {
    family: familyName,
    scorecardEligible: true,
    comparableCount: 1,
    workpaperWins: 1,
    hyperformulaWins: 0,
    worstWorkpaperToHyperFormulaMeanRatio: ratio,
    worstMeanRatioWorkload: `${familyName}-workload`,
    worstWorkpaperToHyperFormulaP95Ratio: ratio,
    worstP95RatioWorkload: `${familyName}-workload`,
  }
}
