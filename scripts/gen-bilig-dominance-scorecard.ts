#!/usr/bin/env bun

import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { parseAuditabilityScorecard, type AuditabilityScorecard } from './gen-auditability-scorecard.ts'
import { parseImportExportFidelityScorecard, type ImportExportFidelityScorecard } from './gen-import-export-fidelity-scorecard.ts'
import { parseSecurityPostureScorecard, type SecurityPostureScorecard } from './gen-security-posture-scorecard.ts'

export type DominanceStatus = 'repo-proved-lead' | 'partial-repo-evidence' | 'target-only'

export interface RatioSummary {
  percent: number
  production: number
  total: number
}

export interface FormulaDominanceSnapshot {
  schemaVersion: 1
  formulaBreadth: {
    officeListed: RatioSummary
    tracked: RatioSummary
    missingOfficeFunctions: string[]
  }
  canonical: {
    summary: RatioSummary
    nonProductionRows: unknown[]
  }
}

export interface LargeWorkbookSloMeasurement {
  id: string
  category: 'large-workbook-scale' | 'ui-responsiveness' | 'collaboration'
  label: string
  materializedCells: number
  corpusCaseId: string | null
  metric: string
  actualP95: number
  budgetP95: number
  gateBudgetP95: number
  sampleCount: number
  passed: boolean
  gatePassed: boolean
}

export interface LargeWorkbookSloScorecard {
  schemaVersion: 1
  suite: 'large-workbook-slo'
  summary: {
    coveredLargeWorkbookRows: number[]
    allSloBudgetsPassed: boolean
    allGateBudgetsPassed: boolean
    headedBrowserFrameP95Evidence: 'not-captured'
    externalGoogleSheetsEvidence: 'not-captured'
    externalMicrosoftExcelEvidence: 'not-captured'
  }
  measurements: LargeWorkbookSloMeasurement[]
}

export interface HyperFormulaSurfaceSnapshot {
  hyperFormulaCommit: string
  hyperFormulaVersion: string
  classSurface: {
    staticMembers: string[]
    staticMethods: string[]
    instanceAccessors: string[]
    instanceMethods: string[]
  }
  configKeys: string[]
}

export interface CompetitiveScorecardSummary {
  comparableCount: number
  directionalMeanRatioGeomean: number
  directionalP95RatioGeomean: number
  hyperformulaWins: number
  worstMeanRatioWorkload: string
  worstP95RatioWorkload: string
  worstWorkpaperToHyperFormulaMeanRatio: number
  worstWorkpaperToHyperFormulaP95Ratio: number
  workpaperWins: number
}

export interface CompetitiveFamilySummary {
  comparableCount: number
  family: string
  hyperformulaWins: number
  scorecardEligible: boolean
  workpaperWins: number
  worstMeanRatioWorkload: string | null
  worstP95RatioWorkload: string | null
  worstWorkpaperToHyperFormulaMeanRatio: number | null
  worstWorkpaperToHyperFormulaP95Ratio: number | null
}

export interface CompetitiveResult {
  comparable: boolean
  workload: string
  comparison?: {
    workpaperToHyperFormulaMeanRatio: number
    workpaperToHyperFormulaP95Ratio: number
  }
}

export interface CompetitiveArtifact {
  generatedAt: string
  engines: {
    hyperformula: {
      commit: string
      version: string
    }
  }
  families: CompetitiveFamilySummary[]
  results: CompetitiveResult[]
  scorecard: CompetitiveScorecardSummary
}

export interface BiligDominanceScorecard {
  schemaVersion: 1
  objective: string
  goalStatus: 'active-not-achieved'
  claimPolicy: {
    blanketTenXClaimAllowed: false
    requiredForBlanketTenXClaim: string[]
    workloadSpecificTenXWins: Array<{
      workload: string
      meanRatio: number
      p95Ratio: number
      comparisonTarget: 'HyperFormula'
    }>
  }
  sourceArtifacts: {
    auditabilityScorecard: string
    formulaDominanceSnapshot: string
    hyperFormulaSurfaceSnapshot: string
    importExportFidelityScorecard: string
    largeWorkbookSloScorecard: string
    securityPostureScorecard: string
    workpaperCompetitiveBenchmark: {
      generatedAt: string
      hyperFormulaCommit: string
      hyperFormulaVersion: string
      path: string
    }
  }
  summary: {
    auditabilityCoveredControls: string[]
    auditabilityPosturePassed: boolean
    auditabilityUncoveredControls: string[]
    externalGoogleSheetsEvidence: 'not-captured-in-repo'
    externalMicrosoftExcelEvidence: 'not-captured-in-repo'
    formulaCanonicalProductionPercent: number
    formulaOfficeListedBreadthPercent: number
    formulaTrackedBreadthPercent: number
    importExportCoveredFeatures: string[]
    importExportFidelityPassed: boolean
    importExportUnsupportedFeatures: string[]
    largeWorkbookSloRowsCovered: number[]
    largeWorkbookSloPassed: boolean
    securityCoveredControls: string[]
    securityPosturePassed: boolean
    securityUncoveredControls: string[]
    hyperFormulaComparableWorkloads: number
    hyperFormulaP95GeomeanRatio: number
    hyperFormulaMeanGeomeanRatio: number
    hyperFormulaWins: number
    tenXMeanAndP95WorkloadCountAgainstHyperFormula: number
    workpaperWins: number
  }
  categories: DominanceCategory[]
}

export interface DominanceCategory {
  id: string
  title: string
  objectiveCategory: string
  target: string
  status: DominanceStatus
  currentEvidence: string[]
  evidenceArtifacts: string[]
  checkCommands: string[]
  blockers: string[]
}

export interface BuildScorecardInput {
  auditabilityScorecard: AuditabilityScorecard
  auditabilityScorecardPath: string
  competitiveArtifact: CompetitiveArtifact
  competitiveArtifactPath: string
  formulaSnapshot: FormulaDominanceSnapshot
  formulaSnapshotPath: string
  importExportFidelityScorecard: ImportExportFidelityScorecard
  importExportFidelityScorecardPath: string
  largeWorkbookSloScorecard: LargeWorkbookSloScorecard
  largeWorkbookSloScorecardPath: string
  securityPostureScorecard: SecurityPostureScorecard
  securityPostureScorecardPath: string
  surfaceSnapshot: HyperFormulaSurfaceSnapshot
  surfaceSnapshotPath: string
}

const TEN_X_RATIO = 0.1
const rootDir = resolve(new URL('..', import.meta.url).pathname)
const auditabilityScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'auditability-scorecard.json')
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'bilig-dominance-scorecard.json')
const competitiveArtifactPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-hyperformula.json')
const formulaSnapshotPath = join(rootDir, 'packages', 'formula', 'src', '__tests__', 'fixtures', 'formula-dominance-snapshot.json')
const importExportFidelityScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'import-export-fidelity-scorecard.json')
const largeWorkbookSloScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'large-workbook-slo-scorecard.json')
const securityPostureScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'security-posture-scorecard.json')
const surfaceSnapshotPath = join(rootDir, 'packages', 'headless', 'src', '__tests__', 'fixtures', 'hyperformula-surface.json')

function main(): void {
  const isCheckMode = process.argv.includes('--check')
  const scorecard = buildBiligDominanceScorecard({
    auditabilityScorecard: parseAuditabilityScorecard(readJsonObject(auditabilityScorecardPath)),
    auditabilityScorecardPath: toRepoPath(auditabilityScorecardPath),
    competitiveArtifact: parseCompetitiveArtifact(readJsonObject(competitiveArtifactPath)),
    competitiveArtifactPath: toRepoPath(competitiveArtifactPath),
    formulaSnapshot: parseFormulaDominanceSnapshot(readJsonObject(formulaSnapshotPath)),
    formulaSnapshotPath: toRepoPath(formulaSnapshotPath),
    importExportFidelityScorecard: parseImportExportFidelityScorecard(readJsonObject(importExportFidelityScorecardPath)),
    importExportFidelityScorecardPath: toRepoPath(importExportFidelityScorecardPath),
    largeWorkbookSloScorecard: parseLargeWorkbookSloScorecard(readJsonObject(largeWorkbookSloScorecardPath)),
    largeWorkbookSloScorecardPath: toRepoPath(largeWorkbookSloScorecardPath),
    securityPostureScorecard: parseSecurityPostureScorecard(readJsonObject(securityPostureScorecardPath)),
    securityPostureScorecardPath: toRepoPath(securityPostureScorecardPath),
    surfaceSnapshot: parseSurfaceSnapshot(readJsonObject(surfaceSnapshotPath)),
    surfaceSnapshotPath: toRepoPath(surfaceSnapshotPath),
  })
  const serializedScorecard = formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`)

  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(`Missing generated bilig dominance scorecard at ${outputPath}. Run: bun scripts/gen-bilig-dominance-scorecard.ts`)
    }
    const currentScorecard = readFileSync(outputPath, 'utf8')
    if (currentScorecard !== serializedScorecard) {
      throw new Error('Generated bilig dominance scorecard is out of date. Run: bun scripts/gen-bilig-dominance-scorecard.ts')
    }
  } else {
    mkdirSync(dirname(outputPath), { recursive: true })
    writeFileSync(outputPath, serializedScorecard)
  }

  console.log(
    JSON.stringify(
      {
        mode: isCheckMode ? 'check' : 'write',
        outputPath,
        goalStatus: scorecard.goalStatus,
        blanketTenXClaimAllowed: scorecard.claimPolicy.blanketTenXClaimAllowed,
        workpaperWins: scorecard.summary.workpaperWins,
        hyperFormulaWins: scorecard.summary.hyperFormulaWins,
        tenXMeanAndP95WorkloadCountAgainstHyperFormula: scorecard.summary.tenXMeanAndP95WorkloadCountAgainstHyperFormula,
      },
      null,
      2,
    ),
  )
}

export function buildBiligDominanceScorecard(input: BuildScorecardInput): BiligDominanceScorecard {
  const tenXWorkloads = input.competitiveArtifact.results
    .filter(isComparableWithComparison)
    .filter(
      (result) =>
        result.comparison.workpaperToHyperFormulaMeanRatio <= TEN_X_RATIO &&
        result.comparison.workpaperToHyperFormulaP95Ratio <= TEN_X_RATIO,
    )
    .map((result) => ({
      workload: result.workload,
      meanRatio: result.comparison.workpaperToHyperFormulaMeanRatio,
      p95Ratio: result.comparison.workpaperToHyperFormulaP95Ratio,
      comparisonTarget: 'HyperFormula' as const,
    }))

  const structuralRows = requiredFamily(input.competitiveArtifact, 'structural-rows')
  const structuralColumns = requiredFamily(input.competitiveArtifact, 'structural-columns')
  const dirtyExecution = requiredFamily(input.competitiveArtifact, 'dirty-execution')
  const batchEdit = requiredFamily(input.competitiveArtifact, 'batch-edit')
  const rebuild = requiredFamily(input.competitiveArtifact, 'rebuild')
  const rangeRead = requiredFamily(input.competitiveArtifact, 'range-read')
  const load100k = requiredSloMeasurement(input.largeWorkbookSloScorecard, 'load100k')
  const load250k = requiredSloMeasurement(input.largeWorkbookSloScorecard, 'load250k')
  const workerWarmStart100k = requiredSloMeasurement(input.largeWorkbookSloScorecard, 'workerWarmStart100k')
  const workerWarmStart250k = requiredSloMeasurement(input.largeWorkbookSloScorecard, 'workerWarmStart250k')
  const workerVisibleEdit10k = requiredSloMeasurement(input.largeWorkbookSloScorecard, 'workerVisibleEdit10k')
  const workerReconnectCatchUp100Pending = requiredSloMeasurement(input.largeWorkbookSloScorecard, 'workerReconnectCatchUp100Pending')
  const totalSurfaceMembers =
    input.surfaceSnapshot.classSurface.staticMembers.length +
    input.surfaceSnapshot.classSurface.staticMethods.length +
    input.surfaceSnapshot.classSurface.instanceAccessors.length +
    input.surfaceSnapshot.classSurface.instanceMethods.length

  return {
    schemaVersion: 1,
    objective:
      'Make bilig decisively better than Google Sheets and Microsoft Excel, targeting at least 10x superiority across major spreadsheet/workbook categories.',
    goalStatus: 'active-not-achieved',
    claimPolicy: {
      blanketTenXClaimAllowed: false,
      requiredForBlanketTenXClaim: [
        'direct Google Sheets evidence for every objective category',
        'direct Microsoft Excel evidence for every objective category',
        '10x mean and p95 wins for every performance category',
        'Excel-compatible correctness coverage across the committed formula and workbook semantics surface',
        'browser UI responsiveness artifacts for large collaborative workbooks',
        'security and reliability artifacts that cover production deployment behavior',
      ],
      workloadSpecificTenXWins: tenXWorkloads,
    },
    sourceArtifacts: {
      auditabilityScorecard: input.auditabilityScorecardPath,
      formulaDominanceSnapshot: input.formulaSnapshotPath,
      hyperFormulaSurfaceSnapshot: input.surfaceSnapshotPath,
      importExportFidelityScorecard: input.importExportFidelityScorecardPath,
      largeWorkbookSloScorecard: input.largeWorkbookSloScorecardPath,
      securityPostureScorecard: input.securityPostureScorecardPath,
      workpaperCompetitiveBenchmark: {
        path: input.competitiveArtifactPath,
        generatedAt: input.competitiveArtifact.generatedAt,
        hyperFormulaVersion: input.competitiveArtifact.engines.hyperformula.version,
        hyperFormulaCommit: input.competitiveArtifact.engines.hyperformula.commit,
      },
    },
    summary: {
      auditabilityCoveredControls: input.auditabilityScorecard.summary.coveredControls,
      auditabilityPosturePassed: input.auditabilityScorecard.summary.allRequiredControlsPassed,
      auditabilityUncoveredControls: input.auditabilityScorecard.summary.uncoveredControls,
      externalGoogleSheetsEvidence: 'not-captured-in-repo',
      externalMicrosoftExcelEvidence: 'not-captured-in-repo',
      formulaCanonicalProductionPercent: input.formulaSnapshot.canonical.summary.percent,
      formulaOfficeListedBreadthPercent: input.formulaSnapshot.formulaBreadth.officeListed.percent,
      formulaTrackedBreadthPercent: input.formulaSnapshot.formulaBreadth.tracked.percent,
      importExportCoveredFeatures: input.importExportFidelityScorecard.summary.coveredFeatures,
      importExportFidelityPassed: input.importExportFidelityScorecard.summary.allRequiredCasesPassed,
      importExportUnsupportedFeatures: input.importExportFidelityScorecard.summary.unsupportedFeatures,
      largeWorkbookSloRowsCovered: input.largeWorkbookSloScorecard.summary.coveredLargeWorkbookRows,
      largeWorkbookSloPassed: input.largeWorkbookSloScorecard.summary.allSloBudgetsPassed,
      securityCoveredControls: input.securityPostureScorecard.summary.coveredControls,
      securityPosturePassed: input.securityPostureScorecard.summary.allRequiredControlsPassed,
      securityUncoveredControls: input.securityPostureScorecard.summary.uncoveredControls,
      hyperFormulaComparableWorkloads: input.competitiveArtifact.scorecard.comparableCount,
      hyperFormulaMeanGeomeanRatio: input.competitiveArtifact.scorecard.directionalMeanRatioGeomean,
      hyperFormulaP95GeomeanRatio: input.competitiveArtifact.scorecard.directionalP95RatioGeomean,
      hyperFormulaWins: input.competitiveArtifact.scorecard.hyperformulaWins,
      tenXMeanAndP95WorkloadCountAgainstHyperFormula: tenXWorkloads.length,
      workpaperWins: input.competitiveArtifact.scorecard.workpaperWins,
    },
    categories: [
      {
        id: 'calculation-correctness',
        title: 'Calculation Correctness',
        objectiveCategory: 'calculation correctness',
        target: 'Excel-compatible semantics on the supported workbook and formula surface, with oracle-backed production routing.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          `canonical formula closure is ${formatRatio(input.formulaSnapshot.canonical.summary)}`,
          `Office-listed formula breadth is ${formatRatio(input.formulaSnapshot.formulaBreadth.officeListed)}`,
          `tracked formula breadth is ${formatRatio(input.formulaSnapshot.formulaBreadth.tracked)}`,
          `strategic canonical rows are production-routed; open canonical rows: ${input.formulaSnapshot.canonical.nonProductionRows.length}`,
        ],
        evidenceArtifacts: [input.formulaSnapshotPath, 'docs/excel-parity-program.md', 'docs/formula-oracle-capture.md'],
        checkCommands: ['pnpm formula:dominance:check', 'pnpm test:correctness:formula'],
        blockers: [
          `${input.formulaSnapshot.formulaBreadth.missingOfficeFunctions.length} Office-listed functions are still missing from the runtime inventory`,
          'no generated scorecard currently compares all committed semantics directly against live Google Sheets and Microsoft Excel',
        ],
      },
      {
        id: 'recalculation-speed',
        title: 'Recalculation Speed',
        objectiveCategory: 'recalculation speed',
        target: '10x mean and p95 wins on named recalculation workloads against each comparison target.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          familyWinSummary(dirtyExecution),
          familyWinSummary(batchEdit),
          familyWinSummary(rebuild),
          `overall HyperFormula comparable scorecard is ${input.competitiveArtifact.scorecard.workpaperWins}/${input.competitiveArtifact.scorecard.comparableCount}`,
        ],
        evidenceArtifacts: [input.competitiveArtifactPath],
        checkCommands: ['pnpm workpaper:bench:competitive:check', 'pnpm bench:contracts'],
        blockers: [
          'current checked-in benchmark is against HyperFormula, not Google Sheets or Microsoft Excel',
          `only ${tenXWorkloads.length} comparable HyperFormula workloads are 10x wins on both mean and p95`,
        ],
      },
      {
        id: 'structural-edit-performance',
        title: 'Structural Edit Performance',
        objectiveCategory: 'structural-edit performance',
        target: '10x mean and p95 wins for insert/delete/move rows and columns at workbook scale.',
        status: 'partial-repo-evidence',
        currentEvidence: [familyWinSummary(structuralRows), familyWinSummary(structuralColumns)],
        evidenceArtifacts: [input.competitiveArtifactPath],
        checkCommands: ['pnpm workpaper:bench:competitive:check'],
        blockers: [
          'structural rows and columns lead HyperFormula, but the worst ratios are not 10x',
          'no direct Sheets or Excel structural-edit timing artifact exists in the repo',
        ],
      },
      {
        id: 'large-workbook-scale',
        title: 'Large Workbook Scale',
        objectiveCategory: 'large-workbook scale',
        target: 'Sub-second warm start, import, viewport, paste, sort, and filter behavior on 100k to 250k row workbooks.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          'local-first worker architecture and OPFS/SQLite model are documented',
          'range-read and build families have HyperFormula comparison evidence',
          familyWinSummary(rangeRead),
          `large-workbook SLO artifact covers ${input.largeWorkbookSloScorecard.summary.coveredLargeWorkbookRows.join(', ')} materialized-cell sessions`,
          sloSummary(load100k),
          sloSummary(load250k),
          sloSummary(workerWarmStart100k),
          sloSummary(workerWarmStart250k),
        ],
        evidenceArtifacts: [input.competitiveArtifactPath, input.largeWorkbookSloScorecardPath, 'docs/05-06-next-phase.md'],
        checkCommands: ['pnpm large-workbook:slo:check', 'CI=1 pnpm bench:contracts', 'pnpm test:browser:full', 'pnpm bench:smoke'],
        blockers: [
          'the generated SLO artifact covers core load and browser-worker runtime, but not headed browser frame/render p95 for 100k and 250k sessions',
          'no direct Sheets or Excel large-workbook scale artifact exists in the repo',
        ],
      },
      {
        id: 'ui-responsiveness',
        title: 'UI Responsiveness',
        objectiveCategory: 'UI responsiveness',
        target: 'Local visible response p95 below 16ms, selection paint p95 below 8ms, and stable viewport frame pacing.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          'worker-first local runtime, projected viewport patches, and tile cache architecture are implemented and documented',
          'browser correctness and performance smoke commands exist',
          sloSummary(workerVisibleEdit10k),
        ],
        evidenceArtifacts: [
          input.largeWorkbookSloScorecardPath,
          'docs/05-06-next-phase.md',
          'apps/web/src/perf/workbook-perf.ts',
          'apps/web/src/perf/workbook-scroll-perf.ts',
        ],
        checkCommands: ['pnpm large-workbook:slo:check', 'CI=1 pnpm bench:contracts', 'pnpm test:browser:full', 'pnpm bench:smoke'],
        blockers: [
          'worker visible-patch p95 is checked, but no headed browser frame/paint p95 artifact proves the stated UI latency targets',
          'no direct Sheets or Excel browser responsiveness comparison artifact exists in the repo',
        ],
      },
      {
        id: 'collaboration',
        title: 'Collaboration',
        objectiveCategory: 'collaboration',
        target: 'Local-first collaboration with durable pending ops, safe rebase, explicit conflicts, private views, and change review.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          'worker runtime, pending-op journal, authoritative delta ingest, presence, and changes-pane tests exist',
          'collaboration product targets are documented',
          sloSummary(workerReconnectCatchUp100Pending),
        ],
        evidenceArtifacts: [
          input.largeWorkbookSloScorecardPath,
          'docs/05-06-next-phase.md',
          'apps/web/src/__tests__/workbook-sync.test.ts',
          'apps/web/src/__tests__/workbook-presence.test.tsx',
        ],
        checkCommands: [
          'pnpm large-workbook:slo:check',
          'CI=1 pnpm bench:contracts',
          'pnpm test:correctness:browser',
          'pnpm test:correctness:server',
        ],
        blockers: [
          'generated collaboration SLO evidence covers reconnect catch-up only; conflict rate and multi-user viewport behavior remain uncovered',
          'no direct Sheets collaboration comparison artifact exists in the repo',
        ],
      },
      {
        id: 'automation-api-extensibility',
        title: 'Automation And API Extensibility',
        objectiveCategory: 'automation/API extensibility',
        target: 'Typed semantic workbook operations and agent tools that replace brittle DOM automation.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          `${totalSurfaceMembers} HyperFormula surface members are snapshotted for parity tracking`,
          `${input.surfaceSnapshot.configKeys.length} HyperFormula config keys are snapshotted for parity tracking`,
          'WorkPaper exposes additional detailed events and performance counters',
        ],
        evidenceArtifacts: [input.surfaceSnapshotPath, 'packages/headless/src/__tests__/hyperformula-surface-parity.test.ts'],
        checkCommands: ['pnpm workpaper:parity:check', 'pnpm workpaper:smoke:external'],
        blockers: [
          'no generated Google Apps Script or Office Scripts capability comparison exists',
          'no 10x workflow automation benchmark exists for semantic operations versus incumbent scripting',
        ],
      },
      {
        id: 'import-export-compatibility',
        title: 'Import And Export Compatibility',
        objectiveCategory: 'import/export compatibility',
        target: 'High-fidelity CSV/XLSX import, preview, authoritative finalization, and compatibility reporting.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          `generated import/export fidelity scorecard passes required cases: ${String(input.importExportFidelityScorecard.summary.allRequiredCasesPassed)}`,
          `covered import/export features: ${input.importExportFidelityScorecard.summary.coveredFeatures.join(', ')}`,
          `unsupported XLSX features are explicitly disclosed: ${input.importExportFidelityScorecard.summary.unsupportedFeatures.join(', ')}`,
        ],
        evidenceArtifacts: [
          input.importExportFidelityScorecardPath,
          'packages/excel-import/src/__tests__/excel-import.test.ts',
          'packages/core/src/__tests__/engine-import-export.fuzz.test.ts',
          'docs/formula-oracle-capture.md',
        ],
        checkCommands: [
          'pnpm import-export:fidelity:check',
          'pnpm exec vitest run packages/excel-import/src/__tests__/excel-import.test.ts packages/core/src/__tests__/engine-import-export.fuzz.test.ts',
        ],
        blockers: [
          'generated XLSX round-trip evidence covers supported snapshot semantics, not full native Excel object-model round trips for styles, charts, pivots, macros, comments, and defined names',
          'no direct Sheets import/export compatibility artifact exists in the repo',
          'no direct Microsoft Excel import/export compatibility artifact exists in the repo',
        ],
      },
      {
        id: 'auditability',
        title: 'Auditability',
        objectiveCategory: 'auditability',
        target: 'Every AI or user workflow is reviewable, previewable, undoable, and tied to durable change evidence.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          `generated auditability scorecard passes required controls: ${String(input.auditabilityScorecard.summary.allRequiredControlsPassed)}`,
          `covered auditability controls: ${input.auditabilityScorecard.summary.coveredControls.join(', ')}`,
          `uncovered auditability controls are explicitly disclosed: ${input.auditabilityScorecard.summary.uncoveredControls.join(', ')}`,
          'change bundles, versions, revertable changes, and agent preview/apply rails are documented',
          'workbook changes and mutation journal tests exist',
        ],
        evidenceArtifacts: [
          input.auditabilityScorecardPath,
          'docs/05-06-next-phase.md',
          'apps/web/src/__tests__/workbook-changes.test.tsx',
          'apps/web/src/__tests__/worker-runtime-mutation-journal.test.ts',
        ],
        checkCommands: [
          'pnpm auditability:check',
          'pnpm exec vitest run packages/agent-api/src/__tests__/workbook-agent-preview.test.ts apps/bilig/src/zero/__tests__/workbook-agent-apply.test.ts packages/zero-sync/src/__tests__/workbook-history-state.test.ts',
          'pnpm test:correctness:browser',
          'pnpm test:correctness:server',
        ],
        blockers: [
          'generated auditability evidence covers preview/apply parity, undo bundle capture, authoritative stale/mismatch guards, and history revert/redo state, but not headed browser preview/apply/revert flow',
          'no direct incumbent auditability comparison artifact exists in the repo',
        ],
      },
      {
        id: 'reliability',
        title: 'Reliability',
        objectiveCategory: 'reliability',
        target: 'Crash-safe local durability, deterministic replay, durable sync, and no accepted-op loss.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          'local pending-op journal and reconnect/rebase architecture are documented',
          'runtime sync replay, fuzz, reconnect, and local persistence tests exist',
        ],
        evidenceArtifacts: ['docs/05-06-next-phase.md', 'apps/web/src/__tests__/runtime-sync.fuzz.test.ts'],
        checkCommands: ['pnpm test:fuzz:main', 'pnpm test:correctness:browser', 'pnpm test:correctness:server'],
        blockers: [
          'no generated crash/reload/offline soak artifact proves zero pending-op loss',
          'no direct Sheets or Excel reliability comparison artifact exists in the repo',
        ],
      },
      {
        id: 'security',
        title: 'Security',
        objectiveCategory: 'security',
        target: 'Safe workbook execution, no hidden runtime requirements, controlled agent permissions, and auditable deployment posture.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          `generated security posture scorecard passes required controls: ${String(input.securityPostureScorecard.summary.allRequiredControlsPassed)}`,
          `covered security controls: ${input.securityPostureScorecard.summary.coveredControls.join(', ')}`,
          `uncovered security controls are explicitly disclosed: ${input.securityPostureScorecard.summary.uncoveredControls.join(', ')}`,
        ],
        evidenceArtifacts: [
          input.securityPostureScorecardPath,
          'packages/excel-import/src/__tests__/excel-import.test.ts',
          'packages/agent-api/src/__tests__/workbook-agent-execution-policy.test.ts',
          'scripts/check-package-publish.ts',
        ],
        checkCommands: [
          'pnpm security:posture:check',
          'pnpm exec vitest run packages/excel-import/src/__tests__/excel-import.test.ts packages/agent-api/src/__tests__/workbook-agent-execution-policy.test.ts',
          'pnpm publish:runtime:check',
        ],
        blockers: [
          'generated security posture evidence covers formula dynamic-code scanning, XLSX macro warning, shared-agent owner review, and runtime package hardening, but not browser CSP, dependency vulnerability audit, or deployment network policy',
          'no direct Sheets or Excel security comparison artifact exists in the repo',
        ],
      },
      {
        id: 'operator-developer-workflow',
        title: 'Operator And Developer Workflow',
        objectiveCategory: 'operator/developer workflow',
        target: 'One-command verification, generated artifacts, external smoke tests, and no unverifiable leadership claims.',
        status: 'repo-proved-lead',
        currentEvidence: [
          'generated parity, formula inventory, formula dominance, workspace resolution, benchmark, publish, and smoke gates exist',
          'this generated scorecard prevents blanket 10x claims from outrunning evidence',
        ],
        evidenceArtifacts: ['package.json', input.competitiveArtifactPath, input.formulaSnapshotPath, input.surfaceSnapshotPath],
        checkCommands: ['pnpm dominance:check', 'pnpm run ci'],
        blockers: [
          'developer workflow evidence is strong, but it does not prove the product is 10x better across every user-facing category',
        ],
      },
    ],
  }
}

function requiredFamily(artifact: CompetitiveArtifact, familyName: string): CompetitiveFamilySummary {
  const family = artifact.families.find((candidate) => candidate.family === familyName)
  if (!family) {
    throw new Error(`Competitive artifact is missing required family: ${familyName}`)
  }
  return family
}

function requiredSloMeasurement(scorecard: LargeWorkbookSloScorecard, measurementId: string): LargeWorkbookSloMeasurement {
  const measurement = scorecard.measurements.find((candidate) => candidate.id === measurementId)
  if (!measurement) {
    throw new Error(`Large workbook SLO scorecard is missing required measurement: ${measurementId}`)
  }
  return measurement
}

function familyWinSummary(family: CompetitiveFamilySummary): string {
  const worstMean =
    family.worstWorkpaperToHyperFormulaMeanRatio === null || family.worstMeanRatioWorkload === null
      ? 'no comparable mean ratio'
      : `worst mean ratio ${family.worstWorkpaperToHyperFormulaMeanRatio} on ${family.worstMeanRatioWorkload}`
  const worstP95 =
    family.worstWorkpaperToHyperFormulaP95Ratio === null || family.worstP95RatioWorkload === null
      ? 'no comparable p95 ratio'
      : `worst p95 ratio ${family.worstWorkpaperToHyperFormulaP95Ratio} on ${family.worstP95RatioWorkload}`
  return `${family.family}: WorkPaper ${family.workpaperWins}/${family.comparableCount}, HyperFormula ${family.hyperformulaWins}/${family.comparableCount}; ${worstMean}; ${worstP95}`
}

function sloSummary(measurement: LargeWorkbookSloMeasurement): string {
  return `${measurement.id}: ${measurement.metric} ${measurement.actualP95}ms against ${measurement.budgetP95}ms SLO (${measurement.sampleCount} samples)`
}

function formatRatio(summary: RatioSummary): string {
  return `${summary.production}/${summary.total} (${summary.percent}%)`
}

function isComparableWithComparison(result: CompetitiveResult): result is CompetitiveResult & {
  comparable: true
  comparison: {
    workpaperToHyperFormulaMeanRatio: number
    workpaperToHyperFormulaP95Ratio: number
  }
} {
  return (
    result.comparable &&
    result.comparison !== undefined &&
    isFiniteNumber(result.comparison.workpaperToHyperFormulaMeanRatio) &&
    isFiniteNumber(result.comparison.workpaperToHyperFormulaP95Ratio)
  )
}

function parseFormulaDominanceSnapshot(value: Record<string, unknown>): FormulaDominanceSnapshot {
  const formulaBreadth = objectField(value, 'formulaBreadth')
  const canonical = objectField(value, 'canonical')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    formulaBreadth: {
      officeListed: ratioField(formulaBreadth, 'officeListed'),
      tracked: ratioField(formulaBreadth, 'tracked'),
      missingOfficeFunctions: stringArrayField(formulaBreadth, 'missingOfficeFunctions'),
    },
    canonical: {
      summary: ratioField(canonical, 'summary'),
      nonProductionRows: arrayField(canonical, 'nonProductionRows'),
    },
  }
}

function parseLargeWorkbookSloScorecard(value: Record<string, unknown>): LargeWorkbookSloScorecard {
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'large-workbook-slo'),
    summary: {
      coveredLargeWorkbookRows: numberArrayField(summary, 'coveredLargeWorkbookRows'),
      allSloBudgetsPassed: booleanField(summary, 'allSloBudgetsPassed'),
      allGateBudgetsPassed: booleanField(summary, 'allGateBudgetsPassed'),
      headedBrowserFrameP95Evidence: literalField(summary, 'headedBrowserFrameP95Evidence', 'not-captured'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'not-captured'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'not-captured'),
    },
    measurements: arrayField(value, 'measurements').map(parseLargeWorkbookSloMeasurement),
  }
}

function parseLargeWorkbookSloMeasurement(value: unknown): LargeWorkbookSloMeasurement {
  const measurement = asObject(value, 'large workbook SLO measurement')
  return {
    id: stringField(measurement, 'id'),
    category: parseSloMeasurementCategory(stringField(measurement, 'category')),
    label: stringField(measurement, 'label'),
    materializedCells: numberField(measurement, 'materializedCells'),
    corpusCaseId: optionalStringField(measurement, 'corpusCaseId'),
    metric: stringField(measurement, 'metric'),
    actualP95: numberField(measurement, 'actualP95'),
    budgetP95: numberField(measurement, 'budgetP95'),
    gateBudgetP95: numberField(measurement, 'gateBudgetP95'),
    sampleCount: numberField(measurement, 'sampleCount'),
    passed: booleanField(measurement, 'passed'),
    gatePassed: booleanField(measurement, 'gatePassed'),
  }
}

function parseSloMeasurementCategory(value: string): LargeWorkbookSloMeasurement['category'] {
  if (value === 'large-workbook-scale' || value === 'ui-responsiveness' || value === 'collaboration') {
    return value
  }
  throw new Error(`Unexpected large workbook SLO measurement category: ${value}`)
}

function parseSurfaceSnapshot(value: Record<string, unknown>): HyperFormulaSurfaceSnapshot {
  const classSurface = objectField(value, 'classSurface')
  return {
    hyperFormulaVersion: stringField(value, 'hyperFormulaVersion'),
    hyperFormulaCommit: stringField(value, 'hyperFormulaCommit'),
    classSurface: {
      staticMembers: stringArrayField(classSurface, 'staticMembers'),
      staticMethods: stringArrayField(classSurface, 'staticMethods'),
      instanceAccessors: stringArrayField(classSurface, 'instanceAccessors'),
      instanceMethods: stringArrayField(classSurface, 'instanceMethods'),
    },
    configKeys: stringArrayField(value, 'configKeys'),
  }
}

function parseCompetitiveArtifact(value: Record<string, unknown>): CompetitiveArtifact {
  const engines = objectField(value, 'engines')
  const hyperformula = objectField(engines, 'hyperformula')
  const scorecard = objectField(value, 'scorecard')
  return {
    generatedAt: stringField(value, 'generatedAt'),
    engines: {
      hyperformula: {
        commit: stringField(hyperformula, 'commit'),
        version: stringField(hyperformula, 'version'),
      },
    },
    families: arrayField(value, 'families').map(parseCompetitiveFamily),
    results: arrayField(value, 'results').map(parseCompetitiveResult),
    scorecard: {
      comparableCount: numberField(scorecard, 'comparableCount'),
      directionalMeanRatioGeomean: numberField(scorecard, 'directionalMeanRatioGeomean'),
      directionalP95RatioGeomean: numberField(scorecard, 'directionalP95RatioGeomean'),
      hyperformulaWins: numberField(scorecard, 'hyperformulaWins'),
      worstMeanRatioWorkload: stringField(scorecard, 'worstMeanRatioWorkload'),
      worstP95RatioWorkload: stringField(scorecard, 'worstP95RatioWorkload'),
      worstWorkpaperToHyperFormulaMeanRatio: numberField(scorecard, 'worstWorkpaperToHyperFormulaMeanRatio'),
      worstWorkpaperToHyperFormulaP95Ratio: numberField(scorecard, 'worstWorkpaperToHyperFormulaP95Ratio'),
      workpaperWins: numberField(scorecard, 'workpaperWins'),
    },
  }
}

function parseCompetitiveFamily(value: unknown): CompetitiveFamilySummary {
  const family = asObject(value, 'competitive family')
  return {
    comparableCount: numberField(family, 'comparableCount'),
    family: stringField(family, 'family'),
    hyperformulaWins: numberField(family, 'hyperformulaWins'),
    scorecardEligible: booleanField(family, 'scorecardEligible'),
    workpaperWins: numberField(family, 'workpaperWins'),
    worstMeanRatioWorkload: optionalStringField(family, 'worstMeanRatioWorkload'),
    worstP95RatioWorkload: optionalStringField(family, 'worstP95RatioWorkload'),
    worstWorkpaperToHyperFormulaMeanRatio: optionalNumberField(family, 'worstWorkpaperToHyperFormulaMeanRatio'),
    worstWorkpaperToHyperFormulaP95Ratio: optionalNumberField(family, 'worstWorkpaperToHyperFormulaP95Ratio'),
  }
}

function parseCompetitiveResult(value: unknown): CompetitiveResult {
  const result = asObject(value, 'competitive result')
  const comparison = result['comparison']
  return {
    comparable: booleanField(result, 'comparable'),
    workload: stringField(result, 'workload'),
    comparison:
      comparison === undefined
        ? undefined
        : {
            workpaperToHyperFormulaMeanRatio: numberField(
              asObject(comparison, 'competitive comparison'),
              'workpaperToHyperFormulaMeanRatio',
            ),
            workpaperToHyperFormulaP95Ratio: numberField(asObject(comparison, 'competitive comparison'), 'workpaperToHyperFormulaP95Ratio'),
          },
  }
}

function readJsonObject(path: string): Record<string, unknown> {
  return asObject(JSON.parse(readFileSync(path, 'utf8')) as unknown, path)
}

function ratioField(value: Record<string, unknown>, field: string): RatioSummary {
  const ratioValue = objectField(value, field)
  return {
    percent: numberField(ratioValue, 'percent'),
    production: numberField(ratioValue, 'production'),
    total: numberField(ratioValue, 'total'),
  }
}

function objectField(value: Record<string, unknown>, field: string): Record<string, unknown> {
  return asObject(value[field], field)
}

function stringField(value: Record<string, unknown>, field: string): string {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${field} to be a string`)
  }
  return fieldValue
}

function optionalStringField(value: Record<string, unknown>, field: string): string | null {
  const fieldValue = value[field]
  if (fieldValue === null) {
    return null
  }
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${field} to be a string or null`)
  }
  return fieldValue
}

function stringArrayField(value: Record<string, unknown>, field: string): string[] {
  const fieldValue = arrayField(value, field)
  if (!fieldValue.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected ${field} to be a string array`)
  }
  return fieldValue
}

function numberArrayField(value: Record<string, unknown>, field: string): number[] {
  const fieldValue = arrayField(value, field)
  if (!fieldValue.every(isFiniteNumber)) {
    throw new Error(`Expected ${field} to be a finite number array`)
  }
  return fieldValue
}

function arrayField(value: Record<string, unknown>, field: string): unknown[] {
  const fieldValue = value[field]
  if (!Array.isArray(fieldValue)) {
    throw new Error(`Expected ${field} to be an array`)
  }
  return fieldValue
}

function numberField(value: Record<string, unknown>, field: string): number {
  const fieldValue = value[field]
  if (!isFiniteNumber(fieldValue)) {
    throw new Error(`Expected ${field} to be a finite number`)
  }
  return fieldValue
}

function optionalNumberField(value: Record<string, unknown>, field: string): number | null {
  const fieldValue = value[field]
  if (fieldValue === null) {
    return null
  }
  if (!isFiniteNumber(fieldValue)) {
    throw new Error(`Expected ${field} to be a finite number or null`)
  }
  return fieldValue
}

function booleanField(value: Record<string, unknown>, field: string): boolean {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${field} to be a boolean`)
  }
  return fieldValue
}

function literalField<const T extends string | number | boolean>(value: Record<string, unknown>, field: string, expected: T): T {
  if (value[field] !== expected) {
    throw new Error(`Expected ${field} to be ${String(expected)}`)
  }
  return expected
}

function asObject(value: unknown, name: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Expected ${name} to be an object`)
  }
  const record: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    record[key] = Reflect.get(value, key)
  }
  return record
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function toRepoPath(path: string): string {
  return path.startsWith(`${rootDir}/`) ? path.slice(rootDir.length + 1) : path
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-dominance-scorecard-'))
  const tempFilePath = join(tempDir, 'scorecard.json')
  writeFileSync(tempFilePath, serializedJson)
  const oxfmtPath = join(rootDir, 'node_modules', '.bin', 'oxfmt')

  const formatResult = Bun.spawnSync([oxfmtPath, '--write', tempFilePath], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (formatResult.exitCode !== 0) {
    rmSync(tempDir, { recursive: true, force: true })
    throw new Error(`Unable to format generated scorecard: ${new TextDecoder().decode(formatResult.stderr).trim()}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
