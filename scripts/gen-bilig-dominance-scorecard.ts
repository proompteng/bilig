#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import type {
  BiligDominanceScorecard,
  BuildScorecardInput,
  CompetitiveArtifact,
  CompetitiveFamilySummary,
  CompetitiveResult,
  LargeWorkbookSloMeasurement,
  LargeWorkbookSloScorecard,
  OverallGoogleSheets10xStatus,
} from './bilig-dominance-scorecard-types.ts'
import {
  familyWinSummary,
  formatList,
  formatRatio,
  formulaMissingFunctionBlockers,
  headedBrowserContractSummary,
  sloSummary,
} from './bilig-dominance-formatters.ts'
import { buildBiligDominanceCompletionAudit, hasUiResponsivenessSameCorpusTenXGap } from './bilig-dominance-completion-audit.ts'
import { loadBiligDominanceScorecardInput, outputPath, rootDir } from './bilig-dominance-scorecard-input.ts'
import { buildBiligDominanceScorecardSummary } from './bilig-dominance-scorecard-summary.ts'
import { loadOperatorWorkflowEvidence, operatorWorkflowGaps } from './bilig-dominance-operator-workflow.ts'
import { isFiniteNumber } from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export type { BiligDominanceScorecard, BuildScorecardInput } from './bilig-dominance-scorecard-types.ts'

const TEN_X_RATIO = 0.1

function main(): void {
  const isCheckMode = process.argv.includes('--check')
  const scorecard = buildBiligDominanceScorecard(loadBiligDominanceScorecardInput())
  const serializedScorecard = formatJsonForRepo({
    rootDir,
    serializedJson: `${JSON.stringify(scorecard, null, 2)}\n`,
    tempPrefix: 'bilig-dominance-scorecard',
  })

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
  const headedBrowserScaleContracts = input.largeWorkbookSloScorecard.headedBrowserFrameP95Contracts.filter(
    (contract) => contract.category === 'large-workbook-scale',
  )
  const headedBrowserUiContracts = input.largeWorkbookSloScorecard.headedBrowserFrameP95Contracts.filter(
    (contract) => contract.category === 'ui-responsiveness',
  )
  const microsoftExcelRecalculationTenXPassed =
    input.microsoftExcelLiveRecalculationScorecard.summary.allRequiredCasesPassed &&
    input.microsoftExcelLiveRecalculationScorecard.summary.tenXMeanAndP95CaseCount ===
      input.microsoftExcelLiveRecalculationScorecard.summary.requiredCaseCount
  const microsoftExcelRecalculationPassedCaseCount = input.microsoftExcelLiveRecalculationScorecard.cases.filter(
    (entry) => entry.passed,
  ).length
  const googleSheetsRecalculationTenXPassed =
    input.googleSheetsLiveRecalculationScorecard.summary.allRequiredCasesPassed &&
    input.googleSheetsLiveRecalculationScorecard.summary.tenXMeanAndP95CaseCount ===
      input.googleSheetsLiveRecalculationScorecard.summary.requiredCaseCount
  const googleSheetsRecalculationPassedCaseCount = input.googleSheetsLiveRecalculationScorecard.cases.filter((entry) => entry.passed).length
  const microsoftExcelLargeWorkbookTenXPassed =
    input.microsoftExcelLiveLargeWorkbookScorecard.summary.allRequiredCasesPassed &&
    input.microsoftExcelLiveLargeWorkbookScorecard.summary.tenXMeanAndP95CaseCount ===
      input.microsoftExcelLiveLargeWorkbookScorecard.summary.requiredCaseCount
  const microsoftExcelLargeWorkbookPassedCaseCount = input.microsoftExcelLiveLargeWorkbookScorecard.cases.filter(
    (entry) => entry.passed,
  ).length
  const googleSheetsLargeWorkbookTenXPassed =
    input.googleSheetsLiveLargeWorkbookScorecard.summary.allRequiredCasesPassed &&
    input.googleSheetsLiveLargeWorkbookScorecard.summary.tenXMeanAndP95CaseCount ===
      input.googleSheetsLiveLargeWorkbookScorecard.summary.requiredCaseCount
  const googleSheetsLargeWorkbookPassedCaseCount = input.googleSheetsLiveLargeWorkbookScorecard.cases.filter((entry) => entry.passed).length
  const microsoftExcelStructuralTenXPassed =
    input.microsoftExcelLiveStructuralScorecard.summary.allRequiredCasesPassed &&
    input.microsoftExcelLiveStructuralScorecard.summary.tenXMeanAndP95CaseCount ===
      input.microsoftExcelLiveStructuralScorecard.summary.requiredCaseCount
  const microsoftExcelStructuralPassedCaseCount = input.microsoftExcelLiveStructuralScorecard.cases.filter((entry) => entry.passed).length
  const googleSheetsStructuralTenXPassed =
    input.googleSheetsLiveStructuralScorecard.summary.allRequiredCasesPassed &&
    input.googleSheetsLiveStructuralScorecard.summary.tenXMeanAndP95CaseCount ===
      input.googleSheetsLiveStructuralScorecard.summary.requiredCaseCount
  const googleSheetsStructuralPassedCaseCount = input.googleSheetsLiveStructuralScorecard.cases.filter((entry) => entry.passed).length
  const recalculationDirectTargetsTenXPassed = microsoftExcelRecalculationTenXPassed && googleSheetsRecalculationTenXPassed
  const structuralDirectTargetsTenXPassed = microsoftExcelStructuralTenXPassed && googleSheetsStructuralTenXPassed
  const largeWorkbookDirectTargetsTenXPassed = microsoftExcelLargeWorkbookTenXPassed && googleSheetsLargeWorkbookTenXPassed
  const calculationSemanticsPassed = input.calculationSemanticsScorecard.summary.allCommittedFormulaSemanticsCovered
  const uiResponsivenessLiveBrowserPassed = input.uiResponsivenessLiveBrowserScorecard.summary.allRequiredCasesPassed
  const totalSurfaceMembers =
    input.surfaceSnapshot.classSurface.staticMembers.length +
    input.surfaceSnapshot.classSurface.staticMethods.length +
    input.surfaceSnapshot.classSurface.instanceAccessors.length +
    input.surfaceSnapshot.classSurface.instanceMethods.length
  const securityUncoveredControls = new Set(input.securityPostureScorecard.summary.uncoveredControls)
  const uiSameCorpusTenXGap = hasUiResponsivenessSameCorpusTenXGap(input.uiResponsivenessLiveBrowserScorecard)
  const operatorWorkflowEvidence = loadOperatorWorkflowEvidence(rootDir)
  const operatorWorkflowBlockers = operatorWorkflowGaps(operatorWorkflowEvidence)
  const completionAudit = buildBiligDominanceCompletionAudit(input, {
    calculationSemanticsPassed,
    googleSheetsLargeWorkbookTenXPassed,
    googleSheetsRecalculationTenXPassed,
    googleSheetsStructuralTenXPassed,
    largeWorkbookDirectTargetsTenXPassed,
    microsoftExcelLargeWorkbookTenXPassed,
    microsoftExcelRecalculationTenXPassed,
    microsoftExcelStructuralTenXPassed,
    recalculationDirectTargetsTenXPassed,
    structuralDirectTargetsTenXPassed,
    uiResponsivenessLiveBrowserPassed,
  })
  const overallGoogleSheets10xStatus = buildOverallGoogleSheets10xStatus(input, {
    googleSheetsLargeWorkbookTenXPassed,
    googleSheetsRecalculationTenXPassed,
    googleSheetsStructuralTenXPassed,
    uiResponsivenessLiveBrowserPassed,
    uiSameCorpusTenXGap,
  })

  return {
    schemaVersion: 1,
    objective:
      'Make bilig decisively better than Google Sheets and Microsoft Excel, targeting at least 10x superiority across major spreadsheet/workbook categories.',
    goalStatus: completionAudit.allCriteriaPassed ? 'achieved' : 'active-not-achieved',
    overallGoogleSheets10xStatus,
    claimPolicy: {
      blanketTenXClaimAllowed: completionAudit.allCriteriaPassed && overallGoogleSheets10xStatus.passed,
      requiredForBlanketTenXClaim: completionAudit.criteria.map((entry) => entry.requirement),
      unmetRequirements: completionAudit.unmetRequirements,
      workloadSpecificTenXWins: tenXWorkloads,
    },
    completionAudit,
    sourceArtifacts: {
      auditabilityScorecard: input.auditabilityScorecardPath,
      automationScorecard: input.automationScorecardPath,
      calculationSemanticsScorecard: input.calculationSemanticsScorecardPath,
      collaborationScorecard: input.collaborationScorecardPath,
      formulaDominanceSnapshot: input.formulaSnapshotPath,
      googleSheetsLiveCalculationScorecard: input.googleSheetsLiveCalculationScorecardPath,
      googleSheetsLiveRecalculationScorecard: input.googleSheetsLiveRecalculationScorecardPath,
      googleSheetsLiveStructuralScorecard: input.googleSheetsLiveStructuralScorecardPath,
      googleSheetsLiveLargeWorkbookScorecard: input.googleSheetsLiveLargeWorkbookScorecardPath,
      hyperFormulaSurfaceSnapshot: input.surfaceSnapshotPath,
      microsoftExcelLiveCalculationScorecard: input.microsoftExcelLiveCalculationScorecardPath,
      microsoftExcelLiveRecalculationScorecard: input.microsoftExcelLiveRecalculationScorecardPath,
      microsoftExcelLiveLargeWorkbookScorecard: input.microsoftExcelLiveLargeWorkbookScorecardPath,
      microsoftExcelLiveStructuralScorecard: input.microsoftExcelLiveStructuralScorecardPath,
      importExportFidelityScorecard: input.importExportFidelityScorecardPath,
      largeWorkbookSloScorecard: input.largeWorkbookSloScorecardPath,
      uiResponsivenessLiveBrowserScorecard: input.uiResponsivenessLiveBrowserScorecardPath,
      reliabilityScorecard: input.reliabilityScorecardPath,
      securityPostureScorecard: input.securityPostureScorecardPath,
      workpaperCompetitiveBenchmark: {
        path: input.competitiveArtifactPath,
        generatedAt: input.competitiveArtifact.generatedAt,
        hyperFormulaVersion: input.competitiveArtifact.engines.hyperformula.version,
        hyperFormulaCommit: input.competitiveArtifact.engines.hyperformula.commit,
      },
    },
    summary: buildBiligDominanceScorecardSummary(input, {
      calculationSemanticsPassed,
      tenXMeanAndP95WorkloadCountAgainstHyperFormula: tenXWorkloads.length,
      uiResponsivenessLiveBrowserPassed,
    }),
    categories: [
      {
        id: 'calculation-correctness',
        title: 'Calculation Correctness',
        objectiveCategory: 'calculation correctness',
        target: 'Excel-compatible semantics on the supported workbook and formula surface, with oracle-backed production routing.',
        status:
          calculationSemanticsPassed &&
          input.formulaSnapshot.formulaBreadth.missingOfficeFunctions.length === 0 &&
          input.microsoftExcelLiveCalculationScorecard.summary.allRequiredCasesPassed &&
          input.googleSheetsLiveCalculationScorecard.summary.allRequiredCasesPassed
            ? 'repo-proved-lead'
            : 'partial-repo-evidence',
        currentEvidence: [
          `canonical formula closure is ${formatRatio(input.formulaSnapshot.canonical.summary)}`,
          `Office-listed formula breadth is ${formatRatio(input.formulaSnapshot.formulaBreadth.officeListed)}`,
          `tracked formula breadth is ${formatRatio(input.formulaSnapshot.formulaBreadth.tracked)}`,
          `strategic canonical rows are production-routed; open canonical rows: ${input.formulaSnapshot.canonical.nonProductionRows.length}`,
          `committed calculation semantics scorecard covers ${String(
            input.calculationSemanticsScorecard.summary.coveredCanonicalFixtureCount,
          )}/${String(input.calculationSemanticsScorecard.summary.canonicalFormulaFixtureCount)} canonical fixtures and ${String(
            input.calculationSemanticsScorecard.summary.coveredWorkbookSemanticsFixtureCount,
          )}/${String(input.calculationSemanticsScorecard.summary.workbookSemanticsFixtureCount)} workbook-semantics fixtures`,
          `live Microsoft Excel calculation scorecard passes ${String(
            input.microsoftExcelLiveCalculationScorecard.summary.matchingCaseCount,
          )}/${String(input.microsoftExcelLiveCalculationScorecard.summary.requiredCaseCount)} required cases on Excel ${
            input.microsoftExcelLiveCalculationScorecard.microsoftExcel.version
          }`,
          `live Microsoft Excel calculation features: ${input.microsoftExcelLiveCalculationScorecard.summary.coveredFeatures.join(', ')}`,
          `live Google Sheets calculation scorecard passes ${String(
            input.googleSheetsLiveCalculationScorecard.summary.matchingCaseCount,
          )}/${String(input.googleSheetsLiveCalculationScorecard.summary.requiredCaseCount)} required cases via native Google Sheets conversion`,
          `live Google Sheets calculation features: ${input.googleSheetsLiveCalculationScorecard.summary.coveredFeatures.join(', ')}`,
        ],
        evidenceArtifacts: [
          input.calculationSemanticsScorecardPath,
          input.formulaSnapshotPath,
          input.googleSheetsLiveCalculationScorecardPath,
          input.microsoftExcelLiveCalculationScorecardPath,
          'docs/excel-parity-program.md',
          'docs/formula-oracle-capture.md',
        ],
        checkCommands: [
          'pnpm calculation:semantics:check',
          'pnpm formula:dominance:check',
          'pnpm calculation:excel-live:check',
          'pnpm calculation:google-sheets-live:check',
          'pnpm test:correctness:formula',
        ],
        blockers: [
          ...formulaMissingFunctionBlockers(input.formulaSnapshot.formulaBreadth.missingOfficeFunctions.length),
          ...(input.microsoftExcelLiveCalculationScorecard.summary.allRequiredCasesPassed
            ? []
            : ['live Microsoft Excel calculation scorecard has failing required cases']),
          ...(input.googleSheetsLiveCalculationScorecard.summary.allRequiredCasesPassed
            ? []
            : ['live Google Sheets calculation scorecard has failing required cases']),
          ...(calculationSemanticsPassed ? [] : ['committed formula semantics scorecard does not cover every canonical fixture']),
        ],
      },
      {
        id: 'recalculation-speed',
        title: 'Recalculation Speed',
        objectiveCategory: 'recalculation speed',
        target: '10x mean and p95 wins on named recalculation workloads against each comparison target.',
        status: recalculationDirectTargetsTenXPassed ? 'repo-proved-lead' : 'partial-repo-evidence',
        currentEvidence: [
          familyWinSummary(dirtyExecution),
          familyWinSummary(batchEdit),
          familyWinSummary(rebuild),
          `HyperFormula proxy scorecard is tracked separately: ${input.competitiveArtifact.scorecard.workpaperWins}/${input.competitiveArtifact.scorecard.comparableCount} WorkPaper wins`,
          `live Microsoft Excel recalculation scorecard passes ${String(
            microsoftExcelRecalculationPassedCaseCount,
          )}/${String(input.microsoftExcelLiveRecalculationScorecard.summary.requiredCaseCount)} required cases on Excel ${
            input.microsoftExcelLiveRecalculationScorecard.microsoftExcel.version
          }`,
          `live Microsoft Excel recalculation workloads with 10x mean+p95 wins: ${String(
            input.microsoftExcelLiveRecalculationScorecard.summary.tenXMeanAndP95CaseCount,
          )}/${String(input.microsoftExcelLiveRecalculationScorecard.summary.requiredCaseCount)}`,
          `live Google Sheets recalculation scorecard passes ${String(
            googleSheetsRecalculationPassedCaseCount,
          )}/${String(input.googleSheetsLiveRecalculationScorecard.summary.requiredCaseCount)} required cases via native Google Sheets conversion`,
          `live Google Sheets recalculation workloads with 10x mean+p95 wins: ${String(
            input.googleSheetsLiveRecalculationScorecard.summary.tenXMeanAndP95CaseCount,
          )}/${String(input.googleSheetsLiveRecalculationScorecard.summary.requiredCaseCount)}`,
        ],
        evidenceArtifacts: [
          input.competitiveArtifactPath,
          input.googleSheetsLiveRecalculationScorecardPath,
          input.microsoftExcelLiveRecalculationScorecardPath,
        ],
        checkCommands: [
          'pnpm workpaper:bench:competitive:check',
          'pnpm recalculation:excel-live:check',
          'pnpm recalculation:google-sheets-live:check',
          'pnpm bench:contracts',
        ],
        blockers: [
          ...(microsoftExcelRecalculationTenXPassed
            ? []
            : ['live Microsoft Excel recalculation timing scorecard does not prove 10x mean+p95 for all recalculation cases']),
          ...(googleSheetsRecalculationTenXPassed
            ? []
            : ['live Google Sheets recalculation timing scorecard does not prove 10x mean+p95 for all recalculation cases']),
        ],
      },
      {
        id: 'structural-edit-performance',
        title: 'Structural Edit Performance',
        objectiveCategory: 'structural-edit performance',
        target: '10x mean and p95 wins for insert/delete/move rows and columns at workbook scale.',
        status: structuralDirectTargetsTenXPassed ? 'repo-proved-lead' : 'partial-repo-evidence',
        currentEvidence: [
          familyWinSummary(structuralRows),
          familyWinSummary(structuralColumns),
          `live Microsoft Excel structural scorecard passes ${String(
            microsoftExcelStructuralPassedCaseCount,
          )}/${String(input.microsoftExcelLiveStructuralScorecard.summary.requiredCaseCount)} required cases on Excel ${
            input.microsoftExcelLiveStructuralScorecard.microsoftExcel.version
          }`,
          `live Microsoft Excel structural operations with 10x mean+p95 wins: ${String(
            input.microsoftExcelLiveStructuralScorecard.summary.tenXMeanAndP95CaseCount,
          )}/${String(input.microsoftExcelLiveStructuralScorecard.summary.requiredCaseCount)}`,
          `live Google Sheets structural scorecard passes ${String(
            googleSheetsStructuralPassedCaseCount,
          )}/${String(input.googleSheetsLiveStructuralScorecard.summary.requiredCaseCount)} required cases via native Google Sheets conversion`,
          `live Google Sheets structural operations with 10x mean+p95 wins: ${String(
            input.googleSheetsLiveStructuralScorecard.summary.tenXMeanAndP95CaseCount,
          )}/${String(input.googleSheetsLiveStructuralScorecard.summary.requiredCaseCount)}`,
        ],
        evidenceArtifacts: [
          input.competitiveArtifactPath,
          input.googleSheetsLiveStructuralScorecardPath,
          input.microsoftExcelLiveStructuralScorecardPath,
        ],
        checkCommands: [
          'pnpm workpaper:bench:competitive:check',
          'pnpm structural:excel-live:check',
          'pnpm structural:google-sheets-live:check',
        ],
        blockers: [
          ...(microsoftExcelStructuralTenXPassed
            ? []
            : ['live Microsoft Excel structural timing scorecard does not prove 10x mean+p95 for all structural cases']),
          ...(googleSheetsStructuralTenXPassed
            ? []
            : ['live Google Sheets structural timing scorecard does not prove 10x mean+p95 for all structural cases']),
        ],
      },
      {
        id: 'large-workbook-scale',
        title: 'Large Workbook Scale',
        objectiveCategory: 'large-workbook scale',
        target: 'Sub-second warm start, import, viewport, paste, sort, and filter behavior on 100k to 250k row workbooks.',
        status: largeWorkbookDirectTargetsTenXPassed ? 'repo-proved-lead' : 'partial-repo-evidence',
        currentEvidence: [
          'local-first worker architecture and Zero-backed sync model are documented',
          'range-read and build families have HyperFormula comparison evidence',
          familyWinSummary(rangeRead),
          `large-workbook SLO artifact covers ${input.largeWorkbookSloScorecard.summary.coveredLargeWorkbookRows.join(', ')} materialized-cell sessions`,
          sloSummary(load100k),
          sloSummary(load250k),
          sloSummary(workerWarmStart100k),
          sloSummary(workerWarmStart250k),
          `external Google Sheets large-workbook evidence: ${input.largeWorkbookSloScorecard.summary.externalGoogleSheetsEvidence}`,
          `external Microsoft Excel large-workbook evidence: ${input.largeWorkbookSloScorecard.summary.externalMicrosoftExcelEvidence}`,
          `live Microsoft Excel large-workbook scorecard passes ${String(
            microsoftExcelLargeWorkbookPassedCaseCount,
          )}/${String(input.microsoftExcelLiveLargeWorkbookScorecard.summary.requiredCaseCount)} required cases on Excel ${
            input.microsoftExcelLiveLargeWorkbookScorecard.microsoftExcel.version
          }`,
          `live Microsoft Excel large-workbook cases with 10x mean+p95 wins: ${String(
            input.microsoftExcelLiveLargeWorkbookScorecard.summary.tenXMeanAndP95CaseCount,
          )}/${String(input.microsoftExcelLiveLargeWorkbookScorecard.summary.requiredCaseCount)}`,
          `live Google Sheets large-workbook scorecard passes ${String(
            googleSheetsLargeWorkbookPassedCaseCount,
          )}/${String(input.googleSheetsLiveLargeWorkbookScorecard.summary.requiredCaseCount)} required cases via native Google Sheets conversion`,
          `live Google Sheets large-workbook cases with 10x mean+p95 wins: ${String(
            input.googleSheetsLiveLargeWorkbookScorecard.summary.tenXMeanAndP95CaseCount,
          )}/${String(input.googleSheetsLiveLargeWorkbookScorecard.summary.requiredCaseCount)}`,
          `external large-workbook comparison dimensions pass: ${String(
            input.largeWorkbookSloScorecard.externalSheetsExcelComparison.requiredDimensionsPassed,
          )}`,
          `headed browser frame p95 contracts pass: ${String(
            input.largeWorkbookSloScorecard.summary.headedBrowserFrameP95ContractsPassed,
          )}`,
          ...headedBrowserScaleContracts.map(headedBrowserContractSummary),
        ],
        evidenceArtifacts: [
          input.competitiveArtifactPath,
          input.largeWorkbookSloScorecardPath,
          input.googleSheetsLiveLargeWorkbookScorecardPath,
          input.microsoftExcelLiveLargeWorkbookScorecardPath,
          input.largeWorkbookSloScorecard.source.externalLargeWorkbookComparisonArtifact,
          'e2e/tests/web-shell-scroll-performance.pw.ts',
          'docs/05-06-next-phase.md',
        ],
        checkCommands: [
          'pnpm large-workbook:slo:check',
          'pnpm large-workbook:excel-live:check',
          'pnpm large-workbook:google-sheets-live:check',
          'CI=1 pnpm bench:contracts',
          'pnpm test:browser:full',
          'pnpm bench:smoke',
        ],
        blockers: [
          ...(microsoftExcelLargeWorkbookTenXPassed
            ? []
            : ['live Microsoft Excel large-workbook timing scorecard does not prove 10x mean+p95 for all large-workbook cases']),
          ...(googleSheetsLargeWorkbookTenXPassed
            ? []
            : ['live Google Sheets large-workbook timing scorecard does not prove 10x mean+p95 for all large-workbook cases']),
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
          `external Google Sheets UI responsiveness evidence: ${input.largeWorkbookSloScorecard.summary.externalUiResponsivenessGoogleSheetsEvidence}`,
          `external Microsoft Excel UI responsiveness evidence: ${input.largeWorkbookSloScorecard.summary.externalUiResponsivenessMicrosoftExcelEvidence}`,
          `external UI responsiveness comparison dimensions pass: ${String(
            input.largeWorkbookSloScorecard.uiResponsivenessExternalSheetsExcelComparison.requiredDimensionsPassed,
          )}`,
          `live incumbent browser timing scorecard passes: ${String(uiResponsivenessLiveBrowserPassed)}`,
          `same-corpus UI 10x proof captured: ${String(input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof.captured)}`,
          `same-corpus UI 10x cases: ${String(
            input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof.tenXMeanAndP95CaseCount,
          )}/${String(input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof.requiredCaseCount)}`,
          ...input.uiResponsivenessLiveBrowserScorecard.cases.map(
            (entry) =>
              `${entry.vendor}: ${entry.workload} scrollResponseMs.p95 ${entry.scrollResponseMs.p95}ms and postScrollFrameMs.p95 ${entry.postScrollFrameMs.p95}ms (${entry.sampleCount} samples, ${entry.accessMode})`,
          ),
          ...input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof.cases.map(
            (entry) =>
              `${entry.corpusCaseId}: ${entry.workload} bilig p95 ratio ${entry.biligToGoogleSheetsP95Ratio} vs Google Sheets and ${entry.biligToMicrosoftExcelWebP95Ratio} vs Excel Web`,
          ),
          ...headedBrowserUiContracts.map(headedBrowserContractSummary),
        ],
        evidenceArtifacts: [
          input.largeWorkbookSloScorecardPath,
          input.largeWorkbookSloScorecard.source.externalUiResponsivenessComparisonArtifact,
          input.uiResponsivenessLiveBrowserScorecardPath,
          'e2e/tests/web-shell-scroll-performance.pw.ts',
          'docs/05-06-next-phase.md',
          'apps/web/src/perf/workbook-perf.ts',
          'apps/web/src/perf/workbook-scroll-perf.ts',
        ],
        checkCommands: [
          'pnpm ui:same-corpus:capture -- --emit-xlsx <directory>',
          'pnpm ui:same-corpus:fixture:check',
          'pnpm ui:same-corpus:capture -- --save-storage-state <state.json> --auth-product google-sheets --google-sheets-url <url>',
          'pnpm ui:same-corpus:capture -- --preflight --google-sheets-url <url> --microsoft-excel-web-url <url> [--google-sheets-storage-state <state.json>]',
          'pnpm ui:same-corpus:capture -- --output <capture.json> --google-sheets-url <url> --microsoft-excel-web-url <url>',
          'pnpm ui:same-corpus:capture -- --output <capture.json> --google-sheets-url <url> --google-sheets-storage-state <state.json> --microsoft-excel-web-url <url>',
          'pnpm ui:browser-live:generate -- --capture <capture.json>',
          'pnpm large-workbook:slo:check',
          'pnpm ui:browser-live:check',
          'CI=1 pnpm bench:contracts',
          'pnpm test:browser:full',
          'pnpm bench:smoke',
        ],
        blockers: [
          ...(uiResponsivenessLiveBrowserPassed
            ? []
            : ['no direct Sheets or Excel browser responsiveness live timing artifact exists in the repo']),
          ...(uiSameCorpusTenXGap ? ['live UI browser evidence is direct, but it is not a same-corpus 10x proof against incumbents'] : []),
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
          `generated collaboration scorecard passes required controls: ${String(
            input.collaborationScorecard.summary.allRequiredControlsPassed,
          )}`,
          `covered collaboration controls: ${input.collaborationScorecard.summary.coveredControls.join(', ')}`,
          `uncovered collaboration controls are explicitly disclosed: ${formatList(input.collaborationScorecard.summary.uncoveredControls)}`,
          `external Google Sheets collaboration evidence: ${input.collaborationScorecard.summary.externalGoogleSheetsEvidence}`,
          `external Microsoft Excel collaboration evidence: ${input.collaborationScorecard.summary.externalMicrosoftExcelEvidence}`,
          sloSummary(workerReconnectCatchUp100Pending),
        ],
        evidenceArtifacts: [
          input.collaborationScorecardPath,
          input.collaborationScorecard.source.externalCollaborationComparisonArtifact,
          input.largeWorkbookSloScorecardPath,
          'e2e/tests/web-shell-scroll-performance.pw.ts',
          'docs/05-06-next-phase.md',
          'apps/web/src/__tests__/workbook-sync.test.ts',
          'apps/web/src/__tests__/workbook-presence.test.tsx',
        ],
        checkCommands: [
          'pnpm collaboration:check',
          'pnpm large-workbook:slo:check',
          'CI=1 pnpm bench:contracts',
          'pnpm exec vitest run scripts/__tests__/collaboration-scorecard.test.ts apps/web/src/__tests__/worker-runtime-reconnect.test.ts apps/web/src/__tests__/workbook-presence.test.tsx apps/web/src/__tests__/projected-viewport-patch-application.test.ts apps/web/src/__tests__/worker-workbook-app-model.test.ts apps/bilig/src/workbook-runtime/document-presence-session-store.test.ts packages/zero-sync/src/__tests__/mutators.test.ts',
          'pnpm test:correctness:browser',
          'pnpm test:correctness:server',
        ],
        blockers: input.collaborationScorecard.summary.uncoveredControls.includes('externalSheetsCollaborationComparison')
          ? [
              `generated collaboration evidence still leaves uncovered controls: ${input.collaborationScorecard.summary.uncoveredControls.join(', ')}`,
              'no direct Sheets collaboration comparison artifact exists in the repo',
            ]
          : [],
      },
      {
        id: 'automation-api-extensibility',
        title: 'Automation And API Extensibility',
        objectiveCategory: 'automation/API extensibility',
        target: 'Typed semantic workbook operations and agent tools that replace brittle DOM automation.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          `generated automation scorecard passes required controls: ${String(input.automationScorecard.summary.allRequiredControlsPassed)}`,
          `covered automation controls: ${input.automationScorecard.summary.coveredControls.join(', ')}`,
          `registered workbook-agent tool count: ${String(input.automationScorecard.summary.registeredToolCount)}`,
          `semantic command kinds exercised by scorecard: ${String(input.automationScorecard.summary.semanticCommandKindCount)}`,
          `uncovered automation controls are explicitly disclosed: ${formatList(input.automationScorecard.summary.uncoveredControls)}`,
          `external Google Apps Script evidence: ${input.automationScorecard.summary.externalGoogleSheetsEvidence}`,
          `external Microsoft Office Scripts evidence: ${input.automationScorecard.summary.externalMicrosoftExcelEvidence}`,
          `${totalSurfaceMembers} HyperFormula surface members are snapshotted for parity tracking`,
          `${input.surfaceSnapshot.configKeys.length} HyperFormula config keys are snapshotted for parity tracking`,
          'WorkPaper exposes additional detailed events and performance counters',
        ],
        evidenceArtifacts: [
          input.automationScorecardPath,
          input.automationScorecard.source.externalAutomationComparisonArtifact,
          input.surfaceSnapshotPath,
          'packages/headless/src/__tests__/hyperformula-surface-parity.test.ts',
          'packages/agent-api/src/__tests__/workbook-agent-bundles.test.ts',
        ],
        checkCommands: [
          'pnpm automation:check',
          'pnpm workpaper:parity:check',
          'pnpm workpaper:smoke:external',
          'pnpm exec vitest run scripts/__tests__/automation-scorecard.test.ts packages/agent-api/src/__tests__/workbook-agent-bundles.test.ts packages/headless/src/__tests__/work-paper.test.ts apps/web/src/__tests__/worker-runtime-authoritative-bootstrap.test.ts',
        ],
        blockers:
          input.automationScorecard.summary.uncoveredControls.length > 0
            ? [
                `generated automation evidence still leaves uncovered controls: ${input.automationScorecard.summary.uncoveredControls.join(', ')}`,
                'no direct generated Google Apps Script or Office Scripts execution comparison exists',
              ]
            : [],
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
          `macro-enabled workbook payload preservation covered: ${String(
            input.importExportFidelityScorecard.summary.coveredFeatures.includes('xlsx.macros.payloadRoundtrip'),
          )}`,
          `macro-enabled workbook code-name preservation covered: ${String(
            input.importExportFidelityScorecard.summary.coveredFeatures.includes('xlsx.macros.codeNameRoundtrip'),
          )}`,
          `unsupported XLSX features are explicitly disclosed: ${
            input.importExportFidelityScorecard.summary.unsupportedFeatures.join(', ') || 'none'
          }`,
          `declined unsafe runtime features: ${input.importExportFidelityScorecard.summary.declinedRuntimeFeatures.join(', ')}`,
          `external Google Sheets import/export evidence: ${input.importExportFidelityScorecard.summary.externalGoogleSheetsEvidence}`,
          `external Microsoft Excel import/export evidence: ${input.importExportFidelityScorecard.summary.externalMicrosoftExcelEvidence}`,
        ],
        evidenceArtifacts: [
          input.importExportFidelityScorecardPath,
          input.importExportFidelityScorecard.source.externalImportExportComparisonArtifact,
          'packages/excel-import/src/__tests__/excel-import.test.ts',
          'packages/core/src/__tests__/engine-import-export.fuzz.test.ts',
          'docs/formula-oracle-capture.md',
        ],
        checkCommands: [
          'pnpm import-export:fidelity:check',
          'pnpm exec vitest run packages/excel-import/src/__tests__/excel-import.test.ts packages/core/src/__tests__/engine-import-export.fuzz.test.ts',
        ],
        blockers: [
          ...(input.importExportFidelityScorecard.summary.unsupportedFeatures.length > 0
            ? [`unsupported import/export features remain: ${input.importExportFidelityScorecard.summary.unsupportedFeatures.join(', ')}`]
            : []),
          ...(input.importExportFidelityScorecard.summary.externalGoogleSheetsEvidence === 'official-docs-comparison-artifact'
            ? []
            : ['no direct Sheets import/export compatibility artifact exists in the repo']),
          ...(input.importExportFidelityScorecard.summary.externalMicrosoftExcelEvidence === 'official-docs-comparison-artifact'
            ? []
            : ['no direct Microsoft Excel import/export compatibility artifact exists in the repo']),
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
          `uncovered auditability controls are explicitly disclosed: ${formatList(input.auditabilityScorecard.summary.uncoveredControls)}`,
          `external Google Sheets auditability evidence: ${input.auditabilityScorecard.summary.externalGoogleSheetsEvidence}`,
          `external Microsoft Excel auditability evidence: ${input.auditabilityScorecard.summary.externalMicrosoftExcelEvidence}`,
          'change bundles, versions, revertable changes, and agent preview/apply rails are documented',
          'workbook changes and mutation journal tests exist',
        ],
        evidenceArtifacts: [
          input.auditabilityScorecardPath,
          input.auditabilityScorecard.source.externalAuditabilityComparisonArtifact,
          'e2e/tests/web-shell-remote-sync.pw.ts',
          'docs/05-06-next-phase.md',
          'apps/web/src/__tests__/workbook-changes.test.tsx',
          'apps/web/src/__tests__/worker-runtime-mutation-journal.test.ts',
        ],
        checkCommands: [
          'pnpm auditability:check',
          'pnpm exec vitest run packages/agent-api/src/__tests__/workbook-agent-preview.test.ts apps/bilig/src/zero/__tests__/workbook-agent-apply.test.ts packages/zero-sync/src/__tests__/workbook-history-state.test.ts',
          'pnpm test:browser e2e/tests/web-shell-remote-sync.pw.ts -g "reverts an authoritative change"',
          'pnpm test:correctness:browser',
          'pnpm test:correctness:server',
        ],
        blockers: input.auditabilityScorecard.summary.uncoveredControls.includes('externalSheetsExcelAuditabilityComparison')
          ? ['no direct incumbent auditability comparison artifact exists in the repo']
          : [],
      },
      {
        id: 'reliability',
        title: 'Reliability',
        objectiveCategory: 'reliability',
        target: 'Crash-safe local durability, deterministic replay, durable sync, and no accepted-op loss.',
        status: 'partial-repo-evidence',
        currentEvidence: [
          `generated reliability scorecard passes required controls: ${String(input.reliabilityScorecard.summary.allRequiredControlsPassed)}`,
          `covered reliability controls: ${input.reliabilityScorecard.summary.coveredControls.join(', ')}`,
          `uncovered reliability controls are explicitly disclosed: ${formatList(input.reliabilityScorecard.summary.uncoveredControls)}`,
          `external Google Sheets reliability evidence: ${input.reliabilityScorecard.summary.externalGoogleSheetsEvidence}`,
          `external Microsoft Excel reliability evidence: ${input.reliabilityScorecard.summary.externalMicrosoftExcelEvidence}`,
          'Zero-backed durability and reconnect/rebase architecture are documented',
          'runtime sync replay, fuzz, reconnect, and Zero sync tests exist',
        ],
        evidenceArtifacts: [
          input.reliabilityScorecardPath,
          input.reliabilityScorecard.source.externalReliabilityComparisonArtifact,
          'e2e/tests/web-shell-remote-sync.pw.ts',
          'docs/05-06-next-phase.md',
          'apps/web/src/__tests__/runtime-sync.fuzz.test.ts',
          'apps/web/src/__tests__/worker-runtime-reconnect.test.ts',
        ],
        checkCommands: [
          'pnpm reliability:check',
          'pnpm exec vitest run apps/web/src/__tests__/worker-runtime-reconnect.test.ts apps/web/src/__tests__/worker-runtime-authoritative-bootstrap.test.ts apps/web/src/__tests__/worker-runtime-mutation-journal.test.ts packages/zero-sync/src/__tests__/workbook-events.test.ts',
          'pnpm test:browser e2e/tests/web-shell-remote-sync.pw.ts -g "restores persisted workbook state after a full reload"',
          'pnpm test:fuzz',
          'pnpm test:correctness:browser',
          'pnpm test:correctness:server',
        ],
        blockers: input.reliabilityScorecard.summary.uncoveredControls.includes('externalSheetsExcelReliabilityComparison')
          ? ['no direct Sheets or Excel reliability comparison artifact exists in the repo']
          : [],
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
          `uncovered security controls are explicitly disclosed: ${formatList(input.securityPostureScorecard.summary.uncoveredControls)}`,
          `external Google Sheets security evidence: ${input.securityPostureScorecard.summary.externalGoogleSheetsEvidence}`,
          `external Microsoft Excel security evidence: ${input.securityPostureScorecard.summary.externalMicrosoftExcelEvidence}`,
        ],
        evidenceArtifacts: [
          input.securityPostureScorecardPath,
          input.securityPostureScorecard.source.externalSecurityComparisonArtifact,
          'pnpm-lock.yaml',
          'apps/bilig/src/http/sync-server-security-headers.ts',
          'apps/bilig/src/http/sync-server-security-headers.test.ts',
          'packages/excel-import/src/__tests__/excel-import.test.ts',
          'packages/agent-api/src/__tests__/workbook-agent-execution-policy.test.ts',
          'scripts/check-package-publish.ts',
        ],
        checkCommands: [
          'pnpm security:posture:check',
          'pnpm security:audit',
          'pnpm exec vitest run apps/bilig/src/http/sync-server-security-headers.test.ts apps/bilig/src/http/sync-server.test.ts packages/excel-import/src/__tests__/excel-import.test.ts packages/agent-api/src/__tests__/workbook-agent-execution-policy.test.ts',
          'pnpm publish:runtime:check',
        ],
        blockers: [
          ...(securityUncoveredControls.has('deployment.runtimeNetworkPolicy')
            ? ['generated security posture evidence has not yet covered deployment runtime network policy']
            : []),
          ...(securityUncoveredControls.has('externalSheetsExcelSecurityComparison')
            ? ['no direct Sheets or Excel security comparison artifact exists in the repo']
            : []),
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
          'generated-source CI checks are serialized to avoid pnpm workspace-state races in the evidence gate',
          'this generated scorecard prevents blanket 10x claims from outrunning evidence',
          `dominance package scripts present: ${String(
            operatorWorkflowEvidence.dominanceGenerateScriptPresent &&
              operatorWorkflowEvidence.dominanceCheckScriptPresent &&
              operatorWorkflowEvidence.dominanceAuditCheckScriptPresent,
          )}`,
          `run-ci executes dominance checks: ${String(
            operatorWorkflowEvidence.runCiDominanceCheckPresent && operatorWorkflowEvidence.runCiDominanceAuditCheckPresent,
          )}`,
          `blanket claim policy coupled to completion audit: ${String(operatorWorkflowEvidence.blanketClaimPolicyCoupledToCompletionAudit)}`,
          `prompt-to-artifact audit coupled to live status: ${String(operatorWorkflowEvidence.promptArtifactAuditCoupledToLiveStatus)}`,
          `completion audit criteria passed: ${String(completionAudit.allCriteriaPassed)}`,
        ],
        evidenceArtifacts: [
          'package.json',
          'scripts/run-ci.ts',
          'scripts/bilig-dominance-operator-workflow.ts',
          input.competitiveArtifactPath,
          input.formulaSnapshotPath,
          input.surfaceSnapshotPath,
        ],
        checkCommands: ['pnpm dominance:check', 'pnpm run ci'],
        blockers: operatorWorkflowBlockers,
      },
    ],
  }
}

function buildOverallGoogleSheets10xStatus(
  input: BuildScorecardInput,
  signals: {
    readonly googleSheetsLargeWorkbookTenXPassed: boolean
    readonly googleSheetsRecalculationTenXPassed: boolean
    readonly googleSheetsStructuralTenXPassed: boolean
    readonly uiResponsivenessLiveBrowserPassed: boolean
    readonly uiSameCorpusTenXGap: boolean
  },
): OverallGoogleSheets10xStatus {
  const categories = [
    {
      id: 'recalculation-speed',
      requirement: 'Every claimed recalculation workload must prove at least 10x better mean and p95 than live Google Sheets.',
      passed: signals.googleSheetsRecalculationTenXPassed,
      evidence: [
        `Google Sheets recalculation 10x cases: ${String(input.googleSheetsLiveRecalculationScorecard.summary.tenXMeanAndP95CaseCount)}/${String(
          input.googleSheetsLiveRecalculationScorecard.summary.requiredCaseCount,
        )}`,
        `evidence kind: ${input.googleSheetsLiveRecalculationScorecard.source.evidenceKind}`,
      ],
      gaps: signals.googleSheetsRecalculationTenXPassed
        ? []
        : ['live Google Sheets recalculation scorecard is not 10x for every required case'],
    },
    {
      id: 'structural-edit-performance',
      requirement: 'Every claimed structural-edit workload must prove at least 10x better mean and p95 than live Google Sheets.',
      passed: signals.googleSheetsStructuralTenXPassed,
      evidence: [
        `Google Sheets structural 10x cases: ${String(input.googleSheetsLiveStructuralScorecard.summary.tenXMeanAndP95CaseCount)}/${String(
          input.googleSheetsLiveStructuralScorecard.summary.requiredCaseCount,
        )}`,
        `evidence kind: ${input.googleSheetsLiveStructuralScorecard.source.evidenceKind}`,
      ],
      gaps: signals.googleSheetsStructuralTenXPassed ? [] : ['live Google Sheets structural scorecard is not 10x for every required case'],
    },
    {
      id: 'large-workbook-scale',
      requirement: 'Every claimed large-workbook workload must prove at least 10x better mean and p95 than live Google Sheets.',
      passed: signals.googleSheetsLargeWorkbookTenXPassed,
      evidence: [
        `Google Sheets large-workbook 10x cases: ${String(
          input.googleSheetsLiveLargeWorkbookScorecard.summary.tenXMeanAndP95CaseCount,
        )}/${String(input.googleSheetsLiveLargeWorkbookScorecard.summary.requiredCaseCount)}`,
        `evidence kind: ${input.googleSheetsLiveLargeWorkbookScorecard.source.evidenceKind}`,
      ],
      gaps: signals.googleSheetsLargeWorkbookTenXPassed
        ? []
        : ['live Google Sheets large-workbook scorecard is not 10x for every required case'],
    },
    {
      id: 'ui-responsiveness',
      requirement:
        'Claimed UI responsiveness must have live same-corpus browser proof against Google Sheets with 10x better mean and p95 plus rendered-grid proof.',
      passed: signals.uiResponsivenessLiveBrowserPassed && !signals.uiSameCorpusTenXGap,
      evidence: [
        `direct live browser timing passed: ${String(signals.uiResponsivenessLiveBrowserPassed)}`,
        `same-corpus capture kind: ${input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof.evidenceKind}`,
        `same-corpus UI 10x cases: ${String(
          input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof.tenXMeanAndP95CaseCount,
        )}/${String(input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof.requiredCaseCount)}`,
      ],
      gaps: [
        ...(signals.uiResponsivenessLiveBrowserPassed ? [] : ['live incumbent browser timing scorecard is not passing']),
        ...(signals.uiSameCorpusTenXGap
          ? ['live UI browser evidence is not a same-corpus 10x proof against Google Sheets with rendered-grid proof']
          : []),
      ],
    },
  ] satisfies OverallGoogleSheets10xStatus['categories']
  const unmetRequirements = categories.filter((entry) => !entry.passed).map((entry) => `${entry.id}: ${entry.gaps.join('; ')}`)
  return {
    passed: unmetRequirements.length === 0,
    status: unmetRequirements.length === 0 ? 'passed' : 'blocked',
    requirement:
      'A broad Google Sheets 10x claim is allowed only when every claimed performance category has checked-in live Google Sheets evidence and every required mean and p95 ratio is at least 10x better.',
    categories,
    unmetRequirements,
    evidenceArtifacts: [
      input.googleSheetsLiveRecalculationScorecardPath,
      input.googleSheetsLiveStructuralScorecardPath,
      input.googleSheetsLiveLargeWorkbookScorecardPath,
      input.uiResponsivenessLiveBrowserScorecardPath,
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

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
