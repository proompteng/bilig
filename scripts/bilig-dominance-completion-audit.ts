import type { BuildScorecardInput, DominanceCompletionAudit, DominanceCompletionCriterion } from './bilig-dominance-scorecard-types.ts'
import type { UiResponsivenessLiveBrowserScorecard } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { requiredUiResponsivenessSameCorpusWorkloads } from './ui-responsiveness-same-corpus-workloads.ts'

export interface CompletionAuditSignals {
  readonly calculationSemanticsPassed: boolean
  readonly googleSheetsLargeWorkbookTenXPassed: boolean
  readonly googleSheetsRecalculationTenXPassed: boolean
  readonly googleSheetsStructuralTenXPassed: boolean
  readonly largeWorkbookDirectTargetsTenXPassed: boolean
  readonly microsoftExcelLargeWorkbookTenXPassed: boolean
  readonly microsoftExcelRecalculationTenXPassed: boolean
  readonly microsoftExcelStructuralTenXPassed: boolean
  readonly recalculationDirectTargetsTenXPassed: boolean
  readonly structuralDirectTargetsTenXPassed: boolean
  readonly uiResponsivenessLiveBrowserPassed: boolean
}

export function buildBiligDominanceCompletionAudit(input: BuildScorecardInput, signals: CompletionAuditSignals): DominanceCompletionAudit {
  const uiSameCorpusTenXGap = hasUiResponsivenessSameCorpusTenXGap(input.uiResponsivenessLiveBrowserScorecard)
  const criteria = [
    criterion({
      id: 'calculation-correctness',
      requirement: 'Calculation correctness must cover the committed formula surface and match live Sheets and Excel required cases.',
      evidence: [
        `${String(input.formulaSnapshot.formulaBreadth.officeListed.production)}/${String(
          input.formulaSnapshot.formulaBreadth.officeListed.total,
        )} Office-listed functions are production-routed`,
        `${String(input.formulaSnapshot.canonical.summary.production)}/${String(
          input.formulaSnapshot.canonical.summary.total,
        )} canonical formula rows are production-routed`,
        `${String(input.calculationSemanticsScorecard.summary.coveredCanonicalFixtureCount)}/${String(
          input.calculationSemanticsScorecard.summary.canonicalFormulaFixtureCount,
        )} canonical fixtures are covered by the calculation semantics scorecard`,
        `${String(input.microsoftExcelLiveCalculationScorecard.summary.matchingCaseCount)}/${String(
          input.microsoftExcelLiveCalculationScorecard.summary.requiredCaseCount,
        )} live Microsoft Excel calculation cases match`,
        `${String(input.googleSheetsLiveCalculationScorecard.summary.matchingCaseCount)}/${String(
          input.googleSheetsLiveCalculationScorecard.summary.requiredCaseCount,
        )} live Google Sheets calculation cases match`,
      ],
      gaps: [
        ...percentGaps('Office-listed formula breadth', input.formulaSnapshot.formulaBreadth.officeListed.percent),
        ...percentGaps('tracked formula breadth', input.formulaSnapshot.formulaBreadth.tracked.percent),
        ...percentGaps('canonical formula production routing', input.formulaSnapshot.canonical.summary.percent),
        ...(signals.calculationSemanticsPassed ? [] : ['calculation semantics scorecard does not cover every committed fixture']),
        ...(input.microsoftExcelLiveCalculationScorecard.summary.allRequiredCasesPassed
          ? []
          : ['live Microsoft Excel calculation scorecard has failing required cases']),
        ...(input.googleSheetsLiveCalculationScorecard.summary.allRequiredCasesPassed
          ? []
          : ['live Google Sheets calculation scorecard has failing required cases']),
      ],
    }),
    criterion({
      id: 'recalculation-speed',
      requirement: 'Recalculation speed must prove 10x mean and p95 wins against both live Sheets and live Excel required cases.',
      evidence: [
        `Microsoft Excel recalculation 10x cases: ${String(input.microsoftExcelLiveRecalculationScorecard.summary.tenXMeanAndP95CaseCount)}/${String(
          input.microsoftExcelLiveRecalculationScorecard.summary.requiredCaseCount,
        )}`,
        `Google Sheets recalculation 10x cases: ${String(input.googleSheetsLiveRecalculationScorecard.summary.tenXMeanAndP95CaseCount)}/${String(
          input.googleSheetsLiveRecalculationScorecard.summary.requiredCaseCount,
        )}`,
      ],
      gaps: [
        ...(signals.microsoftExcelRecalculationTenXPassed
          ? []
          : ['live Microsoft Excel recalculation scorecard is not 10x for every case']),
        ...(signals.googleSheetsRecalculationTenXPassed ? [] : ['live Google Sheets recalculation scorecard is not 10x for every case']),
        ...(signals.recalculationDirectTargetsTenXPassed ? [] : ['direct recalculation targets are not all 10x mean+p95 wins']),
      ],
    }),
    criterion({
      id: 'structural-edit-performance',
      requirement: 'Structural edit performance must prove 10x mean and p95 wins against both live Sheets and live Excel required cases.',
      evidence: [
        `Microsoft Excel structural 10x cases: ${String(input.microsoftExcelLiveStructuralScorecard.summary.tenXMeanAndP95CaseCount)}/${String(
          input.microsoftExcelLiveStructuralScorecard.summary.requiredCaseCount,
        )}`,
        `Google Sheets structural 10x cases: ${String(input.googleSheetsLiveStructuralScorecard.summary.tenXMeanAndP95CaseCount)}/${String(
          input.googleSheetsLiveStructuralScorecard.summary.requiredCaseCount,
        )}`,
      ],
      gaps: [
        ...(signals.microsoftExcelStructuralTenXPassed ? [] : ['live Microsoft Excel structural scorecard is not 10x for every case']),
        ...(signals.googleSheetsStructuralTenXPassed ? [] : ['live Google Sheets structural scorecard is not 10x for every case']),
        ...(signals.structuralDirectTargetsTenXPassed ? [] : ['direct structural targets are not all 10x mean+p95 wins']),
      ],
    }),
    criterion({
      id: 'large-workbook-scale',
      requirement: 'Large workbook scale must meet SLOs and prove 10x mean and p95 wins against both live Sheets and live Excel cases.',
      evidence: [
        `SLO rows covered: ${input.largeWorkbookSloScorecard.summary.coveredLargeWorkbookRows.join(', ')}`,
        `Microsoft Excel large-workbook 10x cases: ${String(
          input.microsoftExcelLiveLargeWorkbookScorecard.summary.tenXMeanAndP95CaseCount,
        )}/${String(input.microsoftExcelLiveLargeWorkbookScorecard.summary.requiredCaseCount)}`,
        `Google Sheets large-workbook 10x cases: ${String(input.googleSheetsLiveLargeWorkbookScorecard.summary.tenXMeanAndP95CaseCount)}/${String(
          input.googleSheetsLiveLargeWorkbookScorecard.summary.requiredCaseCount,
        )}`,
      ],
      gaps: [
        ...(input.largeWorkbookSloScorecard.summary.allSloBudgetsPassed ? [] : ['large-workbook SLO budgets are not all passing']),
        ...(input.largeWorkbookSloScorecard.summary.headedBrowserFrameP95ContractsPassed
          ? []
          : ['headed browser frame p95 contracts are not all passing']),
        ...(signals.microsoftExcelLargeWorkbookTenXPassed
          ? []
          : ['live Microsoft Excel large-workbook scorecard is not 10x for every case']),
        ...(signals.googleSheetsLargeWorkbookTenXPassed ? [] : ['live Google Sheets large-workbook scorecard is not 10x for every case']),
        ...(signals.largeWorkbookDirectTargetsTenXPassed ? [] : ['direct large-workbook targets are not all 10x mean+p95 wins']),
      ],
    }),
    criterion({
      id: 'ui-responsiveness',
      requirement: 'UI responsiveness must include live incumbent browser evidence and headed large-workbook frame-budget contracts.',
      evidence: [
        `live browser vendors: ${input.uiResponsivenessLiveBrowserScorecard.summary.capturedVendors.join(', ')}`,
        `headed browser contracts passed: ${String(input.largeWorkbookSloScorecard.summary.headedBrowserFrameP95ContractsPassed)}`,
      ],
      gaps: [
        ...(signals.uiResponsivenessLiveBrowserPassed ? [] : ['live incumbent browser timing scorecard is not passing']),
        ...(input.largeWorkbookSloScorecard.summary.headedBrowserFrameP95ContractsPassed
          ? []
          : ['headed browser frame p95 contracts are not all passing']),
        ...(uiSameCorpusTenXGap ? ['live UI browser evidence is not a same-corpus 10x proof against incumbents'] : []),
      ],
    }),
    requiredControlCriterion({
      id: 'collaboration',
      requirement: 'Collaboration controls must pass with no uncovered generated controls.',
      passed: input.collaborationScorecard.summary.allRequiredControlsPassed,
      coveredControls: input.collaborationScorecard.summary.coveredControls,
      uncoveredControls: input.collaborationScorecard.summary.uncoveredControls,
    }),
    requiredControlCriterion({
      id: 'automation-api-extensibility',
      requirement: 'Automation and API extensibility controls must pass with no uncovered generated controls.',
      passed: input.automationScorecard.summary.allRequiredControlsPassed,
      coveredControls: input.automationScorecard.summary.coveredControls,
      uncoveredControls: input.automationScorecard.summary.uncoveredControls,
    }),
    criterion({
      id: 'import-export-compatibility',
      requirement: 'Import/export compatibility must pass required cases and leave no unsupported feature required by the blanket claim.',
      evidence: [
        `covered import/export features: ${input.importExportFidelityScorecard.summary.coveredFeatures.join(', ')}`,
        `unsupported import/export features: ${input.importExportFidelityScorecard.summary.unsupportedFeatures.join(', ') || 'none'}`,
        `declined unsafe runtime features: ${input.importExportFidelityScorecard.summary.declinedRuntimeFeatures.join(', ') || 'none'}`,
      ],
      gaps: [
        ...(input.importExportFidelityScorecard.summary.allRequiredCasesPassed
          ? []
          : ['import/export fidelity scorecard has failing cases']),
        ...input.importExportFidelityScorecard.summary.unsupportedFeatures.map(
          (feature) => `unsupported import/export feature: ${feature}`,
        ),
      ],
    }),
    requiredControlCriterion({
      id: 'auditability',
      requirement: 'Auditability controls must pass with no uncovered generated controls.',
      passed: input.auditabilityScorecard.summary.allRequiredControlsPassed,
      coveredControls: input.auditabilityScorecard.summary.coveredControls,
      uncoveredControls: input.auditabilityScorecard.summary.uncoveredControls,
    }),
    requiredControlCriterion({
      id: 'reliability',
      requirement: 'Reliability controls must pass with no uncovered generated controls.',
      passed: input.reliabilityScorecard.summary.allRequiredControlsPassed,
      coveredControls: input.reliabilityScorecard.summary.coveredControls,
      uncoveredControls: input.reliabilityScorecard.summary.uncoveredControls,
    }),
    requiredControlCriterion({
      id: 'security',
      requirement: 'Security controls must pass with no uncovered generated controls.',
      passed: input.securityPostureScorecard.summary.allRequiredControlsPassed,
      coveredControls: input.securityPostureScorecard.summary.coveredControls,
      uncoveredControls: input.securityPostureScorecard.summary.uncoveredControls,
    }),
    criterion({
      id: 'operator-developer-workflow',
      requirement: 'Operator workflow must preserve one-command verification and prevent unverifiable blanket claims.',
      evidence: ['dominance:check is part of fast CI', 'the completion audit drives the blanket claim policy'],
      gaps: [],
    }),
  ]
  const unmetRequirements = criteria.filter((entry) => !entry.passed).map((entry) => `${entry.id}: ${entry.gaps.join('; ')}`)
  return {
    allCriteriaPassed: unmetRequirements.length === 0,
    unmetRequirements,
    criteria,
  }
}

export function hasUiResponsivenessSameCorpusTenXGap(scorecard: UiResponsivenessLiveBrowserScorecard): boolean {
  if (
    !scorecard.sameCorpusProof.captured ||
    scorecard.sameCorpusProof.requiredCaseCount === 0 ||
    scorecard.sameCorpusProof.cases.length !== scorecard.sameCorpusProof.requiredCaseCount ||
    scorecard.sameCorpusProof.tenXMeanAndP95CaseCount !== scorecard.sameCorpusProof.requiredCaseCount ||
    requiredUiResponsivenessSameCorpusWorkloads.some(
      (workload) => !scorecard.sameCorpusProof.cases.some((entry) => entry.workload === workload),
    ) ||
    scorecard.sameCorpusProof.cases.some((entry) => !entry.passed)
  ) {
    return true
  }
  const limitations = [
    ...scorecard.sameCorpusProof.limitations,
    ...scorecard.sameCorpusProof.cases.flatMap((entry) => [
      ...entry.bilig.limitations,
      ...entry.googleSheets.limitations,
      ...entry.microsoftExcelWeb.limitations,
    ]),
  ].map((entry) => entry.toLowerCase())
  return limitations.some(
    (entry) =>
      entry.includes('not a same-corpus 10x proof') ||
      entry.includes('not an authenticated same-corpus') ||
      entry.includes('not live timing') ||
      entry.includes('does not claim bilig is 10x'),
  )
}

function requiredControlCriterion(args: {
  readonly coveredControls: readonly string[]
  readonly id: string
  readonly passed: boolean
  readonly requirement: string
  readonly uncoveredControls: readonly string[]
}): DominanceCompletionCriterion {
  return criterion({
    id: args.id,
    requirement: args.requirement,
    evidence: [`covered controls: ${args.coveredControls.join(', ')}`],
    gaps: [
      ...(args.passed ? [] : ['required controls are not all passing']),
      ...args.uncoveredControls.map((control) => `uncovered control: ${control}`),
    ],
  })
}

function criterion(args: {
  readonly evidence: readonly string[]
  readonly gaps: readonly string[]
  readonly id: string
  readonly requirement: string
}): DominanceCompletionCriterion {
  const gaps = [...args.gaps]
  return {
    id: args.id,
    requirement: args.requirement,
    passed: gaps.length === 0,
    evidence: [...args.evidence],
    gaps,
  }
}

function percentGaps(label: string, percent: number): string[] {
  return percent === 100 ? [] : [`${label} is ${String(percent)}%, not 100%`]
}
