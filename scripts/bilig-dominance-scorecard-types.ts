import type { AuditabilityScorecard } from './gen-auditability-scorecard.ts'
import type { AutomationScorecard } from './gen-automation-scorecard.ts'
import type { CalculationSemanticsScorecard } from './gen-calculation-semantics-scorecard.ts'
import type { CollaborationScorecard } from './gen-collaboration-scorecard.ts'
import type { GoogleSheetsLiveCalculationScorecard } from './gen-google-sheets-live-calculation-scorecard.ts'
import type { GoogleSheetsLiveLargeWorkbookScorecard } from './gen-google-sheets-live-large-workbook-scorecard.ts'
import type { GoogleSheetsLiveRecalculationScorecard } from './gen-google-sheets-live-recalculation-scorecard.ts'
import type { GoogleSheetsLiveStructuralScorecard } from './gen-google-sheets-live-structural-scorecard.ts'
import type { ImportExportFidelityScorecard } from './gen-import-export-fidelity-scorecard.ts'
import type {
  HeadedBrowserFrameP95Contract,
  LargeWorkbookSloMeasurement,
  LargeWorkbookSloScorecard,
} from './gen-large-workbook-slo-scorecard.ts'
import type { MicrosoftExcelLiveCalculationScorecard } from './gen-microsoft-excel-live-calculation-scorecard.ts'
import type { MicrosoftExcelLiveLargeWorkbookScorecard } from './gen-microsoft-excel-live-large-workbook-scorecard.ts'
import type { MicrosoftExcelLiveRecalculationScorecard } from './gen-microsoft-excel-live-recalculation-scorecard.ts'
import type { MicrosoftExcelLiveStructuralScorecard } from './gen-microsoft-excel-live-structural-scorecard.ts'
import type { ReliabilityScorecard } from './gen-reliability-scorecard.ts'
import type { SecurityPostureScorecard } from './gen-security-posture-scorecard.ts'
import type { UiResponsivenessLiveBrowserScorecard } from './gen-ui-responsiveness-live-browser-scorecard.ts'

export type { HeadedBrowserFrameP95Contract, LargeWorkbookSloMeasurement, LargeWorkbookSloScorecard }

export type DominanceStatus = 'repo-proved-lead' | 'partial-repo-evidence' | 'target-only'
export type DominanceGoalStatus = 'achieved' | 'active-not-achieved'

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
  goalStatus: DominanceGoalStatus
  claimPolicy: {
    blanketTenXClaimAllowed: boolean
    requiredForBlanketTenXClaim: string[]
    unmetRequirements: string[]
    workloadSpecificTenXWins: Array<{
      workload: string
      meanRatio: number
      p95Ratio: number
      comparisonTarget: 'HyperFormula'
    }>
  }
  completionAudit: DominanceCompletionAudit
  sourceArtifacts: {
    auditabilityScorecard: string
    automationScorecard: string
    collaborationScorecard: string
    calculationSemanticsScorecard: string
    formulaDominanceSnapshot: string
    googleSheetsLiveCalculationScorecard: string
    googleSheetsLiveRecalculationScorecard: string
    googleSheetsLiveStructuralScorecard: string
    googleSheetsLiveLargeWorkbookScorecard: string
    hyperFormulaSurfaceSnapshot: string
    microsoftExcelLiveCalculationScorecard: string
    microsoftExcelLiveRecalculationScorecard: string
    microsoftExcelLiveLargeWorkbookScorecard: string
    microsoftExcelLiveStructuralScorecard: string
    importExportFidelityScorecard: string
    largeWorkbookSloScorecard: string
    uiResponsivenessLiveBrowserScorecard: string
    reliabilityScorecard: string
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
    automationCoveredControls: string[]
    automationPosturePassed: boolean
    automationRegisteredToolCount: number
    automationSemanticCommandKindCount: number
    automationUncoveredControls: string[]
    collaborationCoveredControls: string[]
    collaborationPosturePassed: boolean
    collaborationUncoveredControls: string[]
    calculationSemanticsCoveredCanonicalFixtureCount: number
    calculationSemanticsCoveredWorkbookSemanticsFixtureCount: number
    calculationSemanticsPassed: boolean
    externalGoogleSheetsEvidence: 'not-captured-in-repo'
    externalMicrosoftExcelEvidence: 'not-captured-in-repo'
    formulaCanonicalProductionPercent: number
    googleSheetsLiveCalculationEvidence: 'live-google-sheets-native-conversion-via-google-drive-connector'
    googleSheetsLiveCalculationCaseCount: number
    googleSheetsLiveCalculationPassed: boolean
    googleSheetsLiveCalculationSpreadsheetId: string
    googleSheetsLiveRecalculationEvidence: 'live-google-sheets-native-conversion-via-google-drive-connector'
    googleSheetsLiveRecalculationCaseCount: number
    googleSheetsLiveRecalculationPassed: boolean
    googleSheetsLiveRecalculationTenXMeanAndP95CaseCount: number
    googleSheetsLiveRecalculationSpreadsheetIds: string[]
    googleSheetsLiveStructuralEvidence: 'live-google-sheets-native-conversion-via-google-drive-connector'
    googleSheetsLiveStructuralCaseCount: number
    googleSheetsLiveStructuralPassed: boolean
    googleSheetsLiveStructuralTenXMeanAndP95CaseCount: number
    googleSheetsLiveStructuralSpreadsheetIds: string[]
    googleSheetsLiveLargeWorkbookEvidence: 'live-google-sheets-native-conversion-via-google-drive-connector'
    googleSheetsLiveLargeWorkbookCaseCount: number
    googleSheetsLiveLargeWorkbookPassed: boolean
    googleSheetsLiveLargeWorkbookTenXMeanAndP95CaseCount: number
    googleSheetsLiveLargeWorkbookSpreadsheetIds: string[]
    microsoftExcelLiveCalculationEvidence: 'live-local-microsoft-excel-automation'
    microsoftExcelLiveCalculationCaseCount: number
    microsoftExcelLiveCalculationPassed: boolean
    microsoftExcelLiveCalculationVersion: string
    microsoftExcelLiveRecalculationEvidence: 'live-local-microsoft-excel-automation'
    microsoftExcelLiveRecalculationCaseCount: number
    microsoftExcelLiveRecalculationPassed: boolean
    microsoftExcelLiveRecalculationTenXMeanAndP95CaseCount: number
    microsoftExcelLiveRecalculationVersion: string
    microsoftExcelLiveLargeWorkbookEvidence: 'live-local-microsoft-excel-automation'
    microsoftExcelLiveLargeWorkbookCaseCount: number
    microsoftExcelLiveLargeWorkbookPassed: boolean
    microsoftExcelLiveLargeWorkbookTenXMeanAndP95CaseCount: number
    microsoftExcelLiveLargeWorkbookVersion: string
    microsoftExcelLiveStructuralEvidence: 'live-local-microsoft-excel-automation'
    microsoftExcelLiveStructuralCaseCount: number
    microsoftExcelLiveStructuralPassed: boolean
    microsoftExcelLiveStructuralTenXMeanAndP95CaseCount: number
    microsoftExcelLiveStructuralVersion: string
    formulaOfficeListedBreadthPercent: number
    formulaTrackedBreadthPercent: number
    importExportCoveredFeatures: string[]
    importExportFidelityPassed: boolean
    importExportUnsupportedFeatures: string[]
    largeWorkbookSloRowsCovered: number[]
    largeWorkbookSloPassed: boolean
    uiResponsivenessLiveBrowserPassed: boolean
    uiResponsivenessLiveBrowserVendors: string[]
    reliabilityCoveredControls: string[]
    reliabilityPosturePassed: boolean
    reliabilityUncoveredControls: string[]
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

export interface DominanceCompletionAudit {
  allCriteriaPassed: boolean
  unmetRequirements: string[]
  criteria: DominanceCompletionCriterion[]
}

export interface DominanceCompletionCriterion {
  id: string
  requirement: string
  passed: boolean
  evidence: string[]
  gaps: string[]
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
  automationScorecard: AutomationScorecard
  automationScorecardPath: string
  collaborationScorecard: CollaborationScorecard
  collaborationScorecardPath: string
  calculationSemanticsScorecard: CalculationSemanticsScorecard
  calculationSemanticsScorecardPath: string
  competitiveArtifact: CompetitiveArtifact
  competitiveArtifactPath: string
  formulaSnapshot: FormulaDominanceSnapshot
  formulaSnapshotPath: string
  googleSheetsLiveCalculationScorecard: GoogleSheetsLiveCalculationScorecard
  googleSheetsLiveCalculationScorecardPath: string
  googleSheetsLiveRecalculationScorecard: GoogleSheetsLiveRecalculationScorecard
  googleSheetsLiveRecalculationScorecardPath: string
  googleSheetsLiveStructuralScorecard: GoogleSheetsLiveStructuralScorecard
  googleSheetsLiveStructuralScorecardPath: string
  googleSheetsLiveLargeWorkbookScorecard: GoogleSheetsLiveLargeWorkbookScorecard
  googleSheetsLiveLargeWorkbookScorecardPath: string
  microsoftExcelLiveCalculationScorecard: MicrosoftExcelLiveCalculationScorecard
  microsoftExcelLiveCalculationScorecardPath: string
  microsoftExcelLiveRecalculationScorecard: MicrosoftExcelLiveRecalculationScorecard
  microsoftExcelLiveRecalculationScorecardPath: string
  microsoftExcelLiveLargeWorkbookScorecard: MicrosoftExcelLiveLargeWorkbookScorecard
  microsoftExcelLiveLargeWorkbookScorecardPath: string
  microsoftExcelLiveStructuralScorecard: MicrosoftExcelLiveStructuralScorecard
  microsoftExcelLiveStructuralScorecardPath: string
  importExportFidelityScorecard: ImportExportFidelityScorecard
  importExportFidelityScorecardPath: string
  largeWorkbookSloScorecard: LargeWorkbookSloScorecard
  largeWorkbookSloScorecardPath: string
  uiResponsivenessLiveBrowserScorecard: UiResponsivenessLiveBrowserScorecard
  uiResponsivenessLiveBrowserScorecardPath: string
  reliabilityScorecard: ReliabilityScorecard
  reliabilityScorecardPath: string
  securityPostureScorecard: SecurityPostureScorecard
  securityPostureScorecardPath: string
  surfaceSnapshot: HyperFormulaSurfaceSnapshot
  surfaceSnapshotPath: string
}
