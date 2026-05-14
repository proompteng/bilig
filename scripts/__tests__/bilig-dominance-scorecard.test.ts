import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildBiligDominanceScorecard } from '../gen-bilig-dominance-scorecard.ts'
import { buildFixtureInput } from './bilig-dominance-scorecard.fixture.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('bilig dominance scorecard', () => {
  it('keeps the active Sheets/Excel goal separated from narrower HyperFormula evidence', () => {
    const scorecard = buildBiligDominanceScorecard(buildFixtureInput())

    expect(scorecard.goalStatus).toBe('active-not-achieved')
    expect(scorecard.overallGoogleSheets10xStatus).toMatchObject({
      passed: false,
      status: 'blocked',
      unmetRequirements: [
        'ui-responsiveness: live UI browser evidence is not a same-corpus 10x proof against Google Sheets with rendered-grid proof',
      ],
    })
    expect(scorecard.overallGoogleSheets10xStatus.categories.map((entry) => entry.id)).toEqual([
      'recalculation-speed',
      'structural-edit-performance',
      'large-workbook-scale',
      'ui-responsiveness',
    ])
    expect(scorecard.claimPolicy.blanketTenXClaimAllowed).toBe(false)
    expect(scorecard.claimPolicy.unmetRequirements).toEqual(
      expect.arrayContaining([
        expect.stringContaining('calculation-correctness:'),
        expect.stringContaining('large-workbook-scale:'),
        'ui-responsiveness: live UI browser evidence is not a same-corpus 10x proof against incumbents',
        expect.stringContaining('security:'),
      ]),
    )
    expect(scorecard.completionAudit.allCriteriaPassed).toBe(false)
    expect(scorecard.completionAudit.criteria.map((entry) => entry.id)).toEqual([
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
    expect(scorecard.completionAudit.criteria.find((entry) => entry.id === 'import-export-compatibility')).toMatchObject({
      passed: true,
      gaps: [],
      evidence: expect.arrayContaining(['declined unsafe runtime features: xlsx.macros.execution']),
    })
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
    expect(scorecard.summary.googleSheetsLiveCalculationPassed).toBe(true)
    expect(scorecard.summary.googleSheetsLiveCalculationCaseCount).toBe(2)
    expect(scorecard.summary.googleSheetsLiveCalculationEvidence).toBe('live-google-sheets-native-conversion-via-google-drive-connector')
    expect(scorecard.summary.googleSheetsLiveCalculationSpreadsheetId).toBe('google-sheet-test-id')
    expect(scorecard.summary.googleSheetsLiveRecalculationPassed).toBe(true)
    expect(scorecard.summary.googleSheetsLiveRecalculationCaseCount).toBe(4)
    expect(scorecard.summary.googleSheetsLiveRecalculationTenXMeanAndP95CaseCount).toBe(4)
    expect(scorecard.summary.googleSheetsLiveRecalculationEvidence).toBe('live-google-sheets-native-conversion-via-google-drive-connector')
    expect(scorecard.summary.googleSheetsLiveStructuralPassed).toBe(true)
    expect(scorecard.summary.googleSheetsLiveStructuralCaseCount).toBe(6)
    expect(scorecard.summary.googleSheetsLiveStructuralTenXMeanAndP95CaseCount).toBe(6)
    expect(scorecard.summary.googleSheetsLiveStructuralEvidence).toBe('live-google-sheets-native-conversion-via-google-drive-connector')
    expect(scorecard.summary.googleSheetsLiveLargeWorkbookPassed).toBe(true)
    expect(scorecard.summary.googleSheetsLiveLargeWorkbookCaseCount).toBe(2)
    expect(scorecard.summary.googleSheetsLiveLargeWorkbookTenXMeanAndP95CaseCount).toBe(2)
    expect(scorecard.summary.googleSheetsLiveLargeWorkbookEvidence).toBe('live-google-sheets-native-conversion-via-google-drive-connector')
    expect(scorecard.summary.googleSheetsLiveLargeWorkbookSpreadsheetIds).toEqual([
      'google-sheets-100k-sample-0',
      'google-sheets-100k-sample-1',
      'google-sheets-100k-sample-2',
      'google-sheets-250k-sample-0',
      'google-sheets-250k-sample-1',
      'google-sheets-250k-sample-2',
    ])
    expect(scorecard.summary.microsoftExcelLiveCalculationPassed).toBe(true)
    expect(scorecard.summary.microsoftExcelLiveCalculationCaseCount).toBe(2)
    expect(scorecard.summary.microsoftExcelLiveCalculationEvidence).toBe('live-local-microsoft-excel-automation')
    expect(scorecard.summary.microsoftExcelLiveRecalculationPassed).toBe(true)
    expect(scorecard.summary.microsoftExcelLiveRecalculationCaseCount).toBe(4)
    expect(scorecard.summary.microsoftExcelLiveRecalculationTenXMeanAndP95CaseCount).toBe(4)
    expect(scorecard.summary.microsoftExcelLiveRecalculationEvidence).toBe('live-local-microsoft-excel-automation')
    expect(scorecard.summary.microsoftExcelLiveLargeWorkbookPassed).toBe(true)
    expect(scorecard.summary.microsoftExcelLiveLargeWorkbookCaseCount).toBe(2)
    expect(scorecard.summary.microsoftExcelLiveLargeWorkbookTenXMeanAndP95CaseCount).toBe(1)
    expect(scorecard.summary.microsoftExcelLiveLargeWorkbookEvidence).toBe('live-local-microsoft-excel-automation')
    expect(scorecard.summary.microsoftExcelLiveStructuralPassed).toBe(true)
    expect(scorecard.summary.microsoftExcelLiveStructuralCaseCount).toBe(6)
    expect(scorecard.summary.microsoftExcelLiveStructuralTenXMeanAndP95CaseCount).toBe(6)
    expect(scorecard.summary.microsoftExcelLiveStructuralEvidence).toBe('live-local-microsoft-excel-automation')
    expect(scorecard.summary.importExportFidelityPassed).toBe(true)
    expect(scorecard.summary.importExportUnsupportedFeatures).toEqual([])
    expect(scorecard.summary.importExportDeclinedRuntimeFeatures).toEqual(['xlsx.macros.execution'])
    expect(scorecard.summary.largeWorkbookSloRowsCovered).toEqual([100_000, 250_000])
    expect(scorecard.summary.largeWorkbookSloPassed).toBe(true)
    expect(scorecard.summary.uiResponsivenessLiveBrowserPassed).toBe(true)
    expect(scorecard.summary.uiResponsivenessLiveBrowserVendors).toEqual(['google-sheets', 'microsoft-excel-web'])
    expect(scorecard.summary.auditabilityPosturePassed).toBe(true)
    expect(scorecard.summary.automationPosturePassed).toBe(true)
    expect(scorecard.summary.collaborationPosturePassed).toBe(true)
    expect(scorecard.summary.calculationSemanticsPassed).toBe(true)
    expect(scorecard.summary.calculationSemanticsCoveredCanonicalFixtureCount).toBe(300)
    expect(scorecard.summary.calculationSemanticsCoveredWorkbookSemanticsFixtureCount).toBe(10)
    expect(scorecard.summary.reliabilityPosturePassed).toBe(true)
    expect(scorecard.summary.securityPosturePassed).toBe(true)
    expect(scorecard.sourceArtifacts.auditabilityScorecard).toBe('packages/benchmarks/baselines/auditability-scorecard.json')
    expect(scorecard.sourceArtifacts.automationScorecard).toBe('packages/benchmarks/baselines/automation-scorecard.json')
    expect(scorecard.sourceArtifacts.calculationSemanticsScorecard).toBe(
      'packages/benchmarks/baselines/calculation-semantics-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.collaborationScorecard).toBe('packages/benchmarks/baselines/collaboration-scorecard.json')
    expect(scorecard.sourceArtifacts.googleSheetsLiveCalculationScorecard).toBe(
      'packages/benchmarks/baselines/google-sheets-live-calculation-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.googleSheetsLiveRecalculationScorecard).toBe(
      'packages/benchmarks/baselines/google-sheets-live-recalculation-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.googleSheetsLiveStructuralScorecard).toBe(
      'packages/benchmarks/baselines/google-sheets-live-structural-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.googleSheetsLiveLargeWorkbookScorecard).toBe(
      'packages/benchmarks/baselines/google-sheets-live-large-workbook-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.microsoftExcelLiveCalculationScorecard).toBe(
      'packages/benchmarks/baselines/microsoft-excel-live-calculation-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.microsoftExcelLiveRecalculationScorecard).toBe(
      'packages/benchmarks/baselines/microsoft-excel-live-recalculation-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.microsoftExcelLiveLargeWorkbookScorecard).toBe(
      'packages/benchmarks/baselines/microsoft-excel-live-large-workbook-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.microsoftExcelLiveStructuralScorecard).toBe(
      'packages/benchmarks/baselines/microsoft-excel-live-structural-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.reliabilityScorecard).toBe('packages/benchmarks/baselines/reliability-scorecard.json')
    expect(scorecard.sourceArtifacts.importExportFidelityScorecard).toBe(
      'packages/benchmarks/baselines/import-export-fidelity-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.largeWorkbookSloScorecard).toBe('packages/benchmarks/baselines/large-workbook-slo-scorecard.json')
    expect(scorecard.sourceArtifacts.uiResponsivenessLiveBrowserScorecard).toBe(
      'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.securityPostureScorecard).toBe('packages/benchmarks/baselines/security-posture-scorecard.json')
    expect(scorecard.categories.find((category) => category.id === 'calculation-correctness')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/calculation-semantics-scorecard.json',
        'packages/formula/src/__tests__/fixtures/formula-dominance-snapshot.json',
        'packages/benchmarks/baselines/google-sheets-live-calculation-scorecard.json',
        'packages/benchmarks/baselines/microsoft-excel-live-calculation-scorecard.json',
      ]),
      checkCommands: expect.arrayContaining(['pnpm calculation:semantics:check', 'pnpm calculation:google-sheets-live:check']),
      blockers: ['2 Office-listed functions are still missing from the runtime inventory'],
    })
    expect(scorecard.categories.find((category) => category.id === 'import-export-compatibility')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/import-export-fidelity-scorecard.json',
        'packages/benchmarks/baselines/import-export-external-sheets-excel-comparison.json',
      ]),
      blockers: [],
    })
    expect(scorecard.categories.find((category) => category.id === 'large-workbook-scale')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/large-workbook-slo-scorecard.json',
        'packages/benchmarks/baselines/google-sheets-live-large-workbook-scorecard.json',
        'packages/benchmarks/baselines/microsoft-excel-live-large-workbook-scorecard.json',
        'packages/benchmarks/baselines/large-workbook-external-sheets-excel-comparison.json',
        'e2e/tests/web-shell-scroll-performance.pw.ts',
      ]),
      checkCommands: expect.arrayContaining(['pnpm large-workbook:excel-live:check', 'pnpm large-workbook:google-sheets-live:check']),
      blockers: ['live Microsoft Excel large-workbook timing scorecard does not prove 10x mean+p95 for all large-workbook cases'],
    })
    expect(scorecard.categories.find((category) => category.id === 'recalculation-speed')).toMatchObject({
      status: 'repo-proved-lead',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
        'packages/benchmarks/baselines/google-sheets-live-recalculation-scorecard.json',
        'packages/benchmarks/baselines/microsoft-excel-live-recalculation-scorecard.json',
      ]),
      checkCommands: expect.arrayContaining(['pnpm recalculation:excel-live:check', 'pnpm recalculation:google-sheets-live:check']),
      blockers: [],
    })
    expect(scorecard.categories.find((category) => category.id === 'structural-edit-performance')).toMatchObject({
      status: 'repo-proved-lead',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
        'packages/benchmarks/baselines/google-sheets-live-structural-scorecard.json',
        'packages/benchmarks/baselines/microsoft-excel-live-structural-scorecard.json',
      ]),
      checkCommands: expect.arrayContaining(['pnpm structural:excel-live:check', 'pnpm structural:google-sheets-live:check']),
      blockers: [],
    })
    expect(scorecard.categories.find((category) => category.id === 'ui-responsiveness')).toMatchObject({
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/large-workbook-slo-scorecard.json',
        'packages/benchmarks/baselines/ui-responsiveness-external-sheets-excel-comparison.json',
        'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json',
      ]),
      checkCommands: expect.arrayContaining([
        'pnpm ui:same-corpus:capture -- --preflight --google-sheets-url <url> --microsoft-excel-web-url <url> [--google-sheets-storage-state <state.json>]',
        'pnpm ui:browser-live:generate -- --capture <capture.json>',
        'pnpm ui:browser-live:check',
      ]),
      blockers: ['live UI browser evidence is direct, but it is not a same-corpus 10x proof against incumbents'],
    })
    expect(scorecard.categories.find((category) => category.id === 'collaboration')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/collaboration-scorecard.json',
        'packages/benchmarks/baselines/collaboration-external-sheets-excel-comparison.json',
        'e2e/tests/web-shell-scroll-performance.pw.ts',
      ]),
      blockers: [],
    })
    expect(scorecard.categories.find((category) => category.id === 'auditability')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/auditability-scorecard.json',
        'packages/benchmarks/baselines/auditability-external-sheets-excel-comparison.json',
        'e2e/tests/web-shell-remote-sync.pw.ts',
      ]),
      blockers: [],
    })
    expect(scorecard.categories.find((category) => category.id === 'automation-api-extensibility')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/automation-scorecard.json',
        'packages/benchmarks/baselines/automation-external-sheets-excel-comparison.json',
      ]),
      blockers: [],
    })
    expect(scorecard.categories.find((category) => category.id === 'reliability')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/reliability-scorecard.json',
        'packages/benchmarks/baselines/reliability-external-sheets-excel-comparison.json',
        'e2e/tests/web-shell-remote-sync.pw.ts',
      ]),
      blockers: [],
    })
    expect(scorecard.categories.find((category) => category.id === 'security')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/security-posture-scorecard.json',
        'packages/benchmarks/baselines/security-external-sheets-excel-comparison.json',
      ]),
      blockers: ['generated security posture evidence has not yet covered deployment runtime network policy'],
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

  it('does not report zero missing Office-listed formulas as a blocker', () => {
    const artifact = parseGeneratedScorecard(
      readFileSync(resolve(repoRoot, 'packages/benchmarks/baselines/bilig-dominance-scorecard.json'), 'utf8'),
    )
    const calculation = artifact.categories.find((category) => category.id === 'calculation-correctness')

    expect(calculation?.blockers).not.toContain('0 Office-listed functions are still missing from the runtime inventory')
    expect(calculation?.blockers).not.toContain(
      'no generated scorecard currently compares committed semantics directly against live Google Sheets',
    )
    expect(calculation?.blockers).not.toContain(
      'live Microsoft Excel calculation scorecard covers representative required cases, not all committed formula semantics',
    )
    expect(calculation?.blockers).not.toContain(
      'live Google Sheets calculation scorecard covers representative required cases, not all committed formula semantics',
    )
  })

  it('does not keep direct incumbent browser UI timing as a blocker after live public-browser evidence exists', () => {
    const artifact = parseGeneratedScorecard(
      readFileSync(resolve(repoRoot, 'packages/benchmarks/baselines/bilig-dominance-scorecard.json'), 'utf8'),
    )
    const uiResponsiveness = artifact.categories.find((category) => category.id === 'ui-responsiveness')

    expect(uiResponsiveness?.blockers).not.toContain(
      'no direct Sheets or Excel browser responsiveness live timing artifact exists in the repo',
    )
    expect(uiResponsiveness?.evidenceArtifacts).toContain('packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')
  })

  it('tracks serialized generated-source CI checks as operator workflow evidence', () => {
    const artifact = parseGeneratedScorecard(
      readFileSync(resolve(repoRoot, 'packages/benchmarks/baselines/bilig-dominance-scorecard.json'), 'utf8'),
    )
    const workflow = artifact.categories.find((category) => category.id === 'operator-developer-workflow')

    expect(workflow?.currentEvidence).toContain(
      'generated-source CI checks are serialized to avoid pnpm workspace-state races in the evidence gate',
    )
    expect(workflow?.currentEvidence).toContain('completion audit criteria passed: false')
    expect(workflow?.blockers).toEqual([])
    expect(workflow?.evidenceArtifacts).toContain('scripts/run-ci.ts')
  })

  it('does not keep deployment network policy as a blocker after security evidence covers it', () => {
    const input = buildFixtureInput()
    input.securityPostureScorecard.summary.coveredControls = [
      ...input.securityPostureScorecard.summary.coveredControls,
      'deployment.runtimeNetworkPolicy',
    ]
    input.securityPostureScorecard.summary.uncoveredControls = []

    const scorecard = buildBiligDominanceScorecard(input)
    const security = scorecard.categories.find((category) => category.id === 'security')

    expect(security?.blockers).toEqual([])
  })

  it('wires the dominance check into fast CI generated checks', () => {
    const packageJson = readFileSync(resolve(repoRoot, 'package.json'), 'utf8')
    const runCi = readFileSync(resolve(repoRoot, 'scripts/run-ci.ts'), 'utf8')

    expect(packageJson).toContain('"dominance:check": "bun scripts/gen-bilig-dominance-scorecard.ts --check"')
    expect(packageJson).toContain('"calculation:semantics:check": "bun scripts/gen-calculation-semantics-scorecard.ts --check"')
    expect(packageJson).toContain('"calculation:excel-live:check": "bun scripts/gen-microsoft-excel-live-calculation-scorecard.ts --check"')
    expect(packageJson).toContain(
      '"calculation:google-sheets-live:check": "bun scripts/gen-google-sheets-live-calculation-scorecard.ts --check"',
    )
    expect(packageJson).toContain(
      '"recalculation:excel-live:check": "bun scripts/gen-microsoft-excel-live-recalculation-scorecard.ts --check"',
    )
    expect(packageJson).toContain(
      '"recalculation:google-sheets-live:check": "bun scripts/gen-google-sheets-live-recalculation-scorecard.ts --check"',
    )
    expect(packageJson).toContain('"structural:excel-live:check": "bun scripts/gen-microsoft-excel-live-structural-scorecard.ts --check"')
    expect(packageJson).toContain(
      '"structural:google-sheets-live:check": "bun scripts/gen-google-sheets-live-structural-scorecard.ts --check"',
    )
    expect(packageJson).toContain(
      '"large-workbook:excel-live:check": "bun scripts/gen-microsoft-excel-live-large-workbook-scorecard.ts --check"',
    )
    expect(packageJson).toContain(
      '"large-workbook:google-sheets-live:check": "bun scripts/gen-google-sheets-live-large-workbook-scorecard.ts --check"',
    )
    expect(packageJson).toContain('"auditability:generate": "bun scripts/gen-auditability-scorecard.ts"')
    expect(packageJson).toContain('"auditability:check": "bun scripts/gen-auditability-scorecard.ts --check"')
    expect(packageJson).toContain('"reliability:generate": "bun scripts/gen-reliability-scorecard.ts"')
    expect(packageJson).toContain('"reliability:check": "bun scripts/gen-reliability-scorecard.ts --check"')
    expect(packageJson).toContain('"collaboration:generate": "bun scripts/gen-collaboration-scorecard.ts"')
    expect(packageJson).toContain('"collaboration:check": "bun scripts/gen-collaboration-scorecard.ts --check"')
    expect(packageJson).toContain('"automation:generate": "bun scripts/gen-automation-scorecard.ts"')
    expect(packageJson).toContain('"automation:check": "bun scripts/gen-automation-scorecard.ts --check"')
    expect(packageJson).toContain('"import-export:fidelity:check": "bun scripts/gen-import-export-fidelity-scorecard.ts --check"')
    expect(packageJson).toContain('"large-workbook:slo:check": "bun scripts/gen-large-workbook-slo-scorecard.ts --check"')
    expect(packageJson).toContain('"ui:same-corpus:capture": "bun scripts/capture-ui-responsiveness-same-corpus.ts"')
    expect(packageJson).toContain('"ui:browser-live:generate": "bun scripts/gen-ui-responsiveness-live-browser-scorecard.ts"')
    expect(packageJson).toContain('"ui:browser-live:check": "bun scripts/gen-ui-responsiveness-live-browser-scorecard.ts --check"')
    expect(packageJson).toContain('"security:posture:check": "bun scripts/gen-security-posture-scorecard.ts --check"')
    expect(runCi).toContain("bunScript('bilig dominance scorecard check', 'scripts/gen-bilig-dominance-scorecard.ts', '--check')")
    expect(runCi).toContain(
      "bunScript('calculation semantics scorecard check', 'scripts/gen-calculation-semantics-scorecard.ts', '--check')",
    )
    expect(runCi).toContain(
      "bunScript('Microsoft Excel live calculation scorecard check', 'scripts/gen-microsoft-excel-live-calculation-scorecard.ts', '--check')",
    )
    expect(runCi).toContain(
      "bunScript('Google Sheets live calculation scorecard check', 'scripts/gen-google-sheets-live-calculation-scorecard.ts', '--check')",
    )
    expect(runCi).toContain(
      "bunScript('Microsoft Excel live recalculation scorecard check', 'scripts/gen-microsoft-excel-live-recalculation-scorecard.ts', '--check')",
    )
    expect(runCi).toContain(
      "bunScript('Google Sheets live recalculation scorecard check', 'scripts/gen-google-sheets-live-recalculation-scorecard.ts', '--check')",
    )
    expect(runCi).toContain(
      "bunScript('Microsoft Excel live structural scorecard check', 'scripts/gen-microsoft-excel-live-structural-scorecard.ts', '--check')",
    )
    expect(runCi).toContain(
      "bunScript('Google Sheets live structural scorecard check', 'scripts/gen-google-sheets-live-structural-scorecard.ts', '--check')",
    )
    expect(runCi).toContain("'Microsoft Excel live large workbook scorecard check'")
    expect(runCi).toContain("'scripts/gen-microsoft-excel-live-large-workbook-scorecard.ts'")
    expect(runCi).toContain(
      "bunScript('Google Sheets live large workbook scorecard check', 'scripts/gen-google-sheets-live-large-workbook-scorecard.ts', '--check')",
    )
    expect(runCi).toContain("bunScript('auditability scorecard check', 'scripts/gen-auditability-scorecard.ts', '--check')")
    expect(runCi).toContain("bunScript('reliability scorecard check', 'scripts/gen-reliability-scorecard.ts', '--check')")
    expect(runCi).toContain("bunScript('collaboration scorecard check', 'scripts/gen-collaboration-scorecard.ts', '--check')")
    expect(runCi).toContain("bunScript('automation scorecard check', 'scripts/gen-automation-scorecard.ts', '--check')")
    expect(runCi).toContain(
      "bunScript('import/export fidelity scorecard check', 'scripts/gen-import-export-fidelity-scorecard.ts', '--check')",
    )
    expect(runCi).toContain("bunScript('large workbook SLO scorecard check', 'scripts/gen-large-workbook-slo-scorecard.ts', '--check')")
    expect(runCi).toContain(
      "bunScript('UI responsiveness live browser scorecard check', 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts', '--check')",
    )
    expect(runCi).toContain("bunScript('security posture scorecard check', 'scripts/gen-security-posture-scorecard.ts', '--check')")
  })
})

function parseGeneratedScorecard(source: string): {
  categories: Array<{ id: unknown; blockers: string[]; currentEvidence: string[]; evidenceArtifacts: string[] }>
} {
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
        blockers: stringList(category['blockers'], 'Generated scorecard category blockers'),
        currentEvidence: stringList(category['currentEvidence'], 'Generated scorecard category evidence'),
        evidenceArtifacts: stringList(category['evidenceArtifacts'], 'Generated scorecard category artifacts'),
      }
    }),
  }
}

function stringList(value: unknown, name: string): string[] {
  if (!Array.isArray(value) || !value.every((entry) => typeof entry === 'string')) {
    throw new Error(`${name} must be a string array`)
  }
  return value
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}
