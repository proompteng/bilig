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
    expect(scorecard.summary.microsoftExcelLiveCalculationPassed).toBe(true)
    expect(scorecard.summary.microsoftExcelLiveCalculationCaseCount).toBe(2)
    expect(scorecard.summary.microsoftExcelLiveCalculationEvidence).toBe('live-local-microsoft-excel-automation')
    expect(scorecard.summary.microsoftExcelLiveStructuralPassed).toBe(true)
    expect(scorecard.summary.microsoftExcelLiveStructuralCaseCount).toBe(6)
    expect(scorecard.summary.microsoftExcelLiveStructuralTenXMeanAndP95CaseCount).toBe(6)
    expect(scorecard.summary.microsoftExcelLiveStructuralEvidence).toBe('live-local-microsoft-excel-automation')
    expect(scorecard.summary.importExportFidelityPassed).toBe(true)
    expect(scorecard.summary.largeWorkbookSloRowsCovered).toEqual([100_000, 250_000])
    expect(scorecard.summary.largeWorkbookSloPassed).toBe(true)
    expect(scorecard.summary.auditabilityPosturePassed).toBe(true)
    expect(scorecard.summary.automationPosturePassed).toBe(true)
    expect(scorecard.summary.collaborationPosturePassed).toBe(true)
    expect(scorecard.summary.reliabilityPosturePassed).toBe(true)
    expect(scorecard.summary.securityPosturePassed).toBe(true)
    expect(scorecard.sourceArtifacts.auditabilityScorecard).toBe('packages/benchmarks/baselines/auditability-scorecard.json')
    expect(scorecard.sourceArtifacts.automationScorecard).toBe('packages/benchmarks/baselines/automation-scorecard.json')
    expect(scorecard.sourceArtifacts.collaborationScorecard).toBe('packages/benchmarks/baselines/collaboration-scorecard.json')
    expect(scorecard.sourceArtifacts.microsoftExcelLiveCalculationScorecard).toBe(
      'packages/benchmarks/baselines/microsoft-excel-live-calculation-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.microsoftExcelLiveStructuralScorecard).toBe(
      'packages/benchmarks/baselines/microsoft-excel-live-structural-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.reliabilityScorecard).toBe('packages/benchmarks/baselines/reliability-scorecard.json')
    expect(scorecard.sourceArtifacts.importExportFidelityScorecard).toBe(
      'packages/benchmarks/baselines/import-export-fidelity-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.largeWorkbookSloScorecard).toBe('packages/benchmarks/baselines/large-workbook-slo-scorecard.json')
    expect(scorecard.sourceArtifacts.securityPostureScorecard).toBe('packages/benchmarks/baselines/security-posture-scorecard.json')
    expect(scorecard.categories.find((category) => category.id === 'calculation-correctness')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/formula/src/__tests__/fixtures/formula-dominance-snapshot.json',
        'packages/benchmarks/baselines/microsoft-excel-live-calculation-scorecard.json',
      ]),
      blockers: [
        '2 Office-listed functions are still missing from the runtime inventory',
        'no generated scorecard currently compares committed semantics directly against live Google Sheets',
        'live Microsoft Excel calculation scorecard covers representative required cases, not all committed formula semantics',
      ],
    })
    expect(scorecard.categories.find((category) => category.id === 'import-export-compatibility')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/import-export-fidelity-scorecard.json',
        'packages/benchmarks/baselines/import-export-external-sheets-excel-comparison.json',
      ]),
      blockers: ['generated XLSX round-trip evidence covers supported snapshot semantics, not full native Excel macro execution semantics'],
    })
    expect(scorecard.categories.find((category) => category.id === 'large-workbook-scale')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/large-workbook-slo-scorecard.json',
        'packages/benchmarks/baselines/large-workbook-external-sheets-excel-comparison.json',
        'e2e/tests/web-shell-scroll-performance.pw.ts',
      ]),
      blockers: ['no direct Sheets or Excel large-workbook live timing artifact exists in the repo'],
    })
    expect(scorecard.categories.find((category) => category.id === 'structural-edit-performance')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
        'packages/benchmarks/baselines/microsoft-excel-live-structural-scorecard.json',
      ]),
      checkCommands: expect.arrayContaining(['pnpm structural:excel-live:check']),
      blockers: [
        'structural rows and columns lead HyperFormula, but the worst ratios are not 10x',
        'no direct Google Sheets structural-edit timing artifact exists in the repo',
      ],
    })
    expect(scorecard.categories.find((category) => category.id === 'ui-responsiveness')).toMatchObject({
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/large-workbook-slo-scorecard.json',
        'packages/benchmarks/baselines/ui-responsiveness-external-sheets-excel-comparison.json',
      ]),
      blockers: ['no direct Sheets or Excel browser responsiveness live timing artifact exists in the repo'],
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
    expect(calculation?.blockers).toContain(
      'no generated scorecard currently compares committed semantics directly against live Google Sheets',
    )
    expect(calculation?.blockers).toContain(
      'live Microsoft Excel calculation scorecard covers representative required cases, not all committed formula semantics',
    )
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
    expect(packageJson).toContain('"calculation:excel-live:check": "bun scripts/gen-microsoft-excel-live-calculation-scorecard.ts --check"')
    expect(packageJson).toContain('"structural:excel-live:check": "bun scripts/gen-microsoft-excel-live-structural-scorecard.ts --check"')
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
    expect(packageJson).toContain('"security:posture:check": "bun scripts/gen-security-posture-scorecard.ts --check"')
    expect(runCi).toContain("pnpm('bilig dominance scorecard check', 'dominance:check')")
    expect(runCi).toContain("pnpm('Microsoft Excel live calculation scorecard check', 'calculation:excel-live:check')")
    expect(runCi).toContain("pnpm('Microsoft Excel live structural scorecard check', 'structural:excel-live:check')")
    expect(runCi).toContain("pnpm('auditability scorecard check', 'auditability:check')")
    expect(runCi).toContain("pnpm('reliability scorecard check', 'reliability:check')")
    expect(runCi).toContain("pnpm('collaboration scorecard check', 'collaboration:check')")
    expect(runCi).toContain("pnpm('automation scorecard check', 'automation:check')")
    expect(runCi).toContain("pnpm('import/export fidelity scorecard check', 'import-export:fidelity:check')")
    expect(runCi).toContain("pnpm('large workbook SLO scorecard check', 'large-workbook:slo:check')")
    expect(runCi).toContain("pnpm('security posture scorecard check', 'security:posture:check')")
  })
})

function parseGeneratedScorecard(source: string): { categories: Array<{ id: unknown; blockers: string[] }> } {
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

function buildFixtureInput(): BuildScorecardInput {
  return {
    competitiveArtifactPath: 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
    formulaSnapshotPath: 'packages/formula/src/__tests__/fixtures/formula-dominance-snapshot.json',
    microsoftExcelLiveCalculationScorecardPath: 'packages/benchmarks/baselines/microsoft-excel-live-calculation-scorecard.json',
    microsoftExcelLiveStructuralScorecardPath: 'packages/benchmarks/baselines/microsoft-excel-live-structural-scorecard.json',
    auditabilityScorecardPath: 'packages/benchmarks/baselines/auditability-scorecard.json',
    automationScorecardPath: 'packages/benchmarks/baselines/automation-scorecard.json',
    collaborationScorecardPath: 'packages/benchmarks/baselines/collaboration-scorecard.json',
    importExportFidelityScorecardPath: 'packages/benchmarks/baselines/import-export-fidelity-scorecard.json',
    largeWorkbookSloScorecardPath: 'packages/benchmarks/baselines/large-workbook-slo-scorecard.json',
    reliabilityScorecardPath: 'packages/benchmarks/baselines/reliability-scorecard.json',
    securityPostureScorecardPath: 'packages/benchmarks/baselines/security-posture-scorecard.json',
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
    microsoftExcelLiveCalculationScorecard: {
      schemaVersion: 1,
      suite: 'microsoft-excel-live-calculation-correctness',
      generatedAt: '2026-05-06T09:00:00.000Z',
      host: {
        arch: 'arm64',
        platform: 'darwin',
      },
      source: {
        artifactGenerator: 'scripts/gen-microsoft-excel-live-calculation-scorecard.ts',
        implementationPackage: 'packages/headless',
        evidenceKind: 'live-local-microsoft-excel-automation',
        appleScriptTransport: 'osascript',
      },
      microsoftExcel: {
        appPath: '/Applications/Microsoft Excel.app',
        version: '16.test',
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 2,
        matchingCaseCount: 2,
        coveredFeatures: ['excelLive.arithmeticPrecedence', 'excelLive.aggregateSumRange'],
        googleSheetsEvidence: 'not-covered-by-this-artifact',
      },
      cases: [
        {
          id: 'arithmetic-precedence',
          formula: '=A2+B2*2',
          formulaCell: 'D2',
          coveredFeature: 'excelLive.arithmeticPrecedence',
          biligValue: 24,
          microsoftExcelRawValue: '24.0',
          microsoftExcelValue: 24,
          passed: true,
        },
        {
          id: 'aggregate-sum-range',
          formula: '=SUM(A3:C3)',
          formulaCell: 'D3',
          coveredFeature: 'excelLive.aggregateSumRange',
          biligValue: 6,
          microsoftExcelRawValue: '6.0',
          microsoftExcelValue: 6,
          passed: true,
        },
      ],
    },
    microsoftExcelLiveStructuralScorecard: {
      schemaVersion: 1,
      suite: 'microsoft-excel-live-structural-performance',
      generatedAt: '2026-05-06T15:00:00.000Z',
      host: {
        arch: 'arm64',
        platform: 'darwin',
      },
      source: {
        artifactGenerator: 'scripts/gen-microsoft-excel-live-structural-scorecard.ts',
        implementationPackage: 'packages/headless',
        evidenceKind: 'live-local-microsoft-excel-automation',
        appleScriptTransport: 'osascript',
      },
      benchmark: {
        rowCount: 500,
        sampleCount: 5,
        screenUpdating: false,
      },
      microsoftExcel: {
        appPath: '/Applications/Microsoft Excel.app',
        version: '16.test',
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 6,
        tenXMeanAndP95CaseCount: 6,
        workpaperWins: 6,
        coveredOperations: ['insert-rows', 'delete-rows', 'move-rows', 'insert-columns', 'delete-columns', 'move-columns'],
        googleSheetsEvidence: 'not-covered-by-this-artifact',
      },
      cases: [
        structuralCase('excel-live-structural-insert-rows', 'insert-rows', 'rows'),
        structuralCase('excel-live-structural-delete-rows', 'delete-rows', 'rows'),
        structuralCase('excel-live-structural-move-rows', 'move-rows', 'rows'),
        structuralCase('excel-live-structural-insert-columns', 'insert-columns', 'columns'),
        structuralCase('excel-live-structural-delete-columns', 'delete-columns', 'columns'),
        structuralCase('excel-live-structural-move-columns', 'move-columns', 'columns'),
      ],
    },
    importExportFidelityScorecard: {
      schemaVersion: 1,
      suite: 'import-export-fidelity',
      generatedAt: '2026-05-06T08:00:00.000Z',
      source: {
        artifactGenerator: 'scripts/gen-import-export-fidelity-scorecard.ts',
        implementationPackage: 'packages/excel-import',
        enginePackage: 'packages/core',
        externalImportExportComparisonArtifact: 'packages/benchmarks/baselines/import-export-external-sheets-excel-comparison.json',
      },
      summary: {
        allRequiredCasesPassed: true,
        csvRoundTripPassed: true,
        xlsxImportPassed: true,
        xlsxSnapshotRoundTripPassed: true,
        coveredFeatures: [
          'csv.import',
          'xlsx.import',
          'xlsx.export',
          'external.googleSheetsImportExportDocs',
          'external.microsoftExcelImportExportDocs',
          'external.sheetsExcelImportExportComparison',
        ],
        unsupportedFeatures: ['xlsx.styles.export'],
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
      cases: [
        {
          id: 'xlsx-snapshot-roundtrip-values-formulas-formats',
          format: 'xlsx',
          direction: 'export-import',
          required: true,
          passed: true,
          coveredFeatures: ['xlsx.values', 'xlsx.formulas', 'xlsx.numberFormats'],
          missingFeatures: [],
          evidence: 'fixture round-tripped',
        },
      ],
    },
    auditabilityScorecard: {
      schemaVersion: 1,
      suite: 'auditability-posture',
      generatedAt: '2026-05-06T10:00:00.000Z',
      source: {
        artifactGenerator: 'scripts/gen-auditability-scorecard.ts',
        previewImplementation: 'packages/agent-api/src/workbook-agent-preview.ts',
        applyImplementation: 'apps/bilig/src/zero/workbook-agent-apply.ts',
        authoritativeApplyImplementation: 'apps/bilig/src/zero/service.ts',
        historyImplementation: 'packages/zero-sync/src/workbook-history-state.ts',
        headedBrowserAuditabilityTestFile: 'e2e/tests/web-shell-remote-sync.pw.ts',
        externalAuditabilityComparisonArtifact: 'packages/benchmarks/baselines/auditability-external-sheets-excel-comparison.json',
      },
      summary: {
        allRequiredControlsPassed: true,
        previewApplyParityPassed: true,
        applyUndoRoundTripPassed: true,
        authoritativeApplyGuardPassed: true,
        historyRevertRedoPassed: true,
        headedBrowserRevertFlowPassed: true,
        coveredControls: [
          'agent.previewDiffParity',
          'agent.applyCapturesUndoBundle',
          'headedBrowser.previewApplyRevertFlow',
          'external.googleSheetsAuditabilityDocs',
          'external.microsoftExcelAuditabilityDocs',
          'external.sheetsExcelAuditabilityComparison',
        ],
        uncoveredControls: [],
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
      controls: [
        {
          id: 'agent-preview-apply-parity',
          category: 'preview-apply',
          required: true,
          passed: true,
          coveredControls: ['agent.previewDiffParity'],
          evidence: 'fixture preview matched apply',
          findings: [],
        },
      ],
    },
    automationScorecard: {
      schemaVersion: 1,
      suite: 'automation-api-extensibility',
      generatedAt: '2026-05-06T12:00:00.000Z',
      source: {
        artifactGenerator: 'scripts/gen-automation-scorecard.ts',
        commandBundleImplementation: 'packages/agent-api/src/workbook-agent-bundles.ts',
        previewImplementation: 'packages/agent-api/src/workbook-agent-preview.ts',
        headlessImplementation: 'packages/headless/src/work-paper-runtime.ts',
        workerRuntimeImplementation: 'apps/web/src/worker-runtime.ts',
        toolRegistryImplementation: 'packages/agent-api/src/workbook-agent-tool-names.ts',
        externalAutomationComparisonArtifact: 'packages/benchmarks/baselines/automation-external-sheets-excel-comparison.json',
      },
      summary: {
        allRequiredControlsPassed: true,
        semanticCommandWorkflowPassed: true,
        headlessServiceWorkflowPassed: true,
        workerPreviewWorkflowPassed: true,
        toolRegistryPassed: true,
        tenXWorkflowAutomationBenchmarkPassed: true,
        registeredToolCount: 98,
        semanticCommandKindCount: 6,
        coveredControls: [
          'agent.semanticBundleValidation',
          'headless.serviceWorkflow',
          'automation.tenXWorkflowBenchmark',
          'googleAppsScriptDirectComparison',
          'officeScriptsDirectComparison',
        ],
        uncoveredControls: [],
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
      controls: [
        {
          id: 'semantic-command-bundle-preview-apply',
          category: 'semantic-command-api',
          required: true,
          passed: true,
          coveredControls: ['agent.semanticBundleValidation'],
          evidence: 'fixture bundle previewed and applied',
          findings: [],
        },
      ],
    },
    collaborationScorecard: {
      schemaVersion: 1,
      suite: 'collaboration-posture',
      generatedAt: '2026-05-06T13:00:00.000Z',
      source: {
        artifactGenerator: 'scripts/gen-collaboration-scorecard.ts',
        workerRuntimeImplementation: 'apps/web/src/worker-runtime.ts',
        presenceImplementation: 'apps/web/src/workbook-presence-model.ts',
        presenceSessionImplementation: 'apps/bilig/src/workbook-runtime/document-presence-session-store.ts',
        viewportPatchImplementation: 'apps/web/src/projected-viewport-patch-application.ts',
        editorConflictImplementation: 'apps/web/src/use-workbook-editor-conflict.tsx',
        headedBrowserViewportTestFile: 'e2e/tests/web-shell-scroll-performance.pw.ts',
        externalCollaborationComparisonArtifact: 'packages/benchmarks/baselines/collaboration-external-sheets-excel-comparison.json',
      },
      summary: {
        allRequiredControlsPassed: true,
        syncRebaseAckPassed: true,
        presenceSelectionPassed: true,
        conflictViewportPassed: true,
        headedBrowserViewportPassed: true,
        longRunningConflictRatePassed: true,
        coveredControls: [
          'sync.pendingRebase',
          'presence.sessionLifecycle',
          'headedBrowser.multiUserViewportSoak',
          'conflict.longRunningZeroUnexpectedConflicts',
          'external.googleSheetsCollaborationDocs',
          'external.microsoftExcelCollaborationDocs',
          'external.sheetsExcelCollaborationComparison',
        ],
        uncoveredControls: [],
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
      controls: [
        {
          id: 'worker-sync-rebase-ack-roundtrip',
          category: 'local-first-sync',
          required: true,
          passed: true,
          coveredControls: ['sync.pendingRebase'],
          evidence: 'fixture pending mutations rebased and acked',
          findings: [],
        },
      ],
    },
    reliabilityScorecard: {
      schemaVersion: 1,
      suite: 'reliability-posture',
      generatedAt: '2026-05-06T11:00:00.000Z',
      source: {
        artifactGenerator: 'scripts/gen-reliability-scorecard.ts',
        workerRuntimeImplementation: 'apps/web/src/worker-runtime.ts',
        mutationJournalImplementation: 'apps/web/src/worker-runtime-mutation-journal.ts',
        localStoreImplementation: 'packages/storage-browser/src/index.ts',
        headedBrowserReliabilityTestFile: 'e2e/tests/web-shell-remote-sync.pw.ts',
        externalReliabilityComparisonArtifact: 'packages/benchmarks/baselines/reliability-external-sheets-excel-comparison.json',
      },
      summary: {
        allRequiredControlsPassed: true,
        pendingReloadPassed: true,
        authoritativeAckPassed: true,
        authoritativeRebasePassed: true,
        failedRetryPassed: true,
        headedBrowserReloadPassed: true,
        headedBrowserCrashSoakPassed: true,
        offlineNetworkPartitionPassed: true,
        coveredControls: [
          'pending.localReloadSurvival',
          'localStore.journalActiveView',
          'headedBrowser.reloadPersistence',
          'headedBrowser.crashSoak',
          'offline.networkPartitionRecoverySoak',
          'external.googleSheetsReliabilityDocs',
          'external.microsoftExcelReliabilityDocs',
          'external.sheetsExcelReliabilityComparison',
        ],
        uncoveredControls: [],
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
      controls: [
        {
          id: 'pending-mutations-survive-reload',
          category: 'pending-durability',
          required: true,
          passed: true,
          coveredControls: ['pending.localReloadSurvival'],
          evidence: 'fixture pending mutation survived reload',
          findings: [],
        },
      ],
    },
    securityPostureScorecard: {
      schemaVersion: 1,
      suite: 'security-posture',
      generatedAt: '2026-05-06T09:00:00.000Z',
      source: {
        artifactGenerator: 'scripts/gen-security-posture-scorecard.ts',
        formulaRuntimeScanRoots: ['packages/formula/src'],
        importImplementation: 'packages/excel-import/src/index.ts',
        agentPolicyImplementation: 'packages/agent-api/src/workbook-agent-execution-policy.ts',
        browserSecurityHeadersImplementation: 'apps/bilig/src/http/sync-server-security-headers.ts',
        externalSecurityComparisonArtifact: 'packages/benchmarks/baselines/security-external-sheets-excel-comparison.json',
        dependencyAuditCommand: 'pnpm audit --prod --json',
        runtimePackageGate: 'pnpm publish:runtime:check',
      },
      summary: {
        allRequiredControlsPassed: true,
        formulaSandboxPassed: true,
        importSafetyPassed: true,
        agentPermissionPolicyPassed: true,
        runtimePackageHardeningPassed: true,
        browserCspPassed: true,
        dependencyAuditPassed: true,
        coveredControls: [
          'formula.noEval',
          'xlsx.macroWarning',
          'browser.contentSecurityPolicy',
          'dependency.vulnerabilityAudit',
          'external.googleSheetsSecurityDocs',
          'external.microsoftExcelSecurityDocs',
          'external.sheetsExcelSecurityComparison',
        ],
        uncoveredControls: ['deployment.runtimeNetworkPolicy'],
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
      controls: [
        {
          id: 'formula-runtime-no-dynamic-code-execution',
          category: 'formula-sandbox',
          required: true,
          passed: true,
          coveredControls: ['formula.noEval'],
          evidence: 'fixture scan passed',
          findings: [],
        },
      ],
    },
    largeWorkbookSloScorecard: {
      schemaVersion: 1,
      suite: 'large-workbook-slo',
      generatedAt: '2026-05-06T16:00:00.000Z',
      source: {
        benchmarkCommand: 'CI=1 pnpm bench:contracts',
        benchmarkScript: 'scripts/bench-contracts.ts',
        headedBrowserCommand: 'pnpm test:browser:full',
        headedBrowserTestFile: 'e2e/tests/web-shell-scroll-performance.pw.ts',
        artifactGenerator: 'scripts/gen-large-workbook-slo-scorecard.ts',
        externalLargeWorkbookComparisonArtifact: 'packages/benchmarks/baselines/large-workbook-external-sheets-excel-comparison.json',
        externalUiResponsivenessComparisonArtifact: 'packages/benchmarks/baselines/ui-responsiveness-external-sheets-excel-comparison.json',
      },
      summary: {
        coveredLargeWorkbookRows: [100_000, 250_000],
        allSloBudgetsPassed: true,
        allGateBudgetsPassed: true,
        headedBrowserFrameP95Evidence: 'playwright-contracts',
        headedBrowserFrameP95ContractsPassed: true,
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
        externalUiResponsivenessGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalUiResponsivenessMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
      measurements: [
        sloMeasurement('load100k', 'large-workbook-scale', 100_000, 230, 1500),
        sloMeasurement('load250k', 'large-workbook-scale', 250_000, 600, 1500),
        sloMeasurement('workerWarmStart100k', 'large-workbook-scale', 100_000, 12, 500),
        sloMeasurement('workerWarmStart250k', 'large-workbook-scale', 250_000, 17, 700),
        sloMeasurement('workerVisibleEdit10k', 'ui-responsiveness', 10_000, 4, 16),
        sloMeasurement('workerReconnectCatchUp100Pending', 'collaboration', 10_000, 270, 2000),
      ],
      headedBrowserFrameP95Contracts: [
        headedBrowserContract('headedDense100kDiagonalBrowse', 'large-workbook-scale', 100_000, 'dense-mixed-100k', 'frameMs.p95', 20),
        headedBrowserContract('headedWide250kMainBodyBrowse', 'large-workbook-scale', 250_000, 'wide-mixed-250k', 'frameMs.p95', 20),
        headedBrowserContract(
          'headedWide250kVisibleEditCommit',
          'ui-responsiveness',
          250_000,
          'wide-mixed-250k',
          'mutationToVisibleMs.p95',
          50,
        ),
      ],
      externalSheetsExcelComparison: {
        artifact: 'packages/benchmarks/baselines/large-workbook-external-sheets-excel-comparison.json',
        sourceBasis: 'official-public-docs-reviewed-2026-05-06',
        officialGoogleSheetsSourceCount: 5,
        officialMicrosoftExcelSourceCount: 2,
        requiredDimensionsPassed: true,
        coveredFeatures: [
          'external.googleSheetsLargeWorkbookDocs',
          'external.microsoftExcelLargeWorkbookDocs',
          'external.sheetsExcelLargeWorkbookScaleComparison',
        ],
        limitations: ['Official-docs comparison, not live timing.'],
        findings: [],
      },
      uiResponsivenessExternalSheetsExcelComparison: {
        artifact: 'packages/benchmarks/baselines/ui-responsiveness-external-sheets-excel-comparison.json',
        sourceBasis: 'official-public-docs-reviewed-2026-05-06',
        officialGoogleSheetsSourceCount: 3,
        officialMicrosoftExcelSourceCount: 4,
        requiredDimensionsPassed: true,
        coveredFeatures: [
          'external.googleSheetsUiResponsivenessDocs',
          'external.microsoftExcelUiResponsivenessDocs',
          'external.sheetsExcelUiResponsivenessComparison',
        ],
        limitations: ['Official-docs comparison, not live timing.'],
        findings: [],
      },
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

function headedBrowserContract(
  id: string,
  category: BuildScorecardInput['largeWorkbookSloScorecard']['headedBrowserFrameP95Contracts'][number]['category'],
  materializedCells: number,
  corpusCaseId: string,
  metric: BuildScorecardInput['largeWorkbookSloScorecard']['headedBrowserFrameP95Contracts'][number]['metric'],
  budgetP95: number,
): BuildScorecardInput['largeWorkbookSloScorecard']['headedBrowserFrameP95Contracts'][number] {
  return {
    id,
    category,
    label: id,
    materializedCells,
    corpusCaseId,
    metric,
    budgetP95,
    minSampleCount: metric === 'frameMs.p95' ? 120 : 1,
    playwrightTestFile: 'e2e/tests/web-shell-scroll-performance.pw.ts',
    playwrightArtifactFile: `${id}.json`,
    command: 'pnpm test:browser:full',
    passed: true,
    findings: [],
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

function structuralCase(
  id: string,
  operation: BuildScorecardInput['microsoftExcelLiveStructuralScorecard']['cases'][number]['operation'],
  axis: BuildScorecardInput['microsoftExcelLiveStructuralScorecard']['cases'][number]['axis'],
): BuildScorecardInput['microsoftExcelLiveStructuralScorecard']['cases'][number] {
  return {
    id,
    operation,
    axis,
    rowCount: 500,
    sampleCount: 5,
    workpaperElapsedMs: numericSummary(1),
    microsoftExcelElapsedMs: numericSummary(20),
    workpaperToMicrosoftExcelMeanRatio: 0.05,
    workpaperToMicrosoftExcelP95Ratio: 0.05,
    tenXMeanAndP95: true,
    verification: {
      workpaper: {
        height: 500,
        width: axis === 'columns' ? 4 : 2,
        value: 500,
      },
      microsoftExcel: {
        height: 500,
        width: axis === 'columns' ? 4 : 2,
        value: 500,
      },
      equivalent: true,
    },
    passed: true,
  }
}

function numericSummary(
  value: number,
): BuildScorecardInput['microsoftExcelLiveStructuralScorecard']['cases'][number]['workpaperElapsedMs'] {
  return {
    samples: [value, value, value, value, value],
    min: value,
    median: value,
    p95: value,
    max: value,
    mean: value,
    standardDeviation: 0,
    relativeStandardDeviation: 0,
    standardError: 0,
    confidence95: {
      low: value,
      high: value,
    },
  }
}
