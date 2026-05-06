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
    expect(scorecard.sourceArtifacts.reliabilityScorecard).toBe('packages/benchmarks/baselines/reliability-scorecard.json')
    expect(scorecard.sourceArtifacts.importExportFidelityScorecard).toBe(
      'packages/benchmarks/baselines/import-export-fidelity-scorecard.json',
    )
    expect(scorecard.sourceArtifacts.largeWorkbookSloScorecard).toBe('packages/benchmarks/baselines/large-workbook-slo-scorecard.json')
    expect(scorecard.sourceArtifacts.securityPostureScorecard).toBe('packages/benchmarks/baselines/security-posture-scorecard.json')
    expect(scorecard.categories.find((category) => category.id === 'import-export-compatibility')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining(['packages/benchmarks/baselines/import-export-fidelity-scorecard.json']),
    })
    expect(scorecard.categories.find((category) => category.id === 'large-workbook-scale')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/large-workbook-slo-scorecard.json',
        'e2e/tests/web-shell-scroll-performance.pw.ts',
      ]),
      blockers: ['no direct Sheets or Excel large-workbook scale artifact exists in the repo'],
    })
    expect(scorecard.categories.find((category) => category.id === 'ui-responsiveness')).toMatchObject({
      blockers: ['no direct Sheets or Excel browser responsiveness comparison artifact exists in the repo'],
    })
    expect(scorecard.categories.find((category) => category.id === 'collaboration')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/collaboration-scorecard.json',
        'e2e/tests/web-shell-scroll-performance.pw.ts',
      ]),
      blockers: [
        'generated collaboration evidence still leaves uncovered controls: externalSheetsCollaborationComparison',
        'no direct Sheets collaboration comparison artifact exists in the repo',
      ],
    })
    expect(scorecard.categories.find((category) => category.id === 'auditability')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/auditability-scorecard.json',
        'e2e/tests/web-shell-remote-sync.pw.ts',
      ]),
      blockers: ['no direct incumbent auditability comparison artifact exists in the repo'],
    })
    expect(scorecard.categories.find((category) => category.id === 'automation-api-extensibility')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining(['packages/benchmarks/baselines/automation-scorecard.json']),
      blockers: [
        'generated automation evidence still leaves uncovered controls: googleAppsScriptDirectComparison, officeScriptsDirectComparison',
        'no direct generated Google Apps Script or Office Scripts execution comparison exists',
      ],
    })
    expect(scorecard.categories.find((category) => category.id === 'reliability')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining([
        'packages/benchmarks/baselines/reliability-scorecard.json',
        'e2e/tests/web-shell-remote-sync.pw.ts',
      ]),
      blockers: ['no direct Sheets or Excel reliability comparison artifact exists in the repo'],
    })
    expect(scorecard.categories.find((category) => category.id === 'security')).toMatchObject({
      status: 'partial-repo-evidence',
      evidenceArtifacts: expect.arrayContaining(['packages/benchmarks/baselines/security-posture-scorecard.json']),
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
    expect(runCi).toContain("pnpm('auditability scorecard check', 'auditability:check')")
    expect(runCi).toContain("pnpm('reliability scorecard check', 'reliability:check')")
    expect(runCi).toContain("pnpm('collaboration scorecard check', 'collaboration:check')")
    expect(runCi).toContain("pnpm('automation scorecard check', 'automation:check')")
    expect(runCi).toContain("pnpm('import/export fidelity scorecard check', 'import-export:fidelity:check')")
    expect(runCi).toContain("pnpm('large workbook SLO scorecard check', 'large-workbook:slo:check')")
    expect(runCi).toContain("pnpm('security posture scorecard check', 'security:posture:check')")
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
    importExportFidelityScorecard: {
      schemaVersion: 1,
      suite: 'import-export-fidelity',
      summary: {
        allRequiredCasesPassed: true,
        csvRoundTripPassed: true,
        xlsxImportPassed: true,
        xlsxSnapshotRoundTripPassed: true,
        coveredFeatures: ['csv.import', 'xlsx.import', 'xlsx.export'],
        unsupportedFeatures: ['xlsx.styles.export'],
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
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
      },
      summary: {
        allRequiredControlsPassed: true,
        previewApplyParityPassed: true,
        applyUndoRoundTripPassed: true,
        authoritativeApplyGuardPassed: true,
        historyRevertRedoPassed: true,
        headedBrowserRevertFlowPassed: true,
        coveredControls: ['agent.previewDiffParity', 'agent.applyCapturesUndoBundle', 'headedBrowser.previewApplyRevertFlow'],
        uncoveredControls: ['externalSheetsExcelAuditabilityComparison'],
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
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
        coveredControls: ['agent.semanticBundleValidation', 'headless.serviceWorkflow', 'automation.tenXWorkflowBenchmark'],
        uncoveredControls: ['googleAppsScriptDirectComparison', 'officeScriptsDirectComparison'],
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
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
        ],
        uncoveredControls: ['externalSheetsCollaborationComparison'],
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
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
        ],
        uncoveredControls: ['externalSheetsExcelReliabilityComparison'],
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
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
        coveredControls: ['formula.noEval', 'xlsx.macroWarning', 'browser.contentSecurityPolicy', 'dependency.vulnerabilityAudit'],
        uncoveredControls: ['deployment.runtimeNetworkPolicy'],
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
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
      summary: {
        coveredLargeWorkbookRows: [100_000, 250_000],
        allSloBudgetsPassed: true,
        allGateBudgetsPassed: true,
        headedBrowserFrameP95Evidence: 'playwright-contracts',
        headedBrowserFrameP95ContractsPassed: true,
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
