#!/usr/bin/env bun

import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { performance } from 'node:perf_hooks'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { SpreadsheetEngine } from '@bilig/core'
import {
  applyWorkbookAgentCommandBundle,
  buildWorkbookAgentPreview,
  createWorkbookAgentCommandBundle,
  isWorkbookAgentCommandBundle,
  isWorkbookAgentToolName,
  normalizeWorkbookAgentToolName,
  projectWorkbookAgentBundle,
  WORKBOOK_AGENT_TOOL_NAMES,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewSummary,
} from '@bilig/agent-api'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { createMemoryWorkbookLocalStoreFactory } from '@bilig/storage-browser'
import {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  WorkPaper,
  type WorkPaperCellAddress,
} from '@bilig/headless'
import { WorkbookWorkerRuntime } from '../apps/web/src/worker-runtime.js'
import { arrayField, asObject, booleanField, literalField, numberField, stringArrayField, stringField } from './json-scorecard-helpers.ts'

export interface AutomationControl {
  readonly id: string
  readonly category: 'semantic-command-api' | 'headless-service-api' | 'worker-runtime-api' | 'tool-registry' | 'workflow-benchmark'
  readonly required: boolean
  readonly passed: boolean
  readonly coveredControls: string[]
  readonly evidence: string
  readonly findings: string[]
}

export interface AutomationScorecard {
  readonly schemaVersion: 1
  readonly suite: 'automation-api-extensibility'
  readonly generatedAt: string
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-automation-scorecard.ts'
    readonly commandBundleImplementation: 'packages/agent-api/src/workbook-agent-bundles.ts'
    readonly previewImplementation: 'packages/agent-api/src/workbook-agent-preview.ts'
    readonly headlessImplementation: 'packages/headless/src/work-paper-runtime.ts'
    readonly workerRuntimeImplementation: 'apps/web/src/worker-runtime.ts'
    readonly toolRegistryImplementation: 'packages/agent-api/src/workbook-agent-tool-names.ts'
  }
  readonly summary: {
    readonly allRequiredControlsPassed: boolean
    readonly semanticCommandWorkflowPassed: boolean
    readonly headlessServiceWorkflowPassed: boolean
    readonly workerPreviewWorkflowPassed: boolean
    readonly toolRegistryPassed: boolean
    readonly tenXWorkflowAutomationBenchmarkPassed: boolean
    readonly registeredToolCount: number
    readonly semanticCommandKindCount: number
    readonly coveredControls: string[]
    readonly uncoveredControls: string[]
    readonly externalGoogleSheetsEvidence: 'not-captured'
    readonly externalMicrosoftExcelEvidence: 'not-captured'
  }
  readonly controls: AutomationControl[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'automation-scorecard.json')
const requiredControlIds = [
  'semantic-command-bundle-preview-apply',
  'headless-service-automation-workflow',
  'worker-runtime-agent-preview',
  'agent-tool-registry-semantic-coverage',
  'semantic-workflow-automation-ten-x-benchmark',
] as const
const coveredControlOrder = [
  'agent.semanticBundleValidation',
  'agent.previewApplyExecution',
  'agent.partialCommandProjection',
  'headless.serviceWorkflow',
  'headless.persistenceRoundTrip',
  'headless.undoRedoAutomation',
  'worker.runtimePreview',
  'tools.semanticWorkbookRegistry',
  'tools.legacyNameNormalization',
  'automation.tenXWorkflowBenchmark',
] as const
const uncoveredControls = ['googleAppsScriptDirectComparison', 'officeScriptsDirectComparison'] as const
const requiredSemanticToolNames = [
  WORKBOOK_AGENT_TOOL_NAMES.applyAndVerify,
  WORKBOOK_AGENT_TOOL_NAMES.undoWorkbookMutation,
  WORKBOOK_AGENT_TOOL_NAMES.writeRange,
  WORKBOOK_AGENT_TOOL_NAMES.setFormula,
  WORKBOOK_AGENT_TOOL_NAMES.formatRange,
  WORKBOOK_AGENT_TOOL_NAMES.copyRange,
  WORKBOOK_AGENT_TOOL_NAMES.moveRange,
  WORKBOOK_AGENT_TOOL_NAMES.insertRows,
  WORKBOOK_AGENT_TOOL_NAMES.setFreezePane,
  WORKBOOK_AGENT_TOOL_NAMES.updateColumnMetadata,
] as const

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error('Automation scorecard is missing. Run: bun scripts/gen-automation-scorecard.ts')
    }
    const scorecard = parseAutomationScorecard(JSON.parse(readFileSync(outputPath, 'utf8')) as unknown)
    validateAutomationScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = await buildAutomationScorecard()
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export async function buildAutomationScorecard(generatedAt = new Date().toISOString()): Promise<AutomationScorecard> {
  const controls = [
    await buildSemanticCommandWorkflowControl(),
    buildHeadlessServiceWorkflowControl(),
    await buildWorkerRuntimePreviewControl(),
    buildToolRegistryControl(),
    await buildSemanticWorkflowAutomationBenchmarkControl(),
  ]
  const coveredControlSet = new Set(controls.flatMap((control) => control.coveredControls))
  const coveredControls = coveredControlOrder.filter((control) => coveredControlSet.has(control))

  return {
    schemaVersion: 1,
    suite: 'automation-api-extensibility',
    generatedAt,
    source: {
      artifactGenerator: 'scripts/gen-automation-scorecard.ts',
      commandBundleImplementation: 'packages/agent-api/src/workbook-agent-bundles.ts',
      previewImplementation: 'packages/agent-api/src/workbook-agent-preview.ts',
      headlessImplementation: 'packages/headless/src/work-paper-runtime.ts',
      workerRuntimeImplementation: 'apps/web/src/worker-runtime.ts',
      toolRegistryImplementation: 'packages/agent-api/src/workbook-agent-tool-names.ts',
    },
    summary: {
      allRequiredControlsPassed: controls.filter((control) => control.required).every((control) => control.passed),
      semanticCommandWorkflowPassed: requiredControl(controls, 'semantic-command-bundle-preview-apply').passed,
      headlessServiceWorkflowPassed: requiredControl(controls, 'headless-service-automation-workflow').passed,
      workerPreviewWorkflowPassed: requiredControl(controls, 'worker-runtime-agent-preview').passed,
      toolRegistryPassed: requiredControl(controls, 'agent-tool-registry-semantic-coverage').passed,
      tenXWorkflowAutomationBenchmarkPassed: requiredControl(controls, 'semantic-workflow-automation-ten-x-benchmark').passed,
      registeredToolCount: Object.values(WORKBOOK_AGENT_TOOL_NAMES).length,
      semanticCommandKindCount: new Set(createSemanticWorkflowCommands().map((command) => command.kind)).size,
      coveredControls,
      uncoveredControls: [...uncoveredControls],
      externalGoogleSheetsEvidence: 'not-captured',
      externalMicrosoftExcelEvidence: 'not-captured',
    },
    controls,
  }
}

async function buildSemanticCommandWorkflowControl(): Promise<AutomationControl> {
  const baseEngine = new SpreadsheetEngine({
    workbookName: 'Automation Workbook',
    replicaId: 'automation:base',
  })
  await baseEngine.ready()
  baseEngine.createSheet('Sheet1')
  baseEngine.setCellValue('Sheet1', 'A1', 10)
  baseEngine.setCellValue('Sheet1', 'B1', 20)
  const snapshot = baseEngine.exportSnapshot()
  const bundle = createSemanticWorkflowBundle()
  const projection = projectWorkbookAgentBundle({
    bundle,
    commandIndexes: [0, 1],
    bundleId: 'automation-semantic-bundle-projected',
  })
  const preview = await buildWorkbookAgentPreview({
    snapshot,
    replicaId: 'automation:preview',
    bundle,
  })
  const applyEngine = new SpreadsheetEngine({
    workbookName: snapshot.workbook.name,
    replicaId: 'automation:apply',
  })
  await applyEngine.ready()
  applyEngine.importSnapshot(snapshot)
  applyWorkbookAgentCommandBundle(applyEngine, bundle)
  const afterSnapshot = applyEngine.exportSnapshot()
  const bundleValidated = isWorkbookAgentCommandBundle(bundle)
  const projectionPassed =
    projection !== null &&
    projection.commands.length === 2 &&
    projection.affectedRanges.length === 2 &&
    projection.estimatedAffectedCells === 3
  const previewApplyPassed =
    preview.effectSummary.displayedCellDiffCount >= 3 &&
    preview.effectSummary.formulaChangeCount >= 2 &&
    previewDiffsMatchEngine(preview, applyEngine) &&
    cellNumber(applyEngine, 'Sheet1', 'C2') === 32 &&
    hasSheetFreeze(afterSnapshot, 'Sheet1', 1, 1) &&
    hasColumnWidth(afterSnapshot, 'Sheet1', 1, 124)

  return automationControl({
    id: 'semantic-command-bundle-preview-apply',
    category: 'semantic-command-api',
    passed: bundleValidated && projectionPassed && previewApplyPassed,
    coveredControls: ['agent.semanticBundleValidation', 'agent.previewApplyExecution', 'agent.partialCommandProjection'],
    evidence:
      `Created a typed workbook-agent bundle with ${String(bundle.commands.length)} semantic commands, ` +
      `validated it, projected a two-command subset, previewed it from a snapshot, applied it to a real SpreadsheetEngine, ` +
      `and verified formula output C2=32 plus freeze-pane and column metadata effects.`,
    findings: [
      ...(bundleValidated ? [] : ['semantic command bundle failed runtime validation']),
      ...(projectionPassed ? [] : [`partial command projection produced an unexpected shape: ${JSON.stringify(projection)}`]),
      ...(previewApplyPassed
        ? []
        : [
            `preview/apply workflow did not produce expected effects: preview=${JSON.stringify(preview.effectSummary)}, C2=${String(
              cellNumber(applyEngine, 'Sheet1', 'C2'),
            )}`,
          ]),
    ],
  })
}

function buildHeadlessServiceWorkflowControl(): AutomationControl {
  const workbook = WorkPaper.buildFromSheets(
    {
      Revenue: [[125], [250], [375], ['=SUM(A1:A3)']],
    },
    {
      maxRows: 1_000,
      maxColumns: 50,
      useColumnIndex: true,
    },
  )
  const revenueId = workbook.getSheetId('Revenue')
  if (revenueId === undefined) {
    return failedAutomationControl('headless-service-automation-workflow', 'headless-service-api', 'Revenue sheet was not created')
  }
  const at = (row: number, col: number): WorkPaperCellAddress => ({ sheet: revenueId, row, col })
  workbook.copy({
    start: at(0, 0),
    end: at(2, 0),
  })
  const batchChanges = workbook.batch(() => {
    workbook.paste(at(0, 1))
    workbook.setCellContents(at(3, 1), '=SUM(B1:B3)')
  })
  const calculatedBeforePersistence = workPaperNumber(workbook, at(3, 1)) === 750
  const document = exportWorkPaperDocument(workbook, { includeConfig: true })
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(document)))
  const restoredRevenueId = restored.getSheetId('Revenue')
  const restoredAt = (row: number, col: number): WorkPaperCellAddress => ({ sheet: restoredRevenueId ?? -1, row, col })
  const persistencePassed = restoredRevenueId !== undefined && workPaperNumber(restored, restoredAt(3, 1)) === 750
  workbook.undo()
  const undoPassed = workbook.getCellSerialized(at(3, 1)) === null
  workbook.redo()
  const redoPassed = workbook.getCellFormula(at(3, 1)) === '=SUM(B1:B3)' && workPaperNumber(workbook, at(3, 1)) === 750
  const serviceWorkflowPassed = batchChanges.length > 0 && calculatedBeforePersistence

  return automationControl({
    id: 'headless-service-automation-workflow',
    category: 'headless-service-api',
    passed: serviceWorkflowPassed && persistencePassed && undoPassed && redoPassed,
    coveredControls: ['headless.serviceWorkflow', 'headless.persistenceRoundTrip', 'headless.undoRedoAutomation'],
    evidence:
      'Executed the production WorkPaper headless API through copy, paste, batch formula write, persistence round trip, undo, and redo without browser DOM automation.',
    findings: [
      ...(serviceWorkflowPassed ? [] : ['headless batch workflow did not calculate the expected 750 total']),
      ...(persistencePassed ? [] : ['headless persistence round trip did not preserve the calculated total']),
      ...(undoPassed ? [] : ['headless undo did not clear the automated formula write']),
      ...(redoPassed ? [] : ['headless redo did not restore the automated formula write']),
    ],
  })
}

async function buildWorkerRuntimePreviewControl(): Promise<AutomationControl> {
  const runtime = new WorkbookWorkerRuntime({
    localStoreFactory: createMemoryWorkbookLocalStoreFactory(),
  })
  await runtime.bootstrap({
    documentId: 'automation-worker-doc',
    replicaId: 'automation:worker',
    persistState: true,
  })
  const bundle = createWorkbookAgentCommandBundle({
    bundleId: 'automation-worker-preview-bundle',
    documentId: 'automation-worker-doc',
    threadId: 'automation-worker-thread',
    turnId: 'automation-worker-turn',
    goalText: 'Preview worker automation',
    baseRevision: 0,
    context: {
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
        range: {
          startAddress: 'A1',
          endAddress: 'B1',
        },
      },
      viewport: {
        rowStart: 0,
        rowEnd: 20,
        colStart: 0,
        colEnd: 8,
      },
    },
    commands: [
      {
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        values: [[9, { formula: '=A1*4' }]],
      },
    ],
    now: 1,
  })
  const preview = await runtime.previewAgentCommandBundle(bundle)
  runtime.dispose()
  const workerPreviewPassed =
    isWorkbookAgentCommandBundle(bundle) &&
    bundle.scope === 'selection' &&
    preview.effectSummary.displayedCellDiffCount >= 1 &&
    preview.effectSummary.formulaChangeCount === 1 &&
    preview.ranges.some((range) => range.sheetName === 'Sheet1' && range.startAddress === 'A1' && range.endAddress === 'B1')

  return automationControl({
    id: 'worker-runtime-agent-preview',
    category: 'worker-runtime-api',
    passed: workerPreviewPassed,
    coveredControls: ['worker.runtimePreview'],
    evidence:
      'Bootstrapped the browser worker runtime against an in-memory local store and previewed a workbook-agent bundle through the worker API.',
    findings: workerPreviewPassed
      ? []
      : [`worker preview produced unexpected bundle or effect summary: ${JSON.stringify(preview.effectSummary)}`],
  })
}

function buildToolRegistryControl(): AutomationControl {
  const toolNames = Object.values(WORKBOOK_AGENT_TOOL_NAMES)
  const uniqueToolNames = new Set(toolNames)
  const allNamesCanonical = toolNames.every((toolName) => /^[a-z0-9_]+$/u.test(toolName) && !toolName.startsWith('bilig'))
  const semanticToolsRegistered = requiredSemanticToolNames.every((toolName) => isWorkbookAgentToolName(toolName))
  const legacyAliasesNormalize =
    normalizeWorkbookAgentToolName('bilig_write_range') === WORKBOOK_AGENT_TOOL_NAMES.writeRange &&
    normalizeWorkbookAgentToolName('bilig.apply_and_verify_workbook_mutation') === WORKBOOK_AGENT_TOOL_NAMES.applyAndVerify &&
    normalizeWorkbookAgentToolName('bilig_update_conditional_format') === WORKBOOK_AGENT_TOOL_NAMES.updateConditionalFormat
  const registryPassed = uniqueToolNames.size === toolNames.length && allNamesCanonical && semanticToolsRegistered && legacyAliasesNormalize

  return automationControl({
    id: 'agent-tool-registry-semantic-coverage',
    category: 'tool-registry',
    passed: registryPassed,
    coveredControls: ['tools.semanticWorkbookRegistry', 'tools.legacyNameNormalization'],
    evidence: `Checked ${String(toolNames.length)} canonical workbook-agent tool names, including semantic mutation, inspection, audit, object, protection, and media tools.`,
    findings: [
      ...(uniqueToolNames.size === toolNames.length ? [] : ['workbook-agent tool registry contains duplicate names']),
      ...(allNamesCanonical ? [] : ['workbook-agent tool registry contains non-canonical or prefixed names']),
      ...(semanticToolsRegistered ? [] : ['required semantic workbook automation tools are missing from the registry']),
      ...(legacyAliasesNormalize ? [] : ['legacy workbook-agent tool aliases do not normalize to canonical names']),
    ],
  })
}

async function buildSemanticWorkflowAutomationBenchmarkControl(): Promise<AutomationControl> {
  const rowCount = 200
  const semanticEngine = await createAutomationBenchmarkEngine('automation:semantic-benchmark')
  const scriptEngine = await createAutomationBenchmarkEngine('automation:script-baseline')
  const bundle = createAutomationBenchmarkBundle(rowCount)
  const semanticCommandCount = bundle.commands.length
  const scriptEquivalentOperationCount = countIncumbentStyleWorkflowOperations(rowCount)

  const semanticStart = performance.now()
  const semanticPreview = await buildWorkbookAgentPreview({
    snapshot: semanticEngine.exportSnapshot(),
    replicaId: 'automation:semantic-benchmark-preview',
    bundle,
  })
  applyWorkbookAgentCommandBundle(semanticEngine, bundle)
  const semanticElapsedMs = performance.now() - semanticStart

  const scriptStart = performance.now()
  applyIncumbentStyleScriptWorkflow(scriptEngine, rowCount)
  const scriptElapsedMs = performance.now() - scriptStart

  const hostCallReductionRatio = scriptEquivalentOperationCount / semanticCommandCount
  const finalRow = rowCount + 1
  const expectedFinalValue = rowCount * 3 * 1.2
  const semanticSnapshot = semanticEngine.exportSnapshot()
  const scriptSnapshot = scriptEngine.exportSnapshot()
  const semanticFinalValue = cellNumber(semanticEngine, 'Sheet1', `D${String(finalRow)}`)
  const scriptFinalValue = cellNumber(scriptEngine, 'Sheet1', `D${String(finalRow)}`)
  const semanticOutputPassed =
    semanticPreview.effectSummary.displayedCellDiffCount > 0 &&
    semanticFinalValue === expectedFinalValue &&
    hasSheetFreeze(semanticSnapshot, 'Sheet1', 1, 1) &&
    hasColumnWidth(semanticSnapshot, 'Sheet1', 0, 128)
  const scriptOutputPassed =
    scriptFinalValue === expectedFinalValue &&
    hasSheetFreeze(scriptSnapshot, 'Sheet1', 1, 1) &&
    hasColumnWidth(scriptSnapshot, 'Sheet1', 0, 128)
  const tenXHostCallReductionPassed = hostCallReductionRatio >= 10

  return automationControl({
    id: 'semantic-workflow-automation-ten-x-benchmark',
    category: 'workflow-benchmark',
    passed: tenXHostCallReductionPassed && semanticOutputPassed && scriptOutputPassed,
    coveredControls: ['automation.tenXWorkflowBenchmark'],
    evidence:
      `Executed a ${String(rowCount)}-row workflow as ${String(semanticCommandCount)} typed semantic commands versus ` +
      `${String(scriptEquivalentOperationCount)} incumbent-style row-by-row script operations, for ` +
      `${formatBenchmarkRatio(hostCallReductionRatio)}x fewer script-visible host calls. ` +
      `Semantic preview+apply took ${formatBenchmarkMs(semanticElapsedMs)}ms; the local script-style baseline took ` +
      `${formatBenchmarkMs(scriptElapsedMs)}ms. Both paths produced Sheet1!D${String(finalRow)}=${String(expectedFinalValue)}.`,
    findings: [
      ...(tenXHostCallReductionPassed
        ? []
        : [`semantic host-call reduction ratio was ${formatBenchmarkRatio(hostCallReductionRatio)}x, below the 10x threshold`]),
      ...(semanticOutputPassed
        ? []
        : [
            `semantic benchmark workflow produced unexpected output: D${String(finalRow)}=${String(semanticFinalValue)}, previewDiffs=${String(
              semanticPreview.effectSummary.displayedCellDiffCount,
            )}`,
          ]),
      ...(scriptOutputPassed ? [] : [`script-style baseline produced unexpected output: D${String(finalRow)}=${String(scriptFinalValue)}`]),
    ],
  })
}

function createSemanticWorkflowBundle(): WorkbookAgentCommandBundle {
  return createWorkbookAgentCommandBundle({
    bundleId: 'automation-semantic-bundle',
    documentId: 'automation-doc',
    threadId: 'automation-thread',
    turnId: 'automation-turn',
    goalText: 'Run semantic workbook automation workflow',
    baseRevision: 1,
    context: {
      selection: {
        sheetName: 'Sheet1',
        address: 'A2',
        range: {
          startAddress: 'A2',
          endAddress: 'C2',
        },
      },
      viewport: {
        rowStart: 0,
        rowEnd: 20,
        colStart: 0,
        colEnd: 10,
      },
    },
    commands: createSemanticWorkflowCommands(),
    now: 1,
  })
}

function createSemanticWorkflowCommands(): WorkbookAgentCommand[] {
  return [
    {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'A2',
      values: [[2, { formula: '=A1+B1' }]],
    },
    {
      kind: 'setRangeFormulas',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'C2',
        endAddress: 'C2',
      },
      formulas: [['=SUM(A2:B2)']],
    },
    {
      kind: 'formatRange',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A2',
        endAddress: 'C2',
      },
      patch: {
        font: {
          bold: true,
        },
      },
      numberFormat: 'currency',
    },
    {
      kind: 'copyRange',
      source: {
        sheetName: 'Sheet1',
        startAddress: 'A2',
        endAddress: 'C2',
      },
      target: {
        sheetName: 'Sheet1',
        startAddress: 'A3',
        endAddress: 'C3',
      },
    },
    {
      kind: 'setFreezePane',
      sheetName: 'Sheet1',
      rows: 1,
      cols: 1,
    },
    {
      kind: 'updateColumnMetadata',
      sheetName: 'Sheet1',
      startCol: 1,
      count: 1,
      width: 124,
    },
  ]
}

async function createAutomationBenchmarkEngine(replicaId: string): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'Automation Benchmark Workbook',
    replicaId,
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  return engine
}

function createAutomationBenchmarkBundle(rowCount: number): WorkbookAgentCommandBundle {
  return createWorkbookAgentCommandBundle({
    bundleId: 'automation-ten-x-workflow-benchmark',
    documentId: 'automation-benchmark-doc',
    threadId: 'automation-benchmark-thread',
    turnId: 'automation-benchmark-turn',
    goalText: 'Benchmark semantic workflow automation against row-by-row scripting',
    baseRevision: 1,
    context: {
      selection: {
        sheetName: 'Sheet1',
        address: 'A2',
        range: {
          startAddress: 'A2',
          endAddress: `D${String(rowCount + 1)}`,
        },
      },
      viewport: {
        rowStart: 0,
        rowEnd: rowCount + 2,
        colStart: 0,
        colEnd: 4,
      },
    },
    commands: createAutomationBenchmarkCommands(rowCount),
    now: 1,
  })
}

function createAutomationBenchmarkCommands(rowCount: number): WorkbookAgentCommand[] {
  return [
    {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'A2',
      values: Array.from({ length: rowCount }, (_, index) => {
        const value = index + 1
        return [value, value * 2]
      }),
    },
    {
      kind: 'setRangeFormulas',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'C2',
        endAddress: `D${String(rowCount + 1)}`,
      },
      formulas: Array.from({ length: rowCount }, (_, index) => {
        const row = index + 2
        return [`=A${String(row)}+B${String(row)}`, `=C${String(row)}*1.2`]
      }),
    },
    {
      kind: 'formatRange',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A2',
        endAddress: `D${String(rowCount + 1)}`,
      },
      patch: {
        font: {
          bold: true,
        },
      },
      numberFormat: 'currency',
    },
    {
      kind: 'setFreezePane',
      sheetName: 'Sheet1',
      rows: 1,
      cols: 1,
    },
    {
      kind: 'updateColumnMetadata',
      sheetName: 'Sheet1',
      startCol: 0,
      count: 4,
      width: 128,
    },
  ]
}

function countIncumbentStyleWorkflowOperations(rowCount: number): number {
  const perRowValueCalls = 2
  const perRowFormulaCalls = 2
  const perRowFormatCalls = 1
  const perRowNumberFormatCalls = 1
  const freezePaneCall = 1
  const columnMetadataCall = 1
  return (
    rowCount * (perRowValueCalls + perRowFormulaCalls + perRowFormatCalls + perRowNumberFormatCalls) + freezePaneCall + columnMetadataCall
  )
}

function applyIncumbentStyleScriptWorkflow(engine: SpreadsheetEngine, rowCount: number): void {
  Array.from({ length: rowCount }, (_, index) => index + 2).forEach((row) => {
    const value = row - 1
    engine.setCellValue('Sheet1', `A${String(row)}`, value)
    engine.setCellValue('Sheet1', `B${String(row)}`, value * 2)
    engine.setCellFormula('Sheet1', `C${String(row)}`, `A${String(row)}+B${String(row)}`)
    engine.setCellFormula('Sheet1', `D${String(row)}`, `C${String(row)}*1.2`)
    engine.setRangeStyle(
      {
        sheetName: 'Sheet1',
        startAddress: `A${String(row)}`,
        endAddress: `D${String(row)}`,
      },
      {
        font: {
          bold: true,
        },
      },
    )
    engine.setRangeNumberFormat(
      {
        sheetName: 'Sheet1',
        startAddress: `A${String(row)}`,
        endAddress: `D${String(row)}`,
      },
      'currency',
    )
  })
  engine.setFreezePane('Sheet1', 1, 1)
  engine.updateColumnMetadata('Sheet1', 0, 4, 128, null)
}

function formatBenchmarkRatio(value: number): string {
  return value.toFixed(1)
}

function formatBenchmarkMs(value: number): string {
  return value.toFixed(3)
}

function previewDiffsMatchEngine(preview: WorkbookAgentPreviewSummary, engine: SpreadsheetEngine): boolean {
  return preview.cellDiffs.every((diff) => {
    const cell = engine.getCell(diff.sheetName, diff.address)
    return normalizePreviewFormula(cell.formula) === diff.afterFormula && normalizePreviewInput(cell.input) === diff.afterInput
  })
}

function normalizePreviewFormula(formula: string | undefined): string | null {
  return formula === undefined ? null : `=${formula}`
}

function normalizePreviewInput(input: string | number | boolean | null | undefined): string | number | boolean | null {
  return input === undefined ? null : input
}

function cellNumber(engine: SpreadsheetEngine, sheetName: string, address: string): number | null {
  const value = engine.getCell(sheetName, address).value
  return value.tag === ValueTag.Number ? value.value : null
}

function workPaperNumber(workbook: WorkPaper, address: WorkPaperCellAddress): number | null {
  const value = workbook.getCellValue(address)
  return value.tag === ValueTag.Number ? value.value : null
}

function hasSheetFreeze(snapshot: WorkbookSnapshot, sheetName: string, rows: number, cols: number): boolean {
  const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
  return sheet?.metadata.freezePane?.rows === rows && sheet.metadata.freezePane.cols === cols
}

function hasColumnWidth(snapshot: WorkbookSnapshot, sheetName: string, columnIndex: number, width: number): boolean {
  const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
  return sheet?.metadata.columns.some((entry) => entry.index === columnIndex && entry.size === width) ?? false
}

function automationControl(input: {
  readonly id: AutomationControl['id']
  readonly category: AutomationControl['category']
  readonly passed: boolean
  readonly coveredControls: readonly string[]
  readonly evidence: string
  readonly findings: readonly string[]
}): AutomationControl {
  return {
    id: input.id,
    category: input.category,
    required: true,
    passed: input.passed,
    coveredControls: [...input.coveredControls],
    evidence: input.evidence,
    findings: [...input.findings],
  }
}

function failedAutomationControl(id: AutomationControl['id'], category: AutomationControl['category'], finding: string): AutomationControl {
  return automationControl({
    id,
    category,
    passed: false,
    coveredControls: [],
    evidence: 'Control did not execute because setup failed.',
    findings: [finding],
  })
}

function requiredControl(controls: readonly AutomationControl[], id: string): AutomationControl {
  const entry = controls.find((control) => control.id === id)
  if (!entry) {
    throw new Error(`Automation scorecard is missing required control: ${id}`)
  }
  return entry
}

export function parseAutomationScorecard(value: unknown): AutomationScorecard {
  const record = asObject(value, 'automation scorecard')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'automation-api-extensibility') {
    throw new Error('Unexpected automation scorecard header')
  }
  const source = asObject(record['source'], 'automation source')
  const summary = asObject(record['summary'], 'automation summary')
  return {
    schemaVersion: 1,
    suite: 'automation-api-extensibility',
    generatedAt: stringField(record, 'generatedAt'),
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-automation-scorecard.ts'),
      commandBundleImplementation: literalField(source, 'commandBundleImplementation', 'packages/agent-api/src/workbook-agent-bundles.ts'),
      previewImplementation: literalField(source, 'previewImplementation', 'packages/agent-api/src/workbook-agent-preview.ts'),
      headlessImplementation: literalField(source, 'headlessImplementation', 'packages/headless/src/work-paper-runtime.ts'),
      workerRuntimeImplementation: literalField(source, 'workerRuntimeImplementation', 'apps/web/src/worker-runtime.ts'),
      toolRegistryImplementation: literalField(source, 'toolRegistryImplementation', 'packages/agent-api/src/workbook-agent-tool-names.ts'),
    },
    summary: {
      allRequiredControlsPassed: booleanField(summary, 'allRequiredControlsPassed'),
      semanticCommandWorkflowPassed: booleanField(summary, 'semanticCommandWorkflowPassed'),
      headlessServiceWorkflowPassed: booleanField(summary, 'headlessServiceWorkflowPassed'),
      workerPreviewWorkflowPassed: booleanField(summary, 'workerPreviewWorkflowPassed'),
      toolRegistryPassed: booleanField(summary, 'toolRegistryPassed'),
      tenXWorkflowAutomationBenchmarkPassed: booleanField(summary, 'tenXWorkflowAutomationBenchmarkPassed'),
      registeredToolCount: numberField(summary, 'registeredToolCount'),
      semanticCommandKindCount: numberField(summary, 'semanticCommandKindCount'),
      coveredControls: stringArrayField(summary, 'coveredControls'),
      uncoveredControls: stringArrayField(summary, 'uncoveredControls'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'not-captured'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'not-captured'),
    },
    controls: arrayField(record, 'controls').map(parseAutomationControl),
  }
}

function parseAutomationControl(value: unknown): AutomationControl {
  const record = asObject(value, 'automation control')
  return {
    id: stringField(record, 'id'),
    category: parseAutomationCategory(stringField(record, 'category')),
    required: booleanField(record, 'required'),
    passed: booleanField(record, 'passed'),
    coveredControls: stringArrayField(record, 'coveredControls'),
    evidence: stringField(record, 'evidence'),
    findings: stringArrayField(record, 'findings'),
  }
}

export function validateAutomationScorecard(scorecard: AutomationScorecard): void {
  for (const id of requiredControlIds) {
    const control = requiredControl(scorecard.controls, id)
    if (!control.required) {
      throw new Error(`Automation scorecard required control is not marked required: ${id}`)
    }
    if (!control.passed) {
      throw new Error(`Automation scorecard contains a failed required control: ${id}`)
    }
  }
  if (!scorecard.summary.allRequiredControlsPassed) {
    throw new Error('Automation scorecard summary reports failed required controls')
  }
  for (const control of coveredControlOrder) {
    if (!scorecard.summary.coveredControls.includes(control)) {
      throw new Error(`Automation scorecard is missing covered control: ${control}`)
    }
  }
  for (const control of uncoveredControls) {
    if (!scorecard.summary.uncoveredControls.includes(control)) {
      throw new Error(`Automation scorecard is missing uncovered control disclosure: ${control}`)
    }
  }
}

function parseAutomationCategory(value: string): AutomationControl['category'] {
  if (
    value === 'semantic-command-api' ||
    value === 'headless-service-api' ||
    value === 'worker-runtime-api' ||
    value === 'tool-registry' ||
    value === 'workflow-benchmark'
  ) {
    return value
  }
  throw new Error(`Unexpected automation category: ${value}`)
}

function logResult(mode: 'check' | 'write', scorecard: AutomationScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        allRequiredControlsPassed: scorecard.summary.allRequiredControlsPassed,
        coveredControls: scorecard.summary.coveredControls.length,
        uncoveredControls: scorecard.summary.uncoveredControls.length,
        registeredToolCount: scorecard.summary.registeredToolCount,
      },
      null,
      2,
    ),
  )
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'automation-scorecard-'))
  const tempFilePath = join(tempDir, 'scorecard.json')
  writeFileSync(tempFilePath, serializedJson)
  const oxfmtPath = join(rootDir, 'node_modules', '.bin', 'oxfmt')

  const formatResult = Bun.spawnSync([oxfmtPath, '--write', tempFilePath], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (formatResult.exitCode !== 0) {
    const stderr = new TextDecoder().decode(formatResult.stderr).trim()
    rmSync(tempDir, { recursive: true, force: true })
    throw new Error(`Unable to format generated automation scorecard: ${stderr}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
