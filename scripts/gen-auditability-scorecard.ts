#!/usr/bin/env bun

import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { SpreadsheetEngine } from '@bilig/core'
import {
  areWorkbookAgentPreviewSummariesEqual,
  buildWorkbookAgentPreview,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewSummary,
} from '@bilig/agent-api'
import { applyWorkbookAgentCommandBundleWithUndoCapture } from '../apps/bilig/src/zero/workbook-agent-apply.js'
import type { CellSnapshot, WorkbookSnapshot } from '../packages/protocol/src/types.js'
import type { WorkbookChangeUndoBundle } from '../packages/zero-sync/src/workbook-events.js'
import { deriveWorkbookActorHistoryState } from '../packages/zero-sync/src/workbook-history-state.js'

export interface AuditabilityControl {
  readonly id: string
  readonly category: 'preview-apply' | 'undo-revert' | 'authoritative-apply' | 'history-state'
  readonly required: boolean
  readonly passed: boolean
  readonly coveredControls: string[]
  readonly evidence: string
  readonly findings: string[]
}

export interface AuditabilityScorecard {
  readonly schemaVersion: 1
  readonly suite: 'auditability-posture'
  readonly generatedAt: string
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-auditability-scorecard.ts'
    readonly previewImplementation: 'packages/agent-api/src/workbook-agent-preview.ts'
    readonly applyImplementation: 'apps/bilig/src/zero/workbook-agent-apply.ts'
    readonly authoritativeApplyImplementation: 'apps/bilig/src/zero/service.ts'
    readonly historyImplementation: 'packages/zero-sync/src/workbook-history-state.ts'
  }
  readonly summary: {
    readonly allRequiredControlsPassed: boolean
    readonly previewApplyParityPassed: boolean
    readonly applyUndoRoundTripPassed: boolean
    readonly authoritativeApplyGuardPassed: boolean
    readonly historyRevertRedoPassed: boolean
    readonly coveredControls: string[]
    readonly uncoveredControls: string[]
    readonly externalGoogleSheetsEvidence: 'not-captured'
    readonly externalMicrosoftExcelEvidence: 'not-captured'
  }
  readonly controls: AuditabilityControl[]
}

interface AgentAuditabilityScenario {
  readonly previewMatchesAuthoritative: boolean
  readonly previewMatchesAppliedWorkbook: boolean
  readonly undoBundleCaptured: boolean
  readonly undoRestoresWorkbook: boolean
  readonly previewEffectSummary: WorkbookAgentPreviewSummary['effectSummary']
  readonly findings: string[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'auditability-scorecard.json')
const requiredControlIds = [
  'agent-preview-apply-parity',
  'agent-apply-undo-roundtrip',
  'authoritative-agent-apply-fails-closed',
  'workbook-history-revert-redo-state',
] as const
const coveredControlOrder = [
  'agent.previewDiffParity',
  'agent.authoritativePreviewMismatchFailsClosed',
  'agent.baseRevisionStaleApplyFailsClosed',
  'agent.applyCapturesUndoBundle',
  'agent.undoBundleRestoresSnapshot',
  'history.revertRedoStack',
  'history.revertLinkage',
] as const
const uncoveredControls = ['headedBrowser.previewApplyRevertFlow', 'externalSheetsExcelAuditabilityComparison'] as const

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error('Auditability scorecard is missing. Run: bun scripts/gen-auditability-scorecard.ts')
    }
    const scorecard = parseAuditabilityScorecard(JSON.parse(readFileSync(outputPath, 'utf8')) as unknown)
    validateAuditabilityScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = await buildAuditabilityScorecard()
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export async function buildAuditabilityScorecard(generatedAt = new Date().toISOString()): Promise<AuditabilityScorecard> {
  const agentScenario = await runAgentAuditabilityScenario()
  const controls = [
    buildAgentPreviewApplyParityControl(agentScenario),
    buildAgentApplyUndoRoundTripControl(agentScenario),
    buildAuthoritativeAgentApplyGuardControl(),
    buildWorkbookHistoryRevertRedoControl(),
  ]
  const coveredControlSet = new Set(controls.flatMap((control) => control.coveredControls))
  const coveredControls = coveredControlOrder.filter((control) => coveredControlSet.has(control))

  return {
    schemaVersion: 1,
    suite: 'auditability-posture',
    generatedAt,
    source: {
      artifactGenerator: 'scripts/gen-auditability-scorecard.ts',
      previewImplementation: 'packages/agent-api/src/workbook-agent-preview.ts',
      applyImplementation: 'apps/bilig/src/zero/workbook-agent-apply.ts',
      authoritativeApplyImplementation: 'apps/bilig/src/zero/service.ts',
      historyImplementation: 'packages/zero-sync/src/workbook-history-state.ts',
    },
    summary: {
      allRequiredControlsPassed: controls.filter((control) => control.required).every((control) => control.passed),
      previewApplyParityPassed: requiredControl(controls, 'agent-preview-apply-parity').passed,
      applyUndoRoundTripPassed: requiredControl(controls, 'agent-apply-undo-roundtrip').passed,
      authoritativeApplyGuardPassed: requiredControl(controls, 'authoritative-agent-apply-fails-closed').passed,
      historyRevertRedoPassed: requiredControl(controls, 'workbook-history-revert-redo-state').passed,
      coveredControls,
      uncoveredControls: [...uncoveredControls],
      externalGoogleSheetsEvidence: 'not-captured',
      externalMicrosoftExcelEvidence: 'not-captured',
    },
    controls,
  }
}

function buildAgentPreviewApplyParityControl(scenario: AgentAuditabilityScenario): AuditabilityControl {
  const passed = scenario.previewMatchesAuthoritative && scenario.previewMatchesAppliedWorkbook
  return auditabilityControl({
    id: 'agent-preview-apply-parity',
    category: 'preview-apply',
    passed,
    coveredControls: [
      'agent.previewDiffParity',
      'agent.authoritativePreviewMismatchFailsClosed',
      'agent.baseRevisionStaleApplyFailsClosed',
    ],
    evidence:
      `Executed workbook-agent preview twice from the same snapshot, compared the summaries for parity, and checked the sampled preview diff against the applied workbook. ` +
      `Effect summary: ${JSON.stringify(scenario.previewEffectSummary)}.`,
    findings: [
      ...(scenario.previewMatchesAuthoritative ? [] : ['local preview did not match authoritative preview for the same snapshot']),
      ...(scenario.previewMatchesAppliedWorkbook ? [] : ['preview diff did not match the post-apply workbook state']),
    ],
  })
}

function buildAgentApplyUndoRoundTripControl(scenario: AgentAuditabilityScenario): AuditabilityControl {
  return auditabilityControl({
    id: 'agent-apply-undo-roundtrip',
    category: 'undo-revert',
    passed: scenario.undoBundleCaptured && scenario.undoRestoresWorkbook,
    coveredControls: ['agent.applyCapturesUndoBundle', 'agent.undoBundleRestoresSnapshot'],
    evidence:
      'Applied a workbook-agent bundle through the server apply helper, required a captured undo bundle, and replayed it to restore the selected workbook snapshot.',
    findings: [
      ...(scenario.undoBundleCaptured ? [] : ['agent apply did not capture an undo bundle']),
      ...(scenario.undoRestoresWorkbook ? [] : ['captured undo bundle did not restore the workbook snapshot']),
      ...scenario.findings,
    ],
  })
}

function buildAuthoritativeAgentApplyGuardControl(): AuditabilityControl {
  const serviceSource = readFileSync(join(rootDir, 'apps', 'bilig', 'src', 'zero', 'service.ts'), 'utf8')
  const staleGuardIndex = serviceSource.indexOf('state.headRevision !== bundle.baseRevision')
  const staleErrorIndex = serviceSource.indexOf('WORKBOOK_AGENT_PREVIEW_STALE')
  const authoritativePreviewIndex = serviceSource.indexOf('const authoritativePreview = await buildWorkbookAgentPreview')
  const mismatchGuardIndex = serviceSource.indexOf('areWorkbookAgentPreviewSummariesEqual(preview, authoritativePreview)')
  const mismatchErrorIndex = serviceSource.indexOf('WORKBOOK_AGENT_PREVIEW_MISMATCH')
  const applyIndex = serviceSource.indexOf('applyWorkbookAgentCommandBundleWithUndoCapture(state.engine, bundle)')
  const persistIndex = serviceSource.indexOf('persistWorkbookMutation(client, documentId')
  const staleGuardBeforeApply = staleGuardIndex >= 0 && staleGuardIndex < applyIndex && staleErrorIndex > staleGuardIndex
  const previewBuiltBeforeMismatch = authoritativePreviewIndex >= 0 && authoritativePreviewIndex < mismatchGuardIndex
  const mismatchGuardBeforeApply =
    mismatchGuardIndex >= 0 && mismatchGuardIndex < applyIndex && mismatchErrorIndex > mismatchGuardIndex && mismatchErrorIndex < applyIndex
  const applyBeforePersist = applyIndex >= 0 && applyIndex < persistIndex

  return auditabilityControl({
    id: 'authoritative-agent-apply-fails-closed',
    category: 'authoritative-apply',
    passed: staleGuardBeforeApply && previewBuiltBeforeMismatch && mismatchGuardBeforeApply && applyBeforePersist,
    coveredControls: ['agent.authoritativePreviewMismatchFailsClosed', 'agent.baseRevisionStaleApplyFailsClosed'],
    evidence:
      'Statically checked the authoritative apply service rejects stale base revisions and preview mismatches before applying and persisting the command bundle.',
    findings: [
      ...(staleGuardBeforeApply ? [] : ['base-revision stale guard does not run before agent bundle apply']),
      ...(previewBuiltBeforeMismatch ? [] : ['authoritative preview is not built before preview comparison']),
      ...(mismatchGuardBeforeApply ? [] : ['preview mismatch guard does not fail closed before agent bundle apply']),
      ...(applyBeforePersist ? [] : ['agent bundle apply is not ordered before mutation persistence']),
    ],
  })
}

function buildWorkbookHistoryRevertRedoControl(): AuditabilityControl {
  const rows = [
    historyRow(10, 'owner@example.com', 'applyAgentCommandBundle', null, null),
    historyRow(11, 'collaborator@example.com', 'applyAgentCommandBundle', null, null),
    historyRow(12, 'owner@example.com', 'applyAgentCommandBundle', 13, null),
    historyRow(13, 'owner@example.com', 'revertChange', null, 12),
  ]
  const revertedState = deriveWorkbookActorHistoryState({
    actorUserId: 'owner@example.com',
    rows,
  })
  const redoneState = deriveWorkbookActorHistoryState({
    actorUserId: 'owner@example.com',
    rows: [...rows, historyRow(14, 'owner@example.com', 'redoChange', null, 13)],
  })
  const revertStackPassed =
    revertedState.canUndo &&
    revertedState.canRedo &&
    revertedState.undoRevision === 10 &&
    revertedState.redoRevision === 13 &&
    JSON.stringify(revertedState.undoStack) === JSON.stringify([10]) &&
    JSON.stringify(revertedState.redoStack) === JSON.stringify([13])
  const redoStackPassed =
    redoneState.canUndo &&
    !redoneState.canRedo &&
    redoneState.undoRevision === 14 &&
    redoneState.redoRevision === null &&
    JSON.stringify(redoneState.undoStack) === JSON.stringify([10, 14]) &&
    JSON.stringify(redoneState.redoStack) === JSON.stringify([])

  return auditabilityControl({
    id: 'workbook-history-revert-redo-state',
    category: 'history-state',
    passed: revertStackPassed && redoStackPassed,
    coveredControls: ['history.revertRedoStack', 'history.revertLinkage'],
    evidence:
      'Executed actor history derivation across own changes, collaborator changes, revert linkage, and redo linkage to prove the UI can expose undo/redo targets deterministically.',
    findings: [
      ...(revertStackPassed ? [] : [`unexpected revert state: ${JSON.stringify(revertedState)}`]),
      ...(redoStackPassed ? [] : [`unexpected redo state: ${JSON.stringify(redoneState)}`]),
    ],
  })
}

async function runAgentAuditabilityScenario(): Promise<AgentAuditabilityScenario> {
  const baseEngine = new SpreadsheetEngine({
    workbookName: 'Auditability Workbook',
    replicaId: 'auditability:base',
  })
  await baseEngine.ready()
  baseEngine.createSheet('Sheet1')
  baseEngine.setCellValue('Sheet1', 'A1', 21)
  baseEngine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A1)')
  const snapshot = baseEngine.exportSnapshot()
  const selectedBefore = selectSnapshotCells(snapshot, ['A1', 'B1'])
  const bundle = createAuditabilityBundle()
  const localPreview = await buildWorkbookAgentPreview({
    snapshot,
    replicaId: 'auditability:local-preview',
    bundle,
  })
  const authoritativePreview = await buildWorkbookAgentPreview({
    snapshot,
    replicaId: 'auditability:authoritative-preview',
    bundle,
  })
  const applyEngine = new SpreadsheetEngine({
    workbookName: snapshot.workbook.name,
    replicaId: 'auditability:apply',
  })
  await applyEngine.ready()
  applyEngine.importSnapshot(snapshot)
  const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(applyEngine, bundle)
  const previewMatchesAppliedWorkbook = localPreview.cellDiffs.every((diff) => {
    const appliedCell = applyEngine.getCell(diff.sheetName, diff.address)
    return (
      normalizePreviewFormula(appliedCell.formula) === diff.afterFormula && normalizePreviewInput(appliedCell.input) === diff.afterInput
    )
  })
  const undoRestoresWorkbook = applyUndoBundle(applyEngine, undoBundle)
    ? JSON.stringify(selectSnapshotCells(applyEngine.exportSnapshot(), ['A1', 'B1'])) === JSON.stringify(selectedBefore)
    : false

  return {
    previewMatchesAuthoritative: areWorkbookAgentPreviewSummariesEqual(localPreview, authoritativePreview),
    previewMatchesAppliedWorkbook,
    undoBundleCaptured: undoBundle !== null,
    undoRestoresWorkbook,
    previewEffectSummary: localPreview.effectSummary,
    findings: [
      ...(localPreview.effectSummary.formulaChangeCount === 1 ? [] : ['preview did not capture exactly one formula change']),
      ...(localPreview.effectSummary.styleChangeCount === 1 ? [] : ['preview did not capture exactly one style change']),
      ...(localPreview.effectSummary.numberFormatChangeCount === 1 ? [] : ['preview did not capture exactly one number-format change']),
    ],
  }
}

function applyUndoBundle(engine: SpreadsheetEngine, undoBundle: WorkbookChangeUndoBundle | null): boolean {
  if (!undoBundle) {
    return false
  }
  switch (undoBundle.kind) {
    case 'engineOps':
      engine.applyOps(undoBundle.ops, { trusted: true })
      return true
    case 'snapshot':
      engine.importSnapshot(undoBundle.snapshot)
      return true
    default: {
      const exhaustive: never = undoBundle
      return exhaustive
    }
  }
}

function normalizePreviewFormula(formula: string | undefined): string | null {
  return formula === undefined ? null : `=${formula}`
}

function normalizePreviewInput(input: CellSnapshot['input']): CellSnapshot['input'] | null {
  return input === undefined ? null : input
}

function createAuditabilityBundle(): WorkbookAgentCommandBundle {
  return {
    id: 'auditability-bundle-1',
    documentId: 'doc-auditability',
    threadId: 'thread-auditability',
    turnId: 'turn-auditability',
    goalText: 'Audit preview apply undo',
    summary: 'Audit preview apply undo',
    scope: 'sheet',
    riskClass: 'medium',
    baseRevision: 1,
    createdAtUnixMs: 1,
    context: null,
    commands: [
      {
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'B1',
        values: [[{ formula: '=A1*3' }]],
      },
      {
        kind: 'formatRange',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
        patch: {
          font: {
            bold: true,
          },
        },
        numberFormat: 'currency',
      },
    ],
    affectedRanges: [
      {
        sheetName: 'Sheet1',
        startAddress: 'B1',
        endAddress: 'B1',
        role: 'target',
      },
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
        role: 'target',
      },
    ],
    estimatedAffectedCells: 2,
  }
}

function historyRow(
  revision: number,
  actorUserId: string,
  eventKind: string,
  revertedByRevision: number | null,
  revertsRevision: number | null,
) {
  return {
    revision,
    actorUserId,
    eventKind,
    undoBundleJson: { kind: 'engineOps', ops: [] },
    revertedByRevision,
    revertsRevision,
  } as const
}

function selectSnapshotCells(snapshot: WorkbookSnapshot, addresses: readonly string[]): readonly CellSnapshot[] {
  const addressSet = new Set(addresses)
  return snapshot.sheets
    .flatMap((sheet) => sheet.cells.map((cell) => ({ ...cell, sheetName: sheet.name })))
    .filter((cell) => addressSet.has(cell.address))
    .toSorted((left, right) => `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`))
}

function auditabilityControl(input: {
  readonly id: AuditabilityControl['id']
  readonly category: AuditabilityControl['category']
  readonly passed: boolean
  readonly coveredControls: readonly string[]
  readonly evidence: string
  readonly findings: readonly string[]
}): AuditabilityControl {
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

function requiredControl(controls: readonly AuditabilityControl[], id: string): AuditabilityControl {
  const entry = controls.find((control) => control.id === id)
  if (!entry) {
    throw new Error(`Auditability scorecard is missing required control: ${id}`)
  }
  return entry
}

export function parseAuditabilityScorecard(value: unknown): AuditabilityScorecard {
  const record = toRecord(value, 'auditability scorecard')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'auditability-posture') {
    throw new Error('Unexpected auditability scorecard header')
  }
  const source = recordField(record, 'source', 'auditability source')
  const summary = recordField(record, 'summary', 'auditability summary')
  return {
    schemaVersion: 1,
    suite: 'auditability-posture',
    generatedAt: stringField(record, 'generatedAt', 'auditability generatedAt'),
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-auditability-scorecard.ts'),
      previewImplementation: literalField(source, 'previewImplementation', 'packages/agent-api/src/workbook-agent-preview.ts'),
      applyImplementation: literalField(source, 'applyImplementation', 'apps/bilig/src/zero/workbook-agent-apply.ts'),
      authoritativeApplyImplementation: literalField(source, 'authoritativeApplyImplementation', 'apps/bilig/src/zero/service.ts'),
      historyImplementation: literalField(source, 'historyImplementation', 'packages/zero-sync/src/workbook-history-state.ts'),
    },
    summary: {
      allRequiredControlsPassed: booleanField(summary, 'allRequiredControlsPassed', 'auditability allRequiredControlsPassed'),
      previewApplyParityPassed: booleanField(summary, 'previewApplyParityPassed', 'auditability previewApplyParityPassed'),
      applyUndoRoundTripPassed: booleanField(summary, 'applyUndoRoundTripPassed', 'auditability applyUndoRoundTripPassed'),
      authoritativeApplyGuardPassed: booleanField(summary, 'authoritativeApplyGuardPassed', 'auditability authoritativeApplyGuardPassed'),
      historyRevertRedoPassed: booleanField(summary, 'historyRevertRedoPassed', 'auditability historyRevertRedoPassed'),
      coveredControls: stringArrayField(summary, 'coveredControls', 'auditability coveredControls'),
      uncoveredControls: stringArrayField(summary, 'uncoveredControls', 'auditability uncoveredControls'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'not-captured'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'not-captured'),
    },
    controls: arrayField(record, 'controls', 'auditability controls').map(parseAuditabilityControl),
  }
}

function parseAuditabilityControl(value: unknown): AuditabilityControl {
  const record = toRecord(value, 'auditability control')
  return {
    id: stringField(record, 'id', 'auditability control id'),
    category: parseAuditabilityCategory(stringField(record, 'category', 'auditability control category')),
    required: booleanField(record, 'required', 'auditability control required'),
    passed: booleanField(record, 'passed', 'auditability control passed'),
    coveredControls: stringArrayField(record, 'coveredControls', 'auditability control coveredControls'),
    evidence: stringField(record, 'evidence', 'auditability control evidence'),
    findings: stringArrayField(record, 'findings', 'auditability control findings'),
  }
}

export function validateAuditabilityScorecard(scorecard: AuditabilityScorecard): void {
  for (const id of requiredControlIds) {
    const control = requiredControl(scorecard.controls, id)
    if (!control.required) {
      throw new Error(`Auditability scorecard required control is not marked required: ${id}`)
    }
    if (!control.passed) {
      throw new Error(`Auditability scorecard contains a failed required control: ${id}`)
    }
  }
  if (!scorecard.summary.allRequiredControlsPassed) {
    throw new Error('Auditability scorecard summary reports failed required controls')
  }
  for (const control of coveredControlOrder) {
    if (!scorecard.summary.coveredControls.includes(control)) {
      throw new Error(`Auditability scorecard is missing covered control: ${control}`)
    }
  }
  for (const control of uncoveredControls) {
    if (!scorecard.summary.uncoveredControls.includes(control)) {
      throw new Error(`Auditability scorecard is missing uncovered control disclosure: ${control}`)
    }
  }
}

function parseAuditabilityCategory(value: string): AuditabilityControl['category'] {
  if (value === 'preview-apply' || value === 'undo-revert' || value === 'authoritative-apply' || value === 'history-state') {
    return value
  }
  throw new Error(`Unexpected auditability category: ${value}`)
}

function logResult(mode: 'check' | 'write', scorecard: AuditabilityScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        allRequiredControlsPassed: scorecard.summary.allRequiredControlsPassed,
        coveredControls: scorecard.summary.coveredControls.length,
        uncoveredControls: scorecard.summary.uncoveredControls.length,
      },
      null,
      2,
    ),
  )
}

function recordField(value: Record<string, unknown>, field: string, name: string): Record<string, unknown> {
  return toRecord(value[field], name)
}

function arrayField(value: Record<string, unknown>, field: string, name: string): unknown[] {
  const fieldValue = value[field]
  if (!Array.isArray(fieldValue)) {
    throw new Error(`Expected ${name} to be an array`)
  }
  return fieldValue
}

function stringArrayField(value: Record<string, unknown>, field: string, name: string): string[] {
  const fieldValue = arrayField(value, field, name)
  if (!fieldValue.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected ${name} to contain only strings`)
  }
  return fieldValue
}

function stringField(value: Record<string, unknown>, field: string, name: string): string {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${name} to be a string`)
  }
  return fieldValue
}

function booleanField(value: Record<string, unknown>, field: string, name: string): boolean {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${name} to be a boolean`)
  }
  return fieldValue
}

function literalField<const T extends string>(value: Record<string, unknown>, field: string, expected: T): T {
  if (value[field] !== expected) {
    throw new Error(`Expected ${field} to be ${expected}`)
  }
  return expected
}

function toRecord(value: unknown, name: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Expected ${name} to be an object`)
  }
  const record: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    record[key] = Reflect.get(value, key)
  }
  return record
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'auditability-scorecard-'))
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
    throw new Error(`Unable to format generated auditability scorecard: ${new TextDecoder().decode(formatResult.stderr).trim()}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
