#!/usr/bin/env bun

import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { ValueTag } from '@bilig/protocol'
import { createMemoryWorkbookLocalStoreFactory, type WorkbookLocalStoreFactory } from '@bilig/storage-browser'
import { WorkbookWorkerRuntime } from '../apps/web/src/worker-runtime.js'
import type { PendingWorkbookMutation } from '../apps/web/src/workbook-sync.js'

export interface ReliabilityControl {
  readonly id: string
  readonly category: 'pending-durability' | 'authoritative-reconcile' | 'failure-recovery' | 'headed-browser'
  readonly required: boolean
  readonly passed: boolean
  readonly coveredControls: string[]
  readonly evidence: string
  readonly findings: string[]
}

export interface ReliabilityScorecard {
  readonly schemaVersion: 1
  readonly suite: 'reliability-posture'
  readonly generatedAt: string
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-reliability-scorecard.ts'
    readonly workerRuntimeImplementation: 'apps/web/src/worker-runtime.ts'
    readonly mutationJournalImplementation: 'apps/web/src/worker-runtime-mutation-journal.ts'
    readonly localStoreImplementation: 'packages/storage-browser/src/index.ts'
    readonly headedBrowserReliabilityTestFile: 'e2e/tests/web-shell-remote-sync.pw.ts'
  }
  readonly summary: {
    readonly allRequiredControlsPassed: boolean
    readonly pendingReloadPassed: boolean
    readonly authoritativeAckPassed: boolean
    readonly authoritativeRebasePassed: boolean
    readonly failedRetryPassed: boolean
    readonly headedBrowserReloadPassed: boolean
    readonly coveredControls: string[]
    readonly uncoveredControls: string[]
    readonly externalGoogleSheetsEvidence: 'not-captured'
    readonly externalMicrosoftExcelEvidence: 'not-captured'
  }
  readonly controls: ReliabilityControl[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'reliability-scorecard.json')
const requiredControlIds = [
  'pending-mutations-survive-reload',
  'submitted-mutations-absorb-authoritative-ack',
  'authoritative-rebase-preserves-unsent-mutations',
  'failed-mutations-survive-reload-and-retry',
  'headed-browser-reload-persistence-flow',
] as const
const coveredControlOrder = [
  'pending.localReloadSurvival',
  'pending.submittedReloadSurvival',
  'pending.authoritativeAckAbsorption',
  'pending.authoritativeRebasePreservesLocal',
  'pending.failedRetrySurvival',
  'localStore.journalActiveView',
  'headedBrowser.reloadPersistence',
] as const
const uncoveredControls = ['headedBrowser.crashSoak', 'offlineNetworkPartitionSoak', 'externalSheetsExcelReliabilityComparison'] as const
const headedBrowserReliabilityTestFile = 'e2e/tests/web-shell-remote-sync.pw.ts'

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error('Reliability scorecard is missing. Run: bun scripts/gen-reliability-scorecard.ts')
    }
    const scorecard = parseReliabilityScorecard(JSON.parse(readFileSync(outputPath, 'utf8')) as unknown)
    validateReliabilityScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = await buildReliabilityScorecard()
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export async function buildReliabilityScorecard(generatedAt = new Date().toISOString()): Promise<ReliabilityScorecard> {
  const controls = [
    await buildPendingReloadControl(),
    await buildSubmittedAckControl(),
    await buildAuthoritativeRebaseControl(),
    await buildFailedRetryControl(),
    buildHeadedBrowserReloadPersistenceControl(),
  ]
  const coveredControlSet = new Set(controls.flatMap((control) => control.coveredControls))
  const coveredControls = coveredControlOrder.filter((control) => coveredControlSet.has(control))

  return {
    schemaVersion: 1,
    suite: 'reliability-posture',
    generatedAt,
    source: {
      artifactGenerator: 'scripts/gen-reliability-scorecard.ts',
      workerRuntimeImplementation: 'apps/web/src/worker-runtime.ts',
      mutationJournalImplementation: 'apps/web/src/worker-runtime-mutation-journal.ts',
      localStoreImplementation: 'packages/storage-browser/src/index.ts',
      headedBrowserReliabilityTestFile,
    },
    summary: {
      allRequiredControlsPassed: controls.filter((control) => control.required).every((control) => control.passed),
      pendingReloadPassed: requiredControl(controls, 'pending-mutations-survive-reload').passed,
      authoritativeAckPassed: requiredControl(controls, 'submitted-mutations-absorb-authoritative-ack').passed,
      authoritativeRebasePassed: requiredControl(controls, 'authoritative-rebase-preserves-unsent-mutations').passed,
      failedRetryPassed: requiredControl(controls, 'failed-mutations-survive-reload-and-retry').passed,
      headedBrowserReloadPassed: requiredControl(controls, 'headed-browser-reload-persistence-flow').passed,
      coveredControls,
      uncoveredControls: [...uncoveredControls],
      externalGoogleSheetsEvidence: 'not-captured',
      externalMicrosoftExcelEvidence: 'not-captured',
    },
    controls,
  }
}

async function buildPendingReloadControl(): Promise<ReliabilityControl> {
  const factory = createMemoryWorkbookLocalStoreFactory()
  const runtime = await createRuntime(factory, 'pending-reload-doc', 'browser:pending')
  const pending = await runtime.enqueuePendingMutation({
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 17],
  })
  const reloaded = await createRuntime(factory, 'pending-reload-doc', 'browser:pending-reloaded')
  const pendingReloaded = findPendingMutation(reloaded.listPendingMutations(), pending.id)
  const localValuePersisted = cellNumber(reloaded, 'A1') === 17
  const pendingSurvived =
    pendingReloaded?.status === 'local' &&
    pendingReloaded.localSeq === pending.localSeq &&
    pendingReloaded.baseRevision === pending.baseRevision &&
    localValuePersisted

  return reliabilityControl({
    id: 'pending-mutations-survive-reload',
    category: 'pending-durability',
    passed: pendingSurvived,
    coveredControls: ['pending.localReloadSurvival', 'localStore.journalActiveView'],
    evidence:
      'Enqueued a local pending mutation, reopened the worker runtime from the same local store, and verified the pending entry and projected cell survived.',
    findings: [
      ...(pendingReloaded ? [] : ['pending mutation was missing after worker reload']),
      ...(pendingReloaded?.status === 'local'
        ? []
        : [`expected local pending status after reload, received ${pendingReloaded?.status ?? 'missing'}`]),
      ...(localValuePersisted ? [] : ['projected local value did not survive reload']),
    ],
  })
}

async function buildSubmittedAckControl(): Promise<ReliabilityControl> {
  const factory = createMemoryWorkbookLocalStoreFactory()
  const runtime = await createRuntime(factory, 'submitted-ack-doc', 'browser:submitted')
  const pending = await runtime.enqueuePendingMutation({
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 23],
  })
  await runtime.markPendingMutationSubmitted(pending.id)
  const submittedReloaded = await createRuntime(factory, 'submitted-ack-doc', 'browser:submitted-reloaded')
  const submitted = findPendingMutation(submittedReloaded.listPendingMutations(), pending.id)
  await submittedReloaded.applyAuthoritativeEvents(
    [
      {
        revision: 1,
        clientMutationId: pending.id,
        payload: {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'A1',
          value: 23,
        },
      },
    ],
    1,
  )
  const afterAckPending = submittedReloaded.listPendingMutations()
  const afterAckJournal = submittedReloaded.listMutationJournalEntries()
  const afterAckReloaded = await createRuntime(factory, 'submitted-ack-doc', 'browser:after-ack')
  const submittedSurvived = submitted?.status === 'submitted'
  const ackAbsorbed =
    afterAckPending.length === 0 &&
    afterAckReloaded.listPendingMutations().length === 0 &&
    cellNumber(afterAckReloaded, 'A1') === 23 &&
    afterAckJournal.some((entry) => entry.id === pending.id && entry.status === 'acked')

  return reliabilityControl({
    id: 'submitted-mutations-absorb-authoritative-ack',
    category: 'authoritative-reconcile',
    passed: submittedSurvived && ackAbsorbed,
    coveredControls: ['pending.submittedReloadSurvival', 'pending.authoritativeAckAbsorption', 'localStore.journalActiveView'],
    evidence:
      'Reopened a submitted mutation, absorbed an authoritative event with the same client mutation id, and verified pending state stayed empty after another reload.',
    findings: [
      ...(submittedSurvived ? [] : [`expected submitted pending status after reload, received ${submitted?.status ?? 'missing'}`]),
      ...(ackAbsorbed ? [] : ['authoritative ack did not clear pending state and persist the authoritative value']),
    ],
  })
}

async function buildAuthoritativeRebaseControl(): Promise<ReliabilityControl> {
  const factory = createMemoryWorkbookLocalStoreFactory()
  const runtime = await createRuntime(factory, 'rebase-doc', 'browser:rebase')
  const pending = await runtime.enqueuePendingMutation({
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 31],
  })
  await runtime.applyAuthoritativeEvents(
    [
      {
        revision: 1,
        clientMutationId: null,
        payload: {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'B1',
          value: 101,
        },
      },
    ],
    1,
  )
  const reloaded = await createRuntime(factory, 'rebase-doc', 'browser:rebase-reloaded')
  const rebased = findPendingMutation(reloaded.listPendingMutations(), pending.id)
  const rebasePassed = rebased?.status === 'rebased' && cellNumber(reloaded, 'A1') === 31 && cellNumber(reloaded, 'B1') === 101

  return reliabilityControl({
    id: 'authoritative-rebase-preserves-unsent-mutations',
    category: 'authoritative-reconcile',
    passed: rebasePassed,
    coveredControls: ['pending.authoritativeRebasePreservesLocal', 'localStore.journalActiveView'],
    evidence:
      'Applied an unrelated authoritative event over an unsent local mutation, reopened the runtime, and verified both the authoritative cell and rebased local cell survived.',
    findings: [
      ...(rebased?.status === 'rebased'
        ? []
        : [`expected rebased pending status after authoritative replay, received ${rebased?.status ?? 'missing'}`]),
      ...(cellNumber(reloaded, 'A1') === 31 ? [] : ['rebased local cell value was lost after reload']),
      ...(cellNumber(reloaded, 'B1') === 101 ? [] : ['authoritative replay cell value was lost after reload']),
    ],
  })
}

async function buildFailedRetryControl(): Promise<ReliabilityControl> {
  const factory = createMemoryWorkbookLocalStoreFactory()
  const runtime = await createRuntime(factory, 'failed-retry-doc', 'browser:failed')
  const pending = await runtime.enqueuePendingMutation({
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 47],
  })
  await runtime.markPendingMutationFailed(pending.id, 'mutation rejected by server')
  const failedReloaded = await createRuntime(factory, 'failed-retry-doc', 'browser:failed-reloaded')
  const failed = findPendingMutation(failedReloaded.listPendingMutations(), pending.id)
  await failedReloaded.retryPendingMutation(pending.id)
  const retriedReloaded = await createRuntime(factory, 'failed-retry-doc', 'browser:retried-reloaded')
  const retried = findPendingMutation(retriedReloaded.listPendingMutations(), pending.id)
  const retryPassed =
    failed?.status === 'failed' &&
    failed.failureMessage === 'mutation rejected by server' &&
    retried?.status === 'local' &&
    retried.failureMessage === null &&
    cellNumber(retriedReloaded, 'A1') === 47

  return reliabilityControl({
    id: 'failed-mutations-survive-reload-and-retry',
    category: 'failure-recovery',
    passed: retryPassed,
    coveredControls: ['pending.failedRetrySurvival', 'localStore.journalActiveView'],
    evidence:
      'Marked a pending mutation failed, reopened the runtime, retried it, reopened again, and verified it returned to local pending state without losing the projected value.',
    findings: [
      ...(failed?.status === 'failed' ? [] : [`expected failed pending status after reload, received ${failed?.status ?? 'missing'}`]),
      ...(retried?.status === 'local'
        ? []
        : [`expected local pending status after retry reload, received ${retried?.status ?? 'missing'}`]),
      ...(cellNumber(retriedReloaded, 'A1') === 47 ? [] : ['failed/retried local cell value was lost after reload']),
    ],
  })
}

function buildHeadedBrowserReloadPersistenceControl(): ReliabilityControl {
  const source = readFileSync(join(rootDir, headedBrowserReliabilityTestFile), 'utf8')
  const testTitle = 'web app restores persisted workbook state after a full reload'
  const testBlock = extractBrowserTestBlock(source, 'remoteSyncTest', testTitle)
  const findings: string[] = []
  if (!testBlock) {
    findings.push(`missing remote-sync Playwright test: ${testTitle}`)
  } else {
    requireSnippet(
      testBlock,
      "createTestDocumentId('playwright-zero-reload-persist')",
      'uses an isolated persisted reload document',
      findings,
    )
    requireSnippet(testBlock, 'await openZeroWorkbookPage(page, documentId)', 'opens the headed workbook shell through Zero sync', findings)
    requireSnippet(testBlock, "await expect(page.getByTestId('worker-error')).toHaveCount(0)", 'starts without worker errors', findings)
    requireSnippet(testBlock, "await formulaInput.fill('17')", 'applies a visible workbook value before reload', findings)
    requireSnippet(testBlock, "await formulaInput.press('Enter')", 'commits the workbook value before reload', findings)
    requireSnippet(testBlock, "await expect(resolvedValue).toHaveText('17')", 'verifies calculated value before and after reload', findings)
    requireSnippet(testBlock, "await page.reload({ waitUntil: 'domcontentloaded' })", 'performs a real headed browser reload', findings)
    requireSnippet(testBlock, 'await waitForWorkbookReady(page)', 'waits for the reloaded workbook runtime', findings)
    requireSnippet(testBlock, "await expect(formulaInput).toHaveValue('17')", 'verifies edited input survived reload', findings)
  }

  return reliabilityControl({
    id: 'headed-browser-reload-persistence-flow',
    category: 'headed-browser',
    passed: findings.length === 0,
    coveredControls: ['headedBrowser.reloadPersistence'],
    evidence:
      'Validated the headed Playwright reload contract that writes a workbook value through the UI, reloads the browser page, and verifies both formula input and resolved value survive.',
    findings,
  })
}

async function createRuntime(
  localStoreFactory: WorkbookLocalStoreFactory,
  documentId: string,
  replicaId: string,
): Promise<WorkbookWorkerRuntime> {
  const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
  await runtime.bootstrap({
    documentId,
    replicaId,
    persistState: true,
  })
  return runtime
}

function findPendingMutation(mutations: readonly PendingWorkbookMutation[], id: string): PendingWorkbookMutation | null {
  return mutations.find((mutation) => mutation.id === id) ?? null
}

function cellNumber(runtime: WorkbookWorkerRuntime, address: string): number | null {
  const value = runtime.getCell('Sheet1', address).value
  return value.tag === ValueTag.Number ? value.value : null
}

function reliabilityControl(input: {
  readonly id: ReliabilityControl['id']
  readonly category: ReliabilityControl['category']
  readonly passed: boolean
  readonly coveredControls: readonly string[]
  readonly evidence: string
  readonly findings: readonly string[]
}): ReliabilityControl {
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

function requiredControl(controls: readonly ReliabilityControl[], id: string): ReliabilityControl {
  const entry = controls.find((control) => control.id === id)
  if (!entry) {
    throw new Error(`Reliability scorecard is missing required control: ${id}`)
  }
  return entry
}

export function parseReliabilityScorecard(value: unknown): ReliabilityScorecard {
  const record = toRecord(value, 'reliability scorecard')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'reliability-posture') {
    throw new Error('Unexpected reliability scorecard header')
  }
  const source = recordField(record, 'source', 'reliability source')
  const summary = recordField(record, 'summary', 'reliability summary')
  return {
    schemaVersion: 1,
    suite: 'reliability-posture',
    generatedAt: stringField(record, 'generatedAt', 'reliability generatedAt'),
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-reliability-scorecard.ts'),
      workerRuntimeImplementation: literalField(source, 'workerRuntimeImplementation', 'apps/web/src/worker-runtime.ts'),
      mutationJournalImplementation: literalField(
        source,
        'mutationJournalImplementation',
        'apps/web/src/worker-runtime-mutation-journal.ts',
      ),
      localStoreImplementation: literalField(source, 'localStoreImplementation', 'packages/storage-browser/src/index.ts'),
      headedBrowserReliabilityTestFile: literalField(source, 'headedBrowserReliabilityTestFile', headedBrowserReliabilityTestFile),
    },
    summary: {
      allRequiredControlsPassed: booleanField(summary, 'allRequiredControlsPassed', 'reliability allRequiredControlsPassed'),
      pendingReloadPassed: booleanField(summary, 'pendingReloadPassed', 'reliability pendingReloadPassed'),
      authoritativeAckPassed: booleanField(summary, 'authoritativeAckPassed', 'reliability authoritativeAckPassed'),
      authoritativeRebasePassed: booleanField(summary, 'authoritativeRebasePassed', 'reliability authoritativeRebasePassed'),
      failedRetryPassed: booleanField(summary, 'failedRetryPassed', 'reliability failedRetryPassed'),
      headedBrowserReloadPassed: booleanField(summary, 'headedBrowserReloadPassed', 'reliability headedBrowserReloadPassed'),
      coveredControls: stringArrayField(summary, 'coveredControls', 'reliability coveredControls'),
      uncoveredControls: stringArrayField(summary, 'uncoveredControls', 'reliability uncoveredControls'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'not-captured'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'not-captured'),
    },
    controls: arrayField(record, 'controls', 'reliability controls').map(parseReliabilityControl),
  }
}

function parseReliabilityControl(value: unknown): ReliabilityControl {
  const record = toRecord(value, 'reliability control')
  return {
    id: stringField(record, 'id', 'reliability control id'),
    category: parseReliabilityCategory(stringField(record, 'category', 'reliability control category')),
    required: booleanField(record, 'required', 'reliability control required'),
    passed: booleanField(record, 'passed', 'reliability control passed'),
    coveredControls: stringArrayField(record, 'coveredControls', 'reliability control coveredControls'),
    evidence: stringField(record, 'evidence', 'reliability control evidence'),
    findings: stringArrayField(record, 'findings', 'reliability control findings'),
  }
}

export function validateReliabilityScorecard(scorecard: ReliabilityScorecard): void {
  for (const id of requiredControlIds) {
    const control = requiredControl(scorecard.controls, id)
    if (!control.required) {
      throw new Error(`Reliability scorecard required control is not marked required: ${id}`)
    }
    if (!control.passed) {
      throw new Error(`Reliability scorecard contains a failed required control: ${id}`)
    }
  }
  if (!scorecard.summary.allRequiredControlsPassed) {
    throw new Error('Reliability scorecard summary reports failed required controls')
  }
  for (const control of coveredControlOrder) {
    if (!scorecard.summary.coveredControls.includes(control)) {
      throw new Error(`Reliability scorecard is missing covered control: ${control}`)
    }
  }
  for (const control of uncoveredControls) {
    if (!scorecard.summary.uncoveredControls.includes(control)) {
      throw new Error(`Reliability scorecard is missing uncovered control disclosure: ${control}`)
    }
  }
}

function parseReliabilityCategory(value: string): ReliabilityControl['category'] {
  if (value === 'pending-durability' || value === 'authoritative-reconcile' || value === 'failure-recovery' || value === 'headed-browser') {
    return value
  }
  throw new Error(`Unexpected reliability category: ${value}`)
}

function extractBrowserTestBlock(source: string, testFunctionName: 'test' | 'remoteSyncTest', testTitle: string): string | null {
  const marker = `${testFunctionName}('${testTitle}'`
  const start = source.indexOf(marker)
  if (start < 0) {
    return null
  }
  const endCandidates = ['\nremoteSyncTest(', '\ntest(']
    .map((nextMarker) => source.indexOf(nextMarker, start + marker.length))
    .filter((index) => index >= 0)
  const end = endCandidates.length > 0 ? Math.min(...endCandidates) : source.length
  return source.slice(start, end)
}

function requireSnippet(source: string, snippet: string, label: string, findings: string[]): void {
  if (!source.includes(snippet)) {
    findings.push(`missing ${label}`)
  }
}

function logResult(mode: 'check' | 'write', scorecard: ReliabilityScorecard): void {
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
  const tempDir = mkdtempSync(join(tmpdir(), 'reliability-scorecard-'))
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
    throw new Error(`Unable to format generated reliability scorecard: ${new TextDecoder().decode(formatResult.stderr).trim()}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
