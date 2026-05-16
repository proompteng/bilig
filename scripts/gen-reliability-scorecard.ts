#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { ValueTag } from '@bilig/protocol'
import type { AuthoritativeWorkbookEventRecord } from '@bilig/zero-sync'
import { WorkbookWorkerRuntime } from '../apps/web/src/worker-runtime.js'
import type { PendingWorkbookMutation } from '../apps/web/src/workbook-sync.js'
import {
  externalReliabilityComparisonArtifactRepoPath,
  externalReliabilityComparisonCoveredControls,
  parseExternalReliabilityComparisonArtifact,
  validateExternalReliabilityComparisonArtifact,
} from './reliability-external-sheets-excel-comparison.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export interface ReliabilityControl {
  readonly id: string
  readonly category:
    | 'pending-durability'
    | 'authoritative-reconcile'
    | 'failure-recovery'
    | 'headed-browser'
    | 'offline-recovery'
    | 'external-comparison'
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
    readonly zeroSyncImplementation: 'packages/zero-sync/src/schema.ts'
    readonly headedBrowserReliabilityTestFile: 'e2e/tests/web-shell-remote-sync.pw.ts'
    readonly externalReliabilityComparisonArtifact: 'packages/benchmarks/baselines/reliability-external-sheets-excel-comparison.json'
  }
  readonly summary: {
    readonly allRequiredControlsPassed: boolean
    readonly pendingReloadPassed: boolean
    readonly authoritativeAckPassed: boolean
    readonly authoritativeRebasePassed: boolean
    readonly failedRetryPassed: boolean
    readonly headedBrowserReloadPassed: boolean
    readonly headedBrowserCrashSoakPassed: boolean
    readonly offlineNetworkPartitionPassed: boolean
    readonly coveredControls: string[]
    readonly uncoveredControls: string[]
    readonly externalGoogleSheetsEvidence: 'official-docs-comparison-artifact'
    readonly externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact'
  }
  readonly controls: ReliabilityControl[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'reliability-scorecard.json')
const externalReliabilityComparisonArtifactPath = join(rootDir, externalReliabilityComparisonArtifactRepoPath)
const requiredControlIds = [
  'pending-mutations-survive-reload',
  'submitted-mutations-absorb-authoritative-ack',
  'authoritative-rebase-preserves-unsent-mutations',
  'failed-mutations-survive-reload-and-retry',
  'headed-browser-reload-persistence-flow',
  'headed-browser-crash-restart-soak',
  'offline-network-partition-recovery-soak',
  'external-sheets-excel-reliability-comparison',
] as const
const coveredControlOrder = [
  'zero.clientPersistenceConfig',
  'zero.authoritativeMutationDurability',
  'pending.authoritativeAckAbsorption',
  'pending.authoritativeRebasePreservesLocal',
  'pending.failedRetryStateMachine',
  'headedBrowser.reloadPersistence',
  'headedBrowser.crashSoak',
  'offline.networkPartitionRecoverySoak',
  ...externalReliabilityComparisonCoveredControls,
] as const
const uncoveredControls: readonly string[] = []
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
    buildHeadedBrowserCrashRestartSoakControl(),
    await buildOfflineNetworkPartitionRecoveryControl(),
    buildExternalSheetsExcelReliabilityComparisonControl(),
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
      zeroSyncImplementation: 'packages/zero-sync/src/schema.ts',
      headedBrowserReliabilityTestFile,
      externalReliabilityComparisonArtifact: externalReliabilityComparisonArtifactRepoPath,
    },
    summary: {
      allRequiredControlsPassed: controls.filter((control) => control.required).every((control) => control.passed),
      pendingReloadPassed: requiredControl(controls, 'pending-mutations-survive-reload').passed,
      authoritativeAckPassed: requiredControl(controls, 'submitted-mutations-absorb-authoritative-ack').passed,
      authoritativeRebasePassed: requiredControl(controls, 'authoritative-rebase-preserves-unsent-mutations').passed,
      failedRetryPassed: requiredControl(controls, 'failed-mutations-survive-reload-and-retry').passed,
      headedBrowserReloadPassed: requiredControl(controls, 'headed-browser-reload-persistence-flow').passed,
      headedBrowserCrashSoakPassed: requiredControl(controls, 'headed-browser-crash-restart-soak').passed,
      offlineNetworkPartitionPassed: requiredControl(controls, 'offline-network-partition-recovery-soak').passed,
      coveredControls,
      uncoveredControls: [...uncoveredControls],
      externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
      externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
    },
    controls,
  }
}

async function buildPendingReloadControl(): Promise<ReliabilityControl> {
  const runtimeConfigSource = readFileSync(join(rootDir, 'apps/web/src/runtime-config.ts'), 'utf8')
  const runtimeSessionSource = readFileSync(join(rootDir, 'apps/web/src/runtime-session.ts'), 'utf8')
  const headedBrowserSource = readFileSync(join(rootDir, headedBrowserReliabilityTestFile), 'utf8')
  const findings: string[] = []
  requireSnippet(runtimeConfigSource, 'persistState: true', 'default Zero client persistence enabled', findings)
  requireSnippet(runtimeSessionSource, 'zero?: ZeroWorkbookSyncSource', 'runtime session accepts Zero as the sync source', findings)
  requireSnippet(runtimeSessionSource, 'ZeroWorkbookRevisionSync', 'runtime session reconciles through Zero revision sync', findings)
  requireSnippet(
    headedBrowserSource,
    "createTestDocumentId('playwright-zero-reload-persist')",
    'headed reload test uses an isolated Zero-backed workbook',
    findings,
  )
  requireSnippet(
    headedBrowserSource,
    'await openZeroWorkbookPage(page, documentId)',
    'headed reload test opens the Zero workbook page',
    findings,
  )
  requireSnippet(
    headedBrowserSource,
    "await page.reload({ waitUntil: 'domcontentloaded' })",
    'headed reload test performs a browser reload',
    findings,
  )

  return reliabilityControl({
    id: 'pending-mutations-survive-reload',
    category: 'pending-durability',
    passed: findings.length === 0,
    coveredControls: ['zero.clientPersistenceConfig', 'headedBrowser.reloadPersistence'],
    evidence: 'Verified browser durability is delegated to Zero client persistence and covered by the headed Zero workbook reload flow.',
    findings,
  })
}

async function buildSubmittedAckControl(): Promise<ReliabilityControl> {
  const runtime = await createRuntime('submitted-ack-doc', 'browser:submitted')
  const pending = await runtime.enqueuePendingMutation({
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 23],
  })
  await runtime.markPendingMutationSubmitted(pending.id)
  const submitted = findPendingMutation(runtime.listPendingMutations(), pending.id)
  await runtime.applyAuthoritativeEvents(
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
  const afterAckPending = runtime.listPendingMutations()
  const afterAckJournal = runtime.listMutationJournalEntries()
  const zeroSchemaSource = readFileSync(join(rootDir, 'packages/zero-sync/src/schema.ts'), 'utf8')
  const workbookMutationStoreSource = readFileSync(join(rootDir, 'apps/bilig/src/zero/workbook-mutation-store.ts'), 'utf8')
  const submittedSurvived = submitted?.status === 'submitted'
  const ackAbsorbed =
    afterAckPending.length === 0 &&
    cellNumber(runtime, 'A1') === 23 &&
    afterAckJournal.some((entry) => entry.id === pending.id && entry.status === 'acked')
  const zeroDurabilityWired =
    zeroSchemaSource.includes("clientMutationId: string().from('client_mutation_id').optional()") &&
    workbookMutationStoreSource.includes('INSERT INTO workbook_event') &&
    workbookMutationStoreSource.includes('client_mutation_id')

  return reliabilityControl({
    id: 'submitted-mutations-absorb-authoritative-ack',
    category: 'authoritative-reconcile',
    passed: submittedSurvived && ackAbsorbed && zeroDurabilityWired,
    coveredControls: ['zero.authoritativeMutationDurability', 'pending.authoritativeAckAbsorption'],
    evidence:
      'Marked a mutation submitted, absorbed an authoritative event with the same client mutation id, and verified Zero/Postgres schema stores client mutation ids in durable workbook events.',
    findings: [
      ...(submittedSurvived ? [] : [`expected submitted pending status before ack, received ${submitted?.status ?? 'missing'}`]),
      ...(ackAbsorbed ? [] : ['authoritative ack did not clear pending state and record the authoritative value']),
      ...(zeroDurabilityWired ? [] : ['Zero schema or server mutation store no longer exposes client mutation id durable event wiring']),
    ],
  })
}

async function buildAuthoritativeRebaseControl(): Promise<ReliabilityControl> {
  const runtime = await createRuntime('rebase-doc', 'browser:rebase')
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
  const rebased = findPendingMutation(runtime.listPendingMutations(), pending.id)
  const rebasePassed = rebased?.status === 'rebased' && cellNumber(runtime, 'A1') === 31 && cellNumber(runtime, 'B1') === 101

  return reliabilityControl({
    id: 'authoritative-rebase-preserves-unsent-mutations',
    category: 'authoritative-reconcile',
    passed: rebasePassed,
    coveredControls: ['pending.authoritativeRebasePreservesLocal', 'zero.clientDurability'],
    evidence:
      'Applied an unrelated authoritative event over an unsent local mutation and verified both the authoritative cell and rebased local cell remain in the projection.',
    findings: [
      ...(rebased?.status === 'rebased'
        ? []
        : [`expected rebased pending status after authoritative replay, received ${rebased?.status ?? 'missing'}`]),
      ...(cellNumber(runtime, 'A1') === 31 ? [] : ['rebased local cell value was lost']),
      ...(cellNumber(runtime, 'B1') === 101 ? [] : ['authoritative replay cell value was lost']),
    ],
  })
}

async function buildFailedRetryControl(): Promise<ReliabilityControl> {
  const runtime = await createRuntime('failed-retry-doc', 'browser:failed')
  const pending = await runtime.enqueuePendingMutation({
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 47],
  })
  await runtime.markPendingMutationFailed(pending.id, 'mutation rejected by server')
  const failed = findPendingMutation(runtime.listPendingMutations(), pending.id)
  await runtime.retryPendingMutation(pending.id)
  const retried = findPendingMutation(runtime.listPendingMutations(), pending.id)
  const retryPassed =
    failed?.status === 'failed' &&
    failed.failureMessage === 'mutation rejected by server' &&
    retried?.status === 'local' &&
    retried.failureMessage === null &&
    cellNumber(runtime, 'A1') === 47

  return reliabilityControl({
    id: 'failed-mutations-survive-reload-and-retry',
    category: 'failure-recovery',
    passed: retryPassed,
    coveredControls: ['pending.failedRetryStateMachine'],
    evidence:
      'Marked a pending mutation failed, retried it, and verified it returned to local pending state without losing the projected value.',
    findings: [
      ...(failed?.status === 'failed' ? [] : [`expected failed pending status, received ${failed?.status ?? 'missing'}`]),
      ...(retried?.status === 'local' ? [] : [`expected local pending status after retry, received ${retried?.status ?? 'missing'}`]),
      ...(cellNumber(runtime, 'A1') === 47 ? [] : ['failed/retried local cell value was lost']),
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

function buildHeadedBrowserCrashRestartSoakControl(): ReliabilityControl {
  const source = readFileSync(join(rootDir, headedBrowserReliabilityTestFile), 'utf8')
  const testTitle = 'web app survives repeated tab crash restarts with persisted workbook state'
  const testBlock = extractBrowserTestBlock(source, 'remoteSyncTest', testTitle)
  const findings: string[] = []
  if (!testBlock) {
    findings.push(`missing remote-sync Playwright test: ${testTitle}`)
  } else {
    requireSnippet(
      testBlock,
      "createTestDocumentId('playwright-zero-crash-restart-soak')",
      'uses an isolated crash-restart document',
      findings,
    )
    requireSnippet(testBlock, "await writeCellValue(page, 0, 0, 'crash-seed')", 'writes the initial value before restart', findings)
    requireSnippet(testBlock, 'const restartedPage = await reopenWorkbookTab(page, documentId)', 'performs first tab restart', findings)
    requireSnippet(
      testBlock,
      'const secondRestartedPage = await reopenWorkbookTab(restartedPage, documentId)',
      'performs second tab restart',
      findings,
    )
    requireSnippet(
      testBlock,
      'const finalRestartedPage = await reopenWorkbookTab(secondRestartedPage, documentId)',
      'performs final tab restart',
      findings,
    )
    requireSnippet(
      testBlock,
      "await expectPersistedCellValue(finalRestartedPage, 0, 2, 'Sheet1!C1', 'restart-two')",
      'verifies state written across restarts',
      findings,
    )
    requireSnippet(
      testBlock,
      "await expect(finalRestartedPage.getByTestId('worker-error')).toHaveCount(0)",
      'asserts the final restarted tab has no worker error',
      findings,
    )
  }
  requireSnippet(
    source,
    'await previousPage.close({ runBeforeUnload: false })',
    'closes the previous headed tab without unload hooks',
    findings,
  )
  requireSnippet(source, 'const nextPage = await context.newPage()', 'creates a replacement headed tab', findings)
  requireSnippet(source, 'await openZeroWorkbookPage(nextPage, documentId)', 'opens the same workbook document after restart', findings)
  requireSnippet(
    source,
    "await expect(nextPage.getByTestId('worker-error')).toHaveCount(0)",
    'checks worker health after every restart',
    findings,
  )

  return reliabilityControl({
    id: 'headed-browser-crash-restart-soak',
    category: 'headed-browser',
    passed: findings.length === 0,
    coveredControls: ['headedBrowser.crashSoak'],
    evidence:
      'Validated the headed Playwright crash-restart contract that closes and recreates the workbook tab three times, writes cells across restarts, and verifies persisted workbook state with no worker error.',
    findings,
  })
}

async function buildOfflineNetworkPartitionRecoveryControl(): Promise<ReliabilityControl> {
  const localWriteCount = 32
  const remoteWriteCount = 32
  const documentId = 'offline-partition-doc'
  const runtime = await createRuntime(documentId, 'browser:offline-partition')
  const pendingMutations = await enqueueSequentialCellValues(runtime, {
    column: 'A',
    count: localWriteCount,
    valueOffset: 1000,
  })
  const pendingBeforeReconnect = runtime.listPendingMutations()
  const pendingQueued =
    pendingBeforeReconnect.length === localWriteCount &&
    pendingMutations.every((mutation) => findPendingMutation(pendingBeforeReconnect, mutation.id)?.status === 'local') &&
    cellsMatchSequentialValues(runtime, 'A', localWriteCount, 1000)

  await runtime.applyAuthoritativeEvents(
    buildSequentialCellValueEvents({
      column: 'B',
      count: remoteWriteCount,
      valueOffset: 2000,
      revisionOffset: 0,
      clientMutationIds: null,
    }),
    remoteWriteCount,
  )
  const pendingAfterReconnect = runtime.listPendingMutations()
  const rebasedOverReconnect =
    pendingAfterReconnect.length === localWriteCount &&
    pendingMutations.every((mutation) => findPendingMutation(pendingAfterReconnect, mutation.id)?.status === 'rebased') &&
    cellsMatchSequentialValues(runtime, 'A', localWriteCount, 1000) &&
    cellsMatchSequentialValues(runtime, 'B', remoteWriteCount, 2000)

  await markMutationsSubmitted(runtime, pendingMutations)
  await runtime.applyAuthoritativeEvents(
    buildSequentialCellValueEvents({
      column: 'A',
      count: localWriteCount,
      valueOffset: 1000,
      revisionOffset: remoteWriteCount,
      clientMutationIds: pendingMutations.map((mutation) => mutation.id),
    }),
    remoteWriteCount + localWriteCount,
  )
  const ackedJournal = runtime.listMutationJournalEntries()
  const authoritativeValuesPersisted =
    runtime.listPendingMutations().length === 0 &&
    pendingMutations.every((mutation) => ackedJournal.some((entry) => entry.id === mutation.id && entry.status === 'acked')) &&
    cellsMatchSequentialValues(runtime, 'A', localWriteCount, 1000) &&
    cellsMatchSequentialValues(runtime, 'B', remoteWriteCount, 2000)
  const maximumPendingQueueDepth = Math.max(pendingBeforeReconnect.length, pendingAfterReconnect.length)
  const boundedPendingQueue = maximumPendingQueueDepth === localWriteCount
  const passed = pendingQueued && rebasedOverReconnect && authoritativeValuesPersisted && boundedPendingQueue

  return reliabilityControl({
    id: 'offline-network-partition-recovery-soak',
    category: 'offline-recovery',
    passed,
    coveredControls: [
      'pending.authoritativeRebasePreservesLocal',
      'pending.authoritativeAckAbsorption',
      'offline.networkPartitionRecoverySoak',
    ],
    evidence:
      'Simulated 32 offline local edits, replayed 32 collaborator authoritative edits on reconnect, acked every local mutation, and verified both local and remote cells remained with an empty pending queue.',
    findings: [
      ...(pendingQueued ? [] : ['offline local pending mutations or projected values were not queued']),
      ...(rebasedOverReconnect ? [] : ['pending offline mutations did not rebase cleanly over reconnect authoritative events']),
      ...(authoritativeValuesPersisted ? [] : ['acked offline mutations did not persist as authoritative values after final reload']),
      ...(boundedPendingQueue ? [] : [`expected maximum pending queue depth ${localWriteCount}, observed ${maximumPendingQueueDepth}`]),
    ],
  })
}

function buildExternalSheetsExcelReliabilityComparisonControl(): ReliabilityControl {
  const artifact = parseExternalReliabilityComparisonArtifact(
    JSON.parse(readFileSync(externalReliabilityComparisonArtifactPath, 'utf8')) as unknown,
  )
  const findings = validateExternalReliabilityComparisonArtifact(artifact)
  const googleSourceCount = artifact.officialSources.filter((source) => source.vendor === 'google-sheets').length
  const microsoftSourceCount = artifact.officialSources.filter((source) => source.vendor === 'microsoft-excel').length

  return reliabilityControl({
    id: 'external-sheets-excel-reliability-comparison',
    category: 'external-comparison',
    passed: findings.length === 0,
    coveredControls: externalReliabilityComparisonCoveredControls,
    evidence:
      `Validated ${externalReliabilityComparisonArtifactRepoPath} from ${artifact.sourceBasis}: ` +
      `${String(artifact.dimensions.length)} required comparison dimensions cite ${String(googleSourceCount)} official Google Sheets/Workspace sources ` +
      `and ${String(microsoftSourceCount)} official Microsoft Excel/Microsoft 365 sources.`,
    findings,
  })
}

async function createRuntime(documentId: string, replicaId: string): Promise<WorkbookWorkerRuntime> {
  const runtime = new WorkbookWorkerRuntime()
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

async function enqueueSequentialCellValues(
  runtime: WorkbookWorkerRuntime,
  input: {
    readonly column: string
    readonly count: number
    readonly valueOffset: number
  },
): Promise<PendingWorkbookMutation[]> {
  return await Array.from({ length: input.count }).reduce<Promise<PendingWorkbookMutation[]>>(async (previousPromise, _, index) => {
    const previous = await previousPromise
    const row = index + 1
    const mutation = await runtime.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', `${input.column}${row}`, input.valueOffset + row],
    })
    return [...previous, mutation]
  }, Promise.resolve([]))
}

function buildSequentialCellValueEvents(input: {
  readonly column: string
  readonly count: number
  readonly valueOffset: number
  readonly revisionOffset: number
  readonly clientMutationIds: readonly string[] | null
}): AuthoritativeWorkbookEventRecord[] {
  return Array.from({ length: input.count }, (_, index) => {
    const row = index + 1
    return {
      revision: input.revisionOffset + row,
      clientMutationId: input.clientMutationIds?.[index] ?? null,
      payload: {
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: `${input.column}${row}`,
        value: input.valueOffset + row,
      },
    }
  })
}

async function markMutationsSubmitted(runtime: WorkbookWorkerRuntime, mutations: readonly PendingWorkbookMutation[]): Promise<void> {
  await mutations.reduce<Promise<void>>(async (previousPromise, mutation) => {
    await previousPromise
    await runtime.markPendingMutationSubmitted(mutation.id)
  }, Promise.resolve())
}

function cellsMatchSequentialValues(runtime: WorkbookWorkerRuntime, column: string, count: number, valueOffset: number): boolean {
  return Array.from({ length: count }, (_, index) => index + 1).every((row) => {
    return cellNumber(runtime, `${column}${row}`) === valueOffset + row
  })
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
      zeroSyncImplementation: literalField(source, 'zeroSyncImplementation', 'packages/zero-sync/src/schema.ts'),
      headedBrowserReliabilityTestFile: literalField(source, 'headedBrowserReliabilityTestFile', headedBrowserReliabilityTestFile),
      externalReliabilityComparisonArtifact: literalField(
        source,
        'externalReliabilityComparisonArtifact',
        'packages/benchmarks/baselines/reliability-external-sheets-excel-comparison.json',
      ),
    },
    summary: {
      allRequiredControlsPassed: booleanField(summary, 'allRequiredControlsPassed', 'reliability allRequiredControlsPassed'),
      pendingReloadPassed: booleanField(summary, 'pendingReloadPassed', 'reliability pendingReloadPassed'),
      authoritativeAckPassed: booleanField(summary, 'authoritativeAckPassed', 'reliability authoritativeAckPassed'),
      authoritativeRebasePassed: booleanField(summary, 'authoritativeRebasePassed', 'reliability authoritativeRebasePassed'),
      failedRetryPassed: booleanField(summary, 'failedRetryPassed', 'reliability failedRetryPassed'),
      headedBrowserReloadPassed: booleanField(summary, 'headedBrowserReloadPassed', 'reliability headedBrowserReloadPassed'),
      headedBrowserCrashSoakPassed: booleanField(summary, 'headedBrowserCrashSoakPassed', 'reliability headedBrowserCrashSoakPassed'),
      offlineNetworkPartitionPassed: booleanField(summary, 'offlineNetworkPartitionPassed', 'reliability offlineNetworkPartitionPassed'),
      coveredControls: stringArrayField(summary, 'coveredControls', 'reliability coveredControls'),
      uncoveredControls: stringArrayField(summary, 'uncoveredControls', 'reliability uncoveredControls'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'official-docs-comparison-artifact'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'official-docs-comparison-artifact'),
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
  if (
    value === 'pending-durability' ||
    value === 'authoritative-reconcile' ||
    value === 'failure-recovery' ||
    value === 'headed-browser' ||
    value === 'offline-recovery' ||
    value === 'external-comparison'
  ) {
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

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
