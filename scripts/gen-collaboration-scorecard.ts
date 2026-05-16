#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { ValueTag } from '@bilig/protocol'
import { createInMemoryDocumentPersistence } from '@bilig/storage-server'
import { updatePresenceArgsSchema, type AuthoritativeWorkbookEventRecord } from '@bilig/zero-sync'
import type { ViewportPatch } from '@bilig/worker-transport'
import {
  closePresenceBackedWorkbookSession,
  countPresenceBackedWorkbookSessions,
  joinOwnedBrowserSession,
  openPresenceBackedWorkbookSession,
} from '../apps/bilig/src/workbook-runtime/document-presence-session-store.js'
import { applyProjectedViewportPatch, type ProjectedViewportPatchState } from '../apps/web/src/projected-viewport-patch-application.js'
import { normalizeWorkbookPresenceRows, selectActiveWorkbookCollaborators } from '../apps/web/src/workbook-presence-model.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../apps/web/src/workbook-optimistic-cell-flags.js'
import { parseEditorInput, parsedEditorInputMatchesSnapshot, sameCellContent } from '../apps/web/src/worker-workbook-app-model.js'
import { WorkbookWorkerRuntime } from '../apps/web/src/worker-runtime.js'
import { arrayField, asObject, booleanField, literalField, stringArrayField, stringField } from './json-scorecard-helpers.ts'
import {
  externalCollaborationComparisonArtifactRepoPath,
  externalCollaborationComparisonCoveredControls,
  parseExternalCollaborationComparisonArtifact,
  validateExternalCollaborationComparisonArtifact,
} from './collaboration-external-sheets-excel-comparison.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export interface CollaborationControl {
  readonly id: string
  readonly category: 'local-first-sync' | 'presence' | 'conflict-viewport' | 'headed-browser' | 'external-comparison'
  readonly required: boolean
  readonly passed: boolean
  readonly coveredControls: string[]
  readonly evidence: string
  readonly findings: string[]
}

export interface CollaborationScorecard {
  readonly schemaVersion: 1
  readonly suite: 'collaboration-posture'
  readonly generatedAt: string
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-collaboration-scorecard.ts'
    readonly workerRuntimeImplementation: 'apps/web/src/worker-runtime.ts'
    readonly presenceImplementation: 'apps/web/src/workbook-presence-model.ts'
    readonly presenceSessionImplementation: 'apps/bilig/src/workbook-runtime/document-presence-session-store.ts'
    readonly viewportPatchImplementation: 'apps/web/src/projected-viewport-patch-application.ts'
    readonly editorConflictImplementation: 'apps/web/src/use-workbook-editor-conflict.tsx'
    readonly headedBrowserViewportTestFile: 'e2e/tests/web-shell-scroll-performance.pw.ts'
    readonly externalCollaborationComparisonArtifact: 'packages/benchmarks/baselines/collaboration-external-sheets-excel-comparison.json'
  }
  readonly summary: {
    readonly allRequiredControlsPassed: boolean
    readonly syncRebaseAckPassed: boolean
    readonly presenceSelectionPassed: boolean
    readonly conflictViewportPassed: boolean
    readonly headedBrowserViewportPassed: boolean
    readonly longRunningConflictRatePassed: boolean
    readonly coveredControls: string[]
    readonly uncoveredControls: string[]
    readonly externalGoogleSheetsEvidence: 'official-docs-comparison-artifact'
    readonly externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact'
  }
  readonly controls: CollaborationControl[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'collaboration-scorecard.json')
const externalCollaborationComparisonArtifactPath = join(rootDir, externalCollaborationComparisonArtifactRepoPath)
const requiredControlIds = [
  'worker-sync-rebase-ack-roundtrip',
  'presence-session-selection-filtering',
  'editor-conflict-and-viewport-protection',
  'headed-browser-multi-user-viewport-soak',
  'long-running-collaboration-conflict-rate',
  'external-sheets-excel-collaboration-comparison',
] as const
const coveredControlOrder = [
  'sync.pendingRebase',
  'sync.authoritativeAck',
  'sync.noAcceptedOpLoss',
  'presence.sessionLifecycle',
  'presence.selectionSchema',
  'presence.collaboratorFiltering',
  'conflict.authoritativeDriftDetection',
  'viewport.optimisticAxisProtection',
  'viewport.authoritativeCatchupClearsPending',
  'headedBrowser.multiUserViewportSoak',
  'conflict.longRunningZeroUnexpectedConflicts',
  'sync.longRunningAcceptedOpConvergence',
  ...externalCollaborationComparisonCoveredControls,
] as const
const uncoveredControls: readonly string[] = []
const headedBrowserViewportTestFile = 'e2e/tests/web-shell-scroll-performance.pw.ts'

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error('Collaboration scorecard is missing. Run: bun scripts/gen-collaboration-scorecard.ts')
    }
    const scorecard = parseCollaborationScorecard(JSON.parse(readFileSync(outputPath, 'utf8')) as unknown)
    validateCollaborationScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = await buildCollaborationScorecard()
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export async function buildCollaborationScorecard(generatedAt = new Date().toISOString()): Promise<CollaborationScorecard> {
  const controls = [
    await buildWorkerSyncRebaseAckControl(),
    await buildPresenceSelectionControl(),
    buildConflictViewportControl(),
    buildHeadedBrowserMultiUserViewportSoakControl(),
    await buildLongRunningCollaborationConflictRateControl(),
    buildExternalSheetsExcelCollaborationComparisonControl(),
  ]
  const coveredControlSet = new Set(controls.flatMap((control) => control.coveredControls))
  const coveredControls = coveredControlOrder.filter((control) => coveredControlSet.has(control))

  return {
    schemaVersion: 1,
    suite: 'collaboration-posture',
    generatedAt,
    source: {
      artifactGenerator: 'scripts/gen-collaboration-scorecard.ts',
      workerRuntimeImplementation: 'apps/web/src/worker-runtime.ts',
      presenceImplementation: 'apps/web/src/workbook-presence-model.ts',
      presenceSessionImplementation: 'apps/bilig/src/workbook-runtime/document-presence-session-store.ts',
      viewportPatchImplementation: 'apps/web/src/projected-viewport-patch-application.ts',
      editorConflictImplementation: 'apps/web/src/use-workbook-editor-conflict.tsx',
      headedBrowserViewportTestFile,
      externalCollaborationComparisonArtifact: externalCollaborationComparisonArtifactRepoPath,
    },
    summary: {
      allRequiredControlsPassed: controls.filter((control) => control.required).every((control) => control.passed),
      syncRebaseAckPassed: requiredControl(controls, 'worker-sync-rebase-ack-roundtrip').passed,
      presenceSelectionPassed: requiredControl(controls, 'presence-session-selection-filtering').passed,
      conflictViewportPassed: requiredControl(controls, 'editor-conflict-and-viewport-protection').passed,
      headedBrowserViewportPassed: requiredControl(controls, 'headed-browser-multi-user-viewport-soak').passed,
      longRunningConflictRatePassed: requiredControl(controls, 'long-running-collaboration-conflict-rate').passed,
      coveredControls,
      uncoveredControls: [...uncoveredControls],
      externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
      externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
    },
    controls,
  }
}

async function buildWorkerSyncRebaseAckControl(): Promise<CollaborationControl> {
  const runtime = new WorkbookWorkerRuntime()
  await runtime.bootstrap({
    documentId: 'collaboration-sync-doc',
    replicaId: 'browser:self',
    persistState: true,
  })

  const pending = await Promise.all(
    Array.from({ length: 5 }, (_, index) =>
      runtime.enqueuePendingMutation({
        method: 'setCellValue',
        args: ['Sheet1', `A${String(index + 1)}`, (index + 1) * 10],
      }),
    ),
  )

  await runtime.applyAuthoritativeEvents(
    [1, 2, 3, 4, 5].map((revision) =>
      buildSetCellValueEvent({
        revision,
        address: `B${String(revision)}`,
        value: revision * 100,
      }),
    ),
    5,
  )
  const rebasedPending = runtime.listPendingMutations()
  const rebasePassed =
    rebasedPending.length === pending.length &&
    rebasedPending.every((mutation) => mutation.status === 'rebased') &&
    [1, 2, 3, 4, 5].every(
      (index) => cellNumber(runtime, `A${String(index)}`) === index * 10 && cellNumber(runtime, `B${String(index)}`) === index * 100,
    )

  await Promise.all(pending.map((mutation) => runtime.markPendingMutationSubmitted(mutation.id)))
  await runtime.applyAuthoritativeEvents(
    pending.map((mutation, index) =>
      buildSetCellValueEvent({
        revision: index + 6,
        address: `A${String(index + 1)}`,
        value: (index + 1) * 10,
        clientMutationId: mutation.id,
      }),
    ),
    10,
  )
  const ackPassed =
    runtime.listPendingMutations().length === 0 &&
    [1, 2, 3, 4, 5].every(
      (index) => cellNumber(runtime, `A${String(index)}`) === index * 10 && cellNumber(runtime, `B${String(index)}`) === index * 100,
    )
  runtime.dispose()

  return collaborationControl({
    id: 'worker-sync-rebase-ack-roundtrip',
    category: 'local-first-sync',
    passed: rebasePassed && ackPassed,
    coveredControls: ['sync.pendingRebase', 'sync.authoritativeAck', 'sync.noAcceptedOpLoss'],
    evidence:
      'Queued five local workbook mutations, replayed unrelated authoritative drift over them, verified rebased local and remote values, then absorbed authoritative acknowledgements without accepted-op loss.',
    findings: [
      ...(rebasePassed ? [] : [`pending rebase state or projected values were unexpected: ${JSON.stringify(rebasedPending)}`]),
      ...(ackPassed ? [] : ['authoritative acknowledgements did not clear pending state while preserving local and remote cell values']),
    ],
  })
}

async function buildPresenceSelectionControl(): Promise<CollaborationControl> {
  const persistence = createInMemoryDocumentPersistence()
  const selfSessionId = await openPresenceBackedWorkbookSession(persistence, 'collab-doc', 'browser:self')
  const otherSessionId = await openPresenceBackedWorkbookSession(persistence, 'collab-doc', 'browser:amy')
  await joinOwnedBrowserSession(persistence, 'bilig-app', 'collab-doc', 'browser:owner')
  const sessionsBeforeClose = await persistence.presence.sessions('collab-doc')
  const sessionCountBeforeClose = await countPresenceBackedWorkbookSessions(persistence, selfSessionId)
  await closePresenceBackedWorkbookSession(persistence, selfSessionId)
  const sessionsAfterClose = await persistence.presence.sessions('collab-doc')
  const validSelection = updatePresenceArgsSchema.safeParse({
    documentId: 'collab-doc',
    clientMutationId: 'presence-update-1',
    sessionId: otherSessionId,
    presenceClientId: 'presence:amy',
    sheetId: 1,
    sheetName: 'Sheet1',
    address: 'B7',
    selection: {
      sheetName: 'Sheet1',
      address: 'B7',
    },
  }).success
  const invalidSelectionRejected = !updatePresenceArgsSchema.safeParse({
    documentId: 'collab-doc',
    clientMutationId: 'presence-update-2',
    sessionId: otherSessionId,
    selection: {
      sheetName: '',
      address: '',
    },
  }).success
  const now = 1_000_000
  const rows = normalizeWorkbookPresenceRows([
    {
      sessionId: otherSessionId,
      userId: 'amy.smith@example.com',
      presenceClientId: 'presence:amy',
      sheetId: 1,
      sheetName: 'Sheet1',
      address: 'B7',
      selectionJson: {
        sheetName: 'Sheet1',
        address: 'B7',
      },
      updatedAt: now,
    },
    {
      sessionId: selfSessionId,
      userId: 'me@example.com',
      presenceClientId: 'presence:self',
      sheetId: 1,
      sheetName: 'Sheet1',
      address: 'A1',
      selectionJson: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      updatedAt: now,
    },
    {
      sessionId: 'collab-doc:browser:guest',
      userId: 'guest:cab7',
      presenceClientId: 'presence:guest',
      sheetId: 1,
      sheetName: 'Sheet1',
      address: 'C9',
      selectionJson: {
        sheetName: 'Sheet1',
        address: 'C9',
      },
      updatedAt: now,
    },
    {
      sessionId: 'collab-doc:browser:stale',
      userId: 'stale@example.com',
      presenceClientId: 'presence:stale',
      sheetId: 1,
      sheetName: 'Sheet1',
      address: 'D10',
      selectionJson: {
        sheetName: 'Sheet1',
        address: 'D10',
      },
      updatedAt: now - 60_000,
    },
    { malformed: true },
  ])
  const collaborators = selectActiveWorkbookCollaborators({
    rows,
    currentUserId: 'me@example.com',
    currentPresenceClientId: 'presence:self',
    currentSessionId: selfSessionId,
    knownSheetNames: ['Sheet1'],
    now,
  })
  const sessionLifecyclePassed =
    selfSessionId === 'collab-doc:browser:self' &&
    sessionsBeforeClose.length === 3 &&
    sessionCountBeforeClose === 3 &&
    sessionsAfterClose.length === 2 &&
    !sessionsAfterClose.includes(selfSessionId) &&
    sessionsAfterClose.includes(otherSessionId) &&
    (await persistence.ownership.owner('collab-doc')) === 'bilig-app'
  const collaboratorFilteringPassed =
    collaborators.length === 1 &&
    collaborators[0]?.label === 'Amy Smith' &&
    collaborators[0]?.sheetName === 'Sheet1' &&
    collaborators[0]?.address === 'B7'

  return collaborationControl({
    id: 'presence-session-selection-filtering',
    category: 'presence',
    passed: sessionLifecyclePassed && validSelection && invalidSelectionRejected && collaboratorFilteringPassed,
    coveredControls: ['presence.sessionLifecycle', 'presence.selectionSchema', 'presence.collaboratorFiltering'],
    evidence:
      'Opened and closed presence-backed sessions through the document runtime store, validated Zero presence selection payloads, and filtered collaborator rows to one active real user.',
    findings: [
      ...(sessionLifecyclePassed
        ? []
        : [
            `presence session lifecycle failed: before=${JSON.stringify(sessionsBeforeClose)}, after=${JSON.stringify(sessionsAfterClose)}`,
          ]),
      ...(validSelection ? [] : ['valid presence selection payload was rejected']),
      ...(invalidSelectionRejected ? [] : ['malformed presence selection payload was accepted']),
      ...(collaboratorFilteringPassed ? [] : [`unexpected collaborator rows: ${JSON.stringify(collaborators)}`]),
    ],
  })
}

function buildConflictViewportControl(): CollaborationControl {
  const baseSnapshot = cellSnapshot('A1', 10, 1)
  const authoritativeSnapshot = cellSnapshot('A1', 12, 2)
  const localDraft = parseEditorInput('99')
  const conflictDetected =
    !sameCellContent(baseSnapshot, authoritativeSnapshot) && !parsedEditorInputMatchesSnapshot(localDraft, authoritativeSnapshot)
  const matchingAuthoritativeDraftDoesNotConflict = parsedEditorInputMatchesSnapshot(parseEditorInput('12'), authoritativeSnapshot)

  const state = createPatchState()
  state.rowSizesBySheet.set('Sheet1', { 1: 22 })
  state.rowHeightsBySheet.set('Sheet1', { 1: 0 })
  state.hiddenRowsBySheet.set('Sheet1', { 1: true })
  state.pendingHiddenRowsBySheet.set('Sheet1', { 1: true })
  const stalePatchResult = applyProjectedViewportPatch({
    state,
    patch: {
      ...createViewportPatch(),
      rows: [{ index: 1, size: 22, hidden: false }],
    },
  })
  const optimisticAxisProtected =
    !stalePatchResult.rowsChanged &&
    state.rowHeightsBySheet.get('Sheet1')?.[1] === 0 &&
    state.hiddenRowsBySheet.get('Sheet1')?.[1] === true &&
    state.pendingHiddenRowsBySheet.get('Sheet1')?.[1] === true
  const catchupPatchResult = applyProjectedViewportPatch({
    state,
    patch: {
      ...createViewportPatch(),
      rows: [{ index: 1, size: 22, hidden: true }],
    },
  })
  const authoritativeCatchupPassed =
    !catchupPatchResult.rowsChanged &&
    state.rowHeightsBySheet.get('Sheet1')?.[1] === 0 &&
    state.hiddenRowsBySheet.get('Sheet1')?.[1] === true &&
    state.pendingHiddenRowsBySheet.get('Sheet1') === undefined

  return collaborationControl({
    id: 'editor-conflict-and-viewport-protection',
    category: 'conflict-viewport',
    passed: conflictDetected && matchingAuthoritativeDraftDoesNotConflict && optimisticAxisProtected && authoritativeCatchupPassed,
    coveredControls: [
      'conflict.authoritativeDriftDetection',
      'viewport.optimisticAxisProtection',
      'viewport.authoritativeCatchupClearsPending',
    ],
    evidence:
      'Exercised editor conflict model helpers against authoritative cell drift, then applied stale and confirming viewport patches to prove optimistic row state survives until authoritative catch-up.',
    findings: [
      ...(conflictDetected ? [] : ['editor conflict model did not detect authoritative drift against a different local draft']),
      ...(matchingAuthoritativeDraftDoesNotConflict ? [] : ['editor conflict model rejected a local draft matching authoritative state']),
      ...(optimisticAxisProtected ? [] : ['stale viewport patch overwrote pending optimistic row visibility state']),
      ...(authoritativeCatchupPassed ? [] : ['authoritative viewport catch-up did not clear pending optimistic row state']),
    ],
  })
}

function buildHeadedBrowserMultiUserViewportSoakControl(): CollaborationControl {
  const source = readFileSync(join(rootDir, headedBrowserViewportTestFile), 'utf8')
  const testTitle = 'keeps shell surfaces quiet and coalesces visible collaborator patch churn while browsing'
  const testBlock = extractBrowserTestBlock(source, 'remoteSyncTest', testTitle)
  const findings: string[] = []
  if (!testBlock) {
    findings.push(`missing remote sync Playwright test: ${testTitle}`)
  } else {
    requireSnippet(testBlock, 'const mirrorPage = await page.context().newPage()', 'opens a second browser tab', findings)
    requireSnippet(testBlock, 'benchmarkCorpus=wide-mixed-250k', 'loads the 250k benchmark corpus in both tabs', findings)
    requireSnippet(
      testBlock,
      'await Promise.all([waitForWorkbookReady(page), waitForWorkbookReady(mirrorPage)])',
      'waits for both workbooks',
      findings,
    )
    requireSnippet(
      testBlock,
      'await Promise.all([waitForBenchmarkCorpus(page), waitForBenchmarkCorpus(mirrorPage)])',
      'waits for both corpora',
      findings,
    )
    requireSnippet(
      testBlock,
      "warmStartWorkbookScrollPerf(page, 'wide-250k-browse-with-visible-patches')",
      'warms the sampled viewport',
      findings,
    )
    requireSnippet(
      testBlock,
      'performHorizontalGridBrowse(page, { distancePx: 2_560, steps: 140 })',
      'browses the viewport while sampling',
      findings,
    )
    requireSnippet(testBlock, 'emitRemoteEdits()', 'applies collaborator edits during browse', findings)
    requireSnippet(
      testBlock,
      "testInfo.outputPath('scroll-perf-wide-250k-visible-patches.json')",
      'writes the headed collaboration perf artifact',
      findings,
    )
    requireSnippet(testBlock, 'expectBoundedVisibleMutation(report', 'asserts bounded visible patch churn', findings)
    requireSnippet(testBlock, 'minDamagePatches: 1', 'requires at least one visible collaborator patch', findings)
    requireSnippet(testBlock, 'maxRendererVisibleDirtyTiles: 24', 'bounds dirty visible tiles', findings)
    requireSnippet(testBlock, 'expectQuietShell(report)', 'keeps shell surfaces quiet', findings)
  }

  return collaborationControl({
    id: 'headed-browser-multi-user-viewport-soak',
    category: 'headed-browser',
    passed: findings.length === 0,
    coveredControls: ['headedBrowser.multiUserViewportSoak'],
    evidence:
      'Validated the headed Playwright remote-sync performance contract for two tabs browsing a 250k workbook while collaborator edits create visible viewport patches.',
    findings,
  })
}

async function buildLongRunningCollaborationConflictRateControl(): Promise<CollaborationControl> {
  const runtime = new WorkbookWorkerRuntime()
  await runtime.bootstrap({
    documentId: 'collaboration-long-running-doc',
    replicaId: 'browser:self',
    persistState: true,
  })

  const localWrites: Array<{ id: string; address: string; value: number }> = []
  const remoteWrites: Array<{ address: string; value: number }> = []
  const failedMutationIds = new Set<string>()
  let revision = 0
  let maxPending = 0

  const captureFailedPending = () => {
    const pending = runtime.listPendingMutations()
    maxPending = Math.max(maxPending, pending.length)
    for (const mutation of pending) {
      if (mutation.status === 'failed') {
        failedMutationIds.add(mutation.id)
      }
    }
  }

  const acknowledgeAllPending = async () => {
    const pending = runtime.listPendingMutations()
    await Promise.all(pending.map((mutation) => runtime.markPendingMutationSubmitted(mutation.id)))
    const ackEvents = pending.map((mutation) => {
      const write = localWrites.find((candidate) => candidate.id === mutation.id)
      if (!write) {
        throw new Error(`Missing local write for pending mutation ${mutation.id}`)
      }
      revision += 1
      return buildSetCellValueEvent({
        revision,
        address: write.address,
        value: write.value,
        clientMutationId: mutation.id,
      })
    })
    await runtime.applyAuthoritativeEvents(ackEvents, revision)
    captureFailedPending()
  }

  try {
    const rounds = 64
    const ackEvery = 8
    await Array.from({ length: rounds }, (_, index) => index).reduce<Promise<void>>(async (previous, index) => {
      await previous
      const localAddress = `A${String(index + 1)}`
      const localValue = 10_000 + index
      const mutation = await runtime.enqueuePendingMutation({
        method: 'setCellValue',
        args: ['Sheet1', localAddress, localValue],
      })
      localWrites.push({ id: mutation.id, address: localAddress, value: localValue })

      revision += 1
      const remoteAddress = `B${String(index + 1)}`
      const remoteValue = 20_000 + index
      remoteWrites.push({ address: remoteAddress, value: remoteValue })
      await runtime.applyAuthoritativeEvents([buildSetCellValueEvent({ revision, address: remoteAddress, value: remoteValue })], revision)
      captureFailedPending()

      if ((index + 1) % ackEvery === 0) {
        await acknowledgeAllPending()
      }
    }, Promise.resolve())
    await acknowledgeAllPending()

    const pending = runtime.listPendingMutations()
    const journal = runtime.listMutationJournalEntries()
    const allLocalWritesAcked = localWrites.every((write) =>
      journal.some((entry) => entry.id === write.id && entry.status === 'acked' && cellNumber(runtime, write.address) === write.value),
    )
    const allRemoteWritesVisible = remoteWrites.every((write) => cellNumber(runtime, write.address) === write.value)
    const operationCount = localWrites.length + remoteWrites.length
    const unexpectedConflictRate = operationCount === 0 ? 0 : failedMutationIds.size / operationCount
    const convergencePassed =
      pending.length === 0 &&
      localWrites.length === 64 &&
      remoteWrites.length === 64 &&
      maxPending === 8 &&
      allLocalWritesAcked &&
      allRemoteWritesVisible
    const conflictRatePassed = unexpectedConflictRate === 0

    return collaborationControl({
      id: 'long-running-collaboration-conflict-rate',
      category: 'local-first-sync',
      passed: convergencePassed && conflictRatePassed,
      coveredControls: ['conflict.longRunningZeroUnexpectedConflicts', 'sync.longRunningAcceptedOpConvergence'],
      evidence:
        'Executed 64 local writes interleaved with 64 collaborator authoritative writes, rebased pending operations across each collaborator edit, acknowledged every local mutation in batches, and verified zero unexpected failed mutations plus final convergence.',
      findings: [
        ...(conflictRatePassed ? [] : [`unexpected conflict rate was ${String(unexpectedConflictRate)}`]),
        ...(pending.length === 0 ? [] : [`pending mutations remained after long-running acknowledgement: ${pending.length}`]),
        ...(allLocalWritesAcked ? [] : ['not all local accepted operations were acknowledged and visible']),
        ...(allRemoteWritesVisible ? [] : ['not all collaborator authoritative operations were visible after convergence']),
        ...(maxPending === 8 ? [] : [`expected max pending batch size 8, received ${String(maxPending)}`]),
      ],
    })
  } finally {
    runtime.dispose()
  }
}

function extractBrowserTestBlock(source: string, testFunctionName: 'test' | 'remoteSyncTest', testTitle: string): string | null {
  const marker = `${testFunctionName}('${testTitle}'`
  const start = source.indexOf(marker)
  if (start < 0) {
    return null
  }
  const endCandidates = ['\n  test(', '\n  remoteSyncTest(']
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

function buildExternalSheetsExcelCollaborationComparisonControl(): CollaborationControl {
  const artifact = parseExternalCollaborationComparisonArtifact(
    JSON.parse(readFileSync(externalCollaborationComparisonArtifactPath, 'utf8')) as unknown,
  )
  const findings = validateExternalCollaborationComparisonArtifact(artifact)
  const googleSourceCount = artifact.officialSources.filter((source) => source.vendor === 'google-sheets').length
  const microsoftSourceCount = artifact.officialSources.filter((source) => source.vendor === 'microsoft-excel').length

  return collaborationControl({
    id: 'external-sheets-excel-collaboration-comparison',
    category: 'external-comparison',
    passed: findings.length === 0,
    coveredControls: externalCollaborationComparisonCoveredControls,
    evidence:
      `Validated ${externalCollaborationComparisonArtifactRepoPath} from ${artifact.sourceBasis}: ` +
      `${String(artifact.dimensions.length)} required comparison dimensions cite ${String(googleSourceCount)} official Google Sheets/Workspace sources ` +
      `and ${String(microsoftSourceCount)} official Microsoft Excel/Microsoft 365 sources.`,
    findings,
  })
}

function buildSetCellValueEvent(input: {
  readonly revision: number
  readonly address: string
  readonly value: number
  readonly clientMutationId?: string | null
}): AuthoritativeWorkbookEventRecord {
  return {
    revision: input.revision,
    clientMutationId: input.clientMutationId ?? null,
    payload: {
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: input.address,
      value: input.value,
    },
  }
}

function cellNumber(runtime: WorkbookWorkerRuntime, address: string): number | null {
  const value = runtime.getCell('Sheet1', address).value
  return value.tag === ValueTag.Number ? value.value : null
}

function cellSnapshot(address: string, value: number, version: number) {
  return {
    sheetName: 'Sheet1',
    address,
    input: value,
    value: {
      tag: ValueTag.Number,
      value,
    },
    flags: 0,
    version,
  }
}

function createPatchState(): ProjectedViewportPatchState {
  return {
    cellSnapshots: new Map(),
    cellKeysBySheet: new Map(),
    cellStyles: new Map([['style-0', { id: 'style-0' }]]),
    columnSizesBySheet: new Map(),
    columnWidthsBySheet: new Map(),
    pendingColumnWidthsBySheet: new Map(),
    pendingHiddenColumnsBySheet: new Map(),
    rowSizesBySheet: new Map(),
    rowHeightsBySheet: new Map(),
    pendingRowHeightsBySheet: new Map(),
    pendingHiddenRowsBySheet: new Map(),
    hiddenColumnsBySheet: new Map(),
    hiddenRowsBySheet: new Map(),
    freezeRowsBySheet: new Map(),
    freezeColsBySheet: new Map(),
    mergeRangesBySheet: new Map(),
    knownSheets: new Set(),
  }
}

function createViewportPatch(): ViewportPatch {
  return {
    version: 1,
    full: false,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 10,
      colStart: 0,
      colEnd: 4,
    },
    metrics: {
      batchId: 0,
      changedInputCount: 0,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0,
    },
    styles: [],
    cells: [
      {
        row: 0,
        col: 0,
        snapshot: {
          sheetName: 'Sheet1',
          address: 'A1',
          value: { tag: ValueTag.Number, value: 99 },
          flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
          input: 99,
          version: 1,
        },
        displayText: '99',
        copyText: '99',
        editorText: '99',
        formatId: 0,
        styleId: 'style-0',
      },
    ],
    columns: [],
    rows: [],
  }
}

function collaborationControl(input: {
  readonly id: CollaborationControl['id']
  readonly category: CollaborationControl['category']
  readonly passed: boolean
  readonly coveredControls: readonly string[]
  readonly evidence: string
  readonly findings: readonly string[]
}): CollaborationControl {
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

function requiredControl(controls: readonly CollaborationControl[], id: string): CollaborationControl {
  const entry = controls.find((control) => control.id === id)
  if (!entry) {
    throw new Error(`Collaboration scorecard is missing required control: ${id}`)
  }
  return entry
}

export function parseCollaborationScorecard(value: unknown): CollaborationScorecard {
  const record = asObject(value, 'collaboration scorecard')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'collaboration-posture') {
    throw new Error('Unexpected collaboration scorecard header')
  }
  const source = asObject(record['source'], 'collaboration source')
  const summary = asObject(record['summary'], 'collaboration summary')
  return {
    schemaVersion: 1,
    suite: 'collaboration-posture',
    generatedAt: stringField(record, 'generatedAt'),
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-collaboration-scorecard.ts'),
      workerRuntimeImplementation: literalField(source, 'workerRuntimeImplementation', 'apps/web/src/worker-runtime.ts'),
      presenceImplementation: literalField(source, 'presenceImplementation', 'apps/web/src/workbook-presence-model.ts'),
      presenceSessionImplementation: literalField(
        source,
        'presenceSessionImplementation',
        'apps/bilig/src/workbook-runtime/document-presence-session-store.ts',
      ),
      viewportPatchImplementation: literalField(
        source,
        'viewportPatchImplementation',
        'apps/web/src/projected-viewport-patch-application.ts',
      ),
      editorConflictImplementation: literalField(source, 'editorConflictImplementation', 'apps/web/src/use-workbook-editor-conflict.tsx'),
      headedBrowserViewportTestFile: literalField(source, 'headedBrowserViewportTestFile', headedBrowserViewportTestFile),
      externalCollaborationComparisonArtifact: literalField(
        source,
        'externalCollaborationComparisonArtifact',
        'packages/benchmarks/baselines/collaboration-external-sheets-excel-comparison.json',
      ),
    },
    summary: {
      allRequiredControlsPassed: booleanField(summary, 'allRequiredControlsPassed'),
      syncRebaseAckPassed: booleanField(summary, 'syncRebaseAckPassed'),
      presenceSelectionPassed: booleanField(summary, 'presenceSelectionPassed'),
      conflictViewportPassed: booleanField(summary, 'conflictViewportPassed'),
      headedBrowserViewportPassed: booleanField(summary, 'headedBrowserViewportPassed'),
      longRunningConflictRatePassed: booleanField(summary, 'longRunningConflictRatePassed'),
      coveredControls: stringArrayField(summary, 'coveredControls'),
      uncoveredControls: stringArrayField(summary, 'uncoveredControls'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'official-docs-comparison-artifact'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'official-docs-comparison-artifact'),
    },
    controls: arrayField(record, 'controls').map(parseCollaborationControl),
  }
}

function parseCollaborationControl(value: unknown): CollaborationControl {
  const record = asObject(value, 'collaboration control')
  return {
    id: stringField(record, 'id'),
    category: parseCollaborationCategory(stringField(record, 'category')),
    required: booleanField(record, 'required'),
    passed: booleanField(record, 'passed'),
    coveredControls: stringArrayField(record, 'coveredControls'),
    evidence: stringField(record, 'evidence'),
    findings: stringArrayField(record, 'findings'),
  }
}

export function validateCollaborationScorecard(scorecard: CollaborationScorecard): void {
  for (const id of requiredControlIds) {
    const control = requiredControl(scorecard.controls, id)
    if (!control.required) {
      throw new Error(`Collaboration scorecard required control is not marked required: ${id}`)
    }
    if (!control.passed) {
      throw new Error(`Collaboration scorecard contains a failed required control: ${id}`)
    }
  }
  if (!scorecard.summary.allRequiredControlsPassed) {
    throw new Error('Collaboration scorecard summary reports failed required controls')
  }
  for (const control of coveredControlOrder) {
    if (!scorecard.summary.coveredControls.includes(control)) {
      throw new Error(`Collaboration scorecard is missing covered control: ${control}`)
    }
  }
  for (const control of uncoveredControls) {
    if (!scorecard.summary.uncoveredControls.includes(control)) {
      throw new Error(`Collaboration scorecard is missing uncovered control disclosure: ${control}`)
    }
  }
}

function parseCollaborationCategory(value: string): CollaborationControl['category'] {
  if (
    value === 'local-first-sync' ||
    value === 'presence' ||
    value === 'conflict-viewport' ||
    value === 'headed-browser' ||
    value === 'external-comparison'
  ) {
    return value
  }
  throw new Error(`Unexpected collaboration category: ${value}`)
}

function logResult(mode: 'check' | 'write', scorecard: CollaborationScorecard): void {
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

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
