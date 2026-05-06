#!/usr/bin/env bun

import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { ValueTag } from '@bilig/protocol'
import { createMemoryWorkbookLocalStoreFactory } from '@bilig/storage-browser'
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

export interface CollaborationControl {
  readonly id: string
  readonly category: 'local-first-sync' | 'presence' | 'conflict-viewport'
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
  }
  readonly summary: {
    readonly allRequiredControlsPassed: boolean
    readonly syncRebaseAckPassed: boolean
    readonly presenceSelectionPassed: boolean
    readonly conflictViewportPassed: boolean
    readonly coveredControls: string[]
    readonly uncoveredControls: string[]
    readonly externalGoogleSheetsEvidence: 'not-captured'
    readonly externalMicrosoftExcelEvidence: 'not-captured'
  }
  readonly controls: CollaborationControl[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'collaboration-scorecard.json')
const requiredControlIds = [
  'worker-sync-rebase-ack-roundtrip',
  'presence-session-selection-filtering',
  'editor-conflict-and-viewport-protection',
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
] as const
const uncoveredControls = [
  'headedBrowser.multiUserViewportSoak',
  'conflictRateLongRunningCollaboration',
  'externalSheetsCollaborationComparison',
] as const

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
  const controls = [await buildWorkerSyncRebaseAckControl(), await buildPresenceSelectionControl(), buildConflictViewportControl()]
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
    },
    summary: {
      allRequiredControlsPassed: controls.filter((control) => control.required).every((control) => control.passed),
      syncRebaseAckPassed: requiredControl(controls, 'worker-sync-rebase-ack-roundtrip').passed,
      presenceSelectionPassed: requiredControl(controls, 'presence-session-selection-filtering').passed,
      conflictViewportPassed: requiredControl(controls, 'editor-conflict-and-viewport-protection').passed,
      coveredControls,
      uncoveredControls: [...uncoveredControls],
      externalGoogleSheetsEvidence: 'not-captured',
      externalMicrosoftExcelEvidence: 'not-captured',
    },
    controls,
  }
}

async function buildWorkerSyncRebaseAckControl(): Promise<CollaborationControl> {
  const runtime = new WorkbookWorkerRuntime({
    localStoreFactory: createMemoryWorkbookLocalStoreFactory(),
  })
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
    },
    summary: {
      allRequiredControlsPassed: booleanField(summary, 'allRequiredControlsPassed'),
      syncRebaseAckPassed: booleanField(summary, 'syncRebaseAckPassed'),
      presenceSelectionPassed: booleanField(summary, 'presenceSelectionPassed'),
      conflictViewportPassed: booleanField(summary, 'conflictViewportPassed'),
      coveredControls: stringArrayField(summary, 'coveredControls'),
      uncoveredControls: stringArrayField(summary, 'uncoveredControls'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'not-captured'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'not-captured'),
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
  if (value === 'local-first-sync' || value === 'presence' || value === 'conflict-viewport') {
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

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'collaboration-scorecard-'))
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
    throw new Error(`Unable to format generated collaboration scorecard: ${stderr}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
