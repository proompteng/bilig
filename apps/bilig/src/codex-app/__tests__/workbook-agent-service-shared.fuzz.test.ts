import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import type { WorkbookAgentTimelineEntry } from '@bilig/contracts'
import {
  buildSnapshot,
  cloneUiContext,
  mergeTimelineEntries,
  normalizeExecutionPolicy,
  toContextRef,
  upsertEntry,
  upsertWorkflowRun,
  type WorkbookAgentThreadState,
} from '../workbook-agent-service-shared.js'

describe('workbook agent service shared state fuzz', () => {
  it('should preserve snapshot isolation across generated thread states', async () => {
    await runProperty({
      suite: 'bilig/codex-app/service-shared/snapshot-isolation',
      arbitrary: fc.record({
        scope: fc.constantFrom('private', 'shared'),
        policy: fc.option(fc.constantFrom('autoApplySafe', 'autoApplyAll', 'ownerReview'), { nil: null }),
        entries: fc.array(timelineEntryArbitrary, { maxLength: 5 }),
      }),
      predicate: async ({ scope, policy, entries }) => {
        const state = createThreadState({
          scope,
          executionPolicy: normalizeExecutionPolicy({ scope, requestedPolicy: policy }),
          entries,
        })
        const snapshot = buildSnapshot(state)

        state.durable.entries.push(createEntry('late-entry'))
        expect(snapshot.entries).toEqual(entries)
        expect(snapshot.executionPolicy).toBe(policy ?? (scope === 'shared' ? 'ownerReview' : 'autoApplyAll'))
      },
      parameters: { numRuns: 100 },
    })
  })

  it('should keep timeline and workflow upserts idempotent by id', async () => {
    await runProperty({
      suite: 'bilig/codex-app/service-shared/upsert-idempotency',
      arbitrary: fc.record({
        first: timelineEntryArbitrary,
        secondText: fc.string({ maxLength: 40 }),
      }),
      predicate: async ({ first, secondText }) => {
        const second = { ...first, text: secondText }
        expect(upsertEntry(upsertEntry([], first), second)).toEqual([second])
        expect(mergeTimelineEntries([first], [second])).toEqual([second])

        const workflow = createWorkflowRun(first.id)
        expect(upsertWorkflowRun(upsertWorkflowRun([], workflow), { ...workflow, status: 'completed' })).toEqual([
          { ...workflow, status: 'completed' },
        ])
      },
      parameters: { numRuns: 100 },
    })
  })

  it('should clone UI contexts and context refs without sharing nested mutable references', async () => {
    await runProperty({
      suite: 'bilig/codex-app/service-shared/context-clone-isolation',
      arbitrary: fc.record({
        rowStart: fc.integer({ min: 0, max: 10 }),
        colStart: fc.integer({ min: 0, max: 5 }),
      }),
      predicate: async ({ rowStart, colStart }) => {
        const context = {
          selection: { sheetName: 'Sheet1', address: 'B2', range: { startAddress: 'B2', endAddress: 'C3' } },
          viewport: { rowStart, rowEnd: rowStart + 10, colStart, colEnd: colStart + 5 },
        }

        const clone = cloneUiContext(context)
        const ref = toContextRef(context)
        context.selection.sheetName = 'Changed'

        expect(clone?.selection.sheetName).toBe('Sheet1')
        expect(ref?.selection.sheetName).toBe('Sheet1')
      },
      parameters: { numRuns: 100 },
    })
  })
})

// Helpers

const timelineEntryArbitrary = fc.record({
  id: fc.uuid(),
  kind: fc.constantFrom('user', 'assistant', 'plan', 'reasoning', 'tool', 'system'),
  turnId: fc.option(fc.uuid(), { nil: null }),
  text: fc.option(fc.string({ maxLength: 80 }), { nil: null }),
  phase: fc.option(fc.string({ maxLength: 20 }), { nil: null }),
  toolName: fc.option(fc.constantFrom('read_range', 'write_range'), { nil: null }),
  toolStatus: fc.option(fc.constantFrom('inProgress', 'completed', 'failed'), { nil: null }),
  argumentsText: fc.option(fc.string({ maxLength: 80 }), { nil: null }),
  outputText: fc.option(fc.string({ maxLength: 80 }), { nil: null }),
  success: fc.option(fc.boolean(), { nil: null }),
  citations: fc.constant([]),
}) satisfies fc.Arbitrary<WorkbookAgentTimelineEntry>

function createEntry(id: string): WorkbookAgentTimelineEntry {
  return {
    id,
    kind: 'system',
    turnId: null,
    text: 'late',
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [],
  }
}

function createThreadState(input: {
  scope: 'private' | 'shared'
  executionPolicy: WorkbookAgentThreadState['executionPolicy']
  entries: WorkbookAgentTimelineEntry[]
}): WorkbookAgentThreadState {
  return {
    documentId: 'doc-1',
    userId: 'alex@example.com',
    storageActorUserId: 'alex@example.com',
    scope: input.scope,
    executionPolicy: input.executionPolicy,
    threadId: 'thr-1',
    durable: {
      context: null,
      entries: [...input.entries],
      reviewQueueItems: [],
      executionRecords: [],
      workflowRuns: [],
    },
    live: {
      activeTurnId: null,
      status: 'idle',
      lastError: null,
      authorizedUserIds: new Set(['alex@example.com']),
      stagedPrivateBundleByTurn: new Map(),
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      turnActorUserIdByTurn: new Map(),
      turnContextByTurn: new Map(),
      lastAccessedAt: 1,
    },
  }
}

function createWorkflowRun(runId: string) {
  return {
    runId,
    threadId: 'thr-1',
    startedByUserId: 'alex@example.com',
    workflowTemplate: 'summarizeWorkbook' as const,
    title: 'Summary',
    summary: 'Summarize workbook',
    status: 'running' as const,
    createdAtUnixMs: 1,
    updatedAtUnixMs: 1,
    completedAtUnixMs: null,
    errorMessage: null,
    steps: [],
    artifact: null,
  }
}
