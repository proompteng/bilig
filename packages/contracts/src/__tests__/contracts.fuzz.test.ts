import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { cloneJsonValue, corruptRecord, fuzzCellRangeRefArbitrary, fuzzWorkbookAddressArbitrary, runProperty } from '@bilig/test-fuzz'
import {
  decodeUnknownSync,
  RuntimeSessionSchema,
  stringifyWorkbookAgentUiContextSemanticKey,
  WorkbookAgentStreamEventSchema,
  WorkbookAgentTimelineEntrySchema,
  WorkbookAgentUiContextSchema,
  WorkbookAgentWorkflowRunSchema,
  type WorkbookAgentUiContext,
} from '../index.js'

describe('contract schema fuzz', () => {
  it('should decode generated runtime, ui context, timeline, workflow, and stream payloads', async () => {
    await runProperty({
      suite: 'contracts/schema/generated-workbook-agent-payloads',
      arbitrary: fc.record({
        session: runtimeSessionArbitrary,
        context: workbookAgentUiContextArbitrary,
        timeline: timelineEntryArbitrary,
        workflow: workflowRunArbitrary,
        streamDelta: fc.string({ maxLength: 40 }),
      }),
      predicate: async ({ session, context, timeline, workflow, streamDelta }) => {
        expect(decodeUnknownSync(RuntimeSessionSchema, cloneJsonValue(session))).toEqual(session)
        expect(decodeUnknownSync(WorkbookAgentUiContextSchema, cloneJsonValue(context))).toEqual(context)
        expect(decodeUnknownSync(WorkbookAgentTimelineEntrySchema, cloneJsonValue(timeline))).toEqual(timeline)
        expect(decodeUnknownSync(WorkbookAgentWorkflowRunSchema, cloneJsonValue(workflow))).toEqual(workflow)
        expect(
          decodeUnknownSync(WorkbookAgentStreamEventSchema, {
            type: 'entryTextDelta',
            entryKind: 'assistant',
            itemId: timeline.id,
            turnId: 'turn-1',
            delta: streamDelta,
          }),
        ).toEqual({
          type: 'entryTextDelta',
          entryKind: 'assistant',
          itemId: timeline.id,
          turnId: 'turn-1',
          delta: streamDelta,
        })
      },
      parameters: { numRuns: 80 },
    })
  })

  it('should reject required-field corruptions in generated contract payloads', async () => {
    await runProperty({
      suite: 'contracts/schema/reject-required-field-corruption',
      arbitrary: fc.record({
        context: workbookAgentUiContextArbitrary,
        timeline: timelineEntryArbitrary,
        workflow: workflowRunArbitrary,
      }),
      predicate: async ({ context, timeline, workflow }) => {
        expect(() => decodeUnknownSync(WorkbookAgentUiContextSchema, corruptRecord(context, 'viewport'))).toThrow()
        expect(() => decodeUnknownSync(WorkbookAgentTimelineEntrySchema, corruptRecord(timeline, 'kind'))).toThrow()
        expect(() => decodeUnknownSync(WorkbookAgentWorkflowRunSchema, corruptRecord(workflow, 'status', 'done'))).toThrow()
      },
      parameters: { numRuns: 80 },
    })
  })

  it('should ignore string pool ids in workbook agent rendered context semantic keys', async () => {
    await runProperty({
      suite: 'contracts/context-key/string-id-stability',
      arbitrary: fc.record({
        context: workbookAgentUiContextArbitrary,
        stringValue: fc.string({ maxLength: 40 }),
        leftStringId: fc.integer({ min: 0, max: 1_000_000 }),
        rightStringId: fc.integer({ min: 0, max: 1_000_000 }),
      }),
      predicate: async ({ context, stringValue, leftStringId, rightStringId }) => {
        const left = withRenderedStringValue(context, stringValue, leftStringId)
        const right = withRenderedStringValue(context, stringValue, rightStringId)

        expect(stringifyWorkbookAgentUiContextSemanticKey(left)).toBe(stringifyWorkbookAgentUiContextSemanticKey(right))
      },
      parameters: { numRuns: 80 },
    })
  })
})

// Helpers

const runtimeSessionArbitrary = fc.record({
  authToken: fc.uuid(),
  userId: fc.emailAddress(),
  roles: fc.array(fc.constantFrom('viewer', 'editor', 'owner'), { maxLength: 3 }),
  isAuthenticated: fc.boolean(),
  authSource: fc.constantFrom('header', 'cookie', 'guest'),
})

const viewportArbitrary = fc
  .record({
    rowStart: fc.integer({ min: 0, max: 50 }),
    colStart: fc.integer({ min: 0, max: 20 }),
    rowSpan: fc.integer({ min: 0, max: 40 }),
    colSpan: fc.integer({ min: 0, max: 20 }),
  })
  .map(({ rowStart, colStart, rowSpan, colSpan }) => ({
    rowStart,
    rowEnd: rowStart + rowSpan,
    colStart,
    colEnd: colStart + colSpan,
  }))

const workbookAgentUiContextArbitrary: fc.Arbitrary<WorkbookAgentUiContext> = fc
  .record({
    range: fuzzCellRangeRefArbitrary,
    address: fuzzWorkbookAddressArbitrary,
    viewport: viewportArbitrary,
  })
  .map(({ range, address, viewport }) => ({
    selection: {
      sheetName: range.sheetName,
      address,
      range: {
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      },
    },
    viewport,
  }))

const citationArbitrary = fc.oneof(
  fuzzCellRangeRefArbitrary.map((range) => ({
    kind: 'range' as const,
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
    role: 'target' as const,
  })),
  fc.integer({ min: 0, max: 10_000 }).map((revision) => ({
    kind: 'revision' as const,
    revision,
  })),
)

const timelineEntryArbitrary = fc.record({
  id: fc.uuid(),
  kind: fc.constantFrom('user', 'assistant', 'plan', 'reasoning', 'tool', 'system'),
  turnId: fc.option(fc.uuid(), { nil: null }),
  text: fc.option(fc.string({ maxLength: 80 }), { nil: null }),
  phase: fc.option(fc.string({ maxLength: 20 }), { nil: null }),
  toolName: fc.option(fc.constantFrom('read_range', 'write_range', 'verify_invariants'), { nil: null }),
  toolStatus: fc.option(fc.constantFrom('inProgress', 'completed', 'failed'), { nil: null }),
  argumentsText: fc.option(fc.string({ maxLength: 80 }), { nil: null }),
  outputText: fc.option(fc.string({ maxLength: 80 }), { nil: null }),
  success: fc.option(fc.boolean(), { nil: null }),
  citations: fc.array(citationArbitrary, { maxLength: 3 }),
})

const workflowStepArbitrary = fc.record({
  stepId: fc.uuid(),
  label: fc.string({ minLength: 1, maxLength: 30 }),
  status: fc.constantFrom('pending', 'running', 'completed', 'failed', 'cancelled'),
  summary: fc.string({ maxLength: 80 }),
  updatedAtUnixMs: fc.integer({ min: 0, max: 10_000 }),
})

const workflowRunArbitrary = fc.record({
  runId: fc.uuid(),
  threadId: fc.uuid(),
  startedByUserId: fc.emailAddress(),
  workflowTemplate: fc.constantFrom('summarizeWorkbook', 'findFormulaIssues', 'searchWorkbookQuery', 'createSheet'),
  title: fc.string({ minLength: 1, maxLength: 30 }),
  summary: fc.string({ maxLength: 80 }),
  status: fc.constantFrom('running', 'completed', 'failed', 'cancelled'),
  createdAtUnixMs: fc.integer({ min: 0, max: 10_000 }),
  updatedAtUnixMs: fc.integer({ min: 0, max: 10_000 }),
  completedAtUnixMs: fc.option(fc.integer({ min: 0, max: 10_000 }), { nil: null }),
  errorMessage: fc.option(fc.string({ maxLength: 80 }), { nil: null }),
  steps: fc.array(workflowStepArbitrary, { maxLength: 3 }),
  artifact: fc.option(
    fc.record({
      kind: fc.constant('markdown' as const),
      title: fc.string({ minLength: 1, maxLength: 30 }),
      text: fc.string({ maxLength: 120 }),
    }),
    { nil: null },
  ),
})

function withRenderedStringValue(context: WorkbookAgentUiContext, value: string, stringId: number): WorkbookAgentUiContext {
  const cellValue = { tag: 3, value, stringId }
  const cell = {
    address: context.selection.address,
    input: value,
    value: cellValue,
    formula: null,
    displayFormat: null,
    styleId: null,
    numberFormatId: null,
    style: null,
  }
  const renderedRange = {
    range: {
      sheetName: context.selection.sheetName,
      startAddress: context.selection.address,
      endAddress: context.selection.address,
    },
    rowCount: 1,
    columnCount: 1,
    cellCount: 1,
    truncated: false,
    rows: [[cell]],
  }
  return {
    ...context,
    rendered: {
      capturedAtUnixMs: 1,
      capturedRevision: 1,
      batchId: null,
      selection: renderedRange,
      visibleRange: renderedRange,
    },
  }
}
