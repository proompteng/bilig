import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { fuzzCellRangeRefArbitrary, runProperty } from '@bilig/test-fuzz'
import type { WorkbookAgentStructuralCommand } from '../workbook-agent-structural-commands.js'
import {
  deriveWorkbookAgentStructuralCommandPreviewRanges,
  describeWorkbookAgentStructuralCommand,
  estimateWorkbookAgentStructuralCommandAffectedCells,
  isHighRiskWorkbookAgentStructuralCommand,
  isWorkbookAgentStructuralCommandKind,
  isWorkbookAgentStructuralCommandValue,
  isWorkbookScopeStructuralCommand,
} from '../workbook-agent-structural-commands.js'

describe('workbook agent structural command fuzz', () => {
  it('should keep structural guards, classification, descriptions, and preview ranges coherent', async () => {
    await runProperty({
      suite: 'agent-api/structural-commands/classification-preview-coherence',
      arbitrary: structuralCommandArbitrary,
      predicate: async (command) => {
        expect(isWorkbookAgentStructuralCommandKind(command.kind)).toBe(true)
        expect(isWorkbookAgentStructuralCommandValue(command)).toBe(true)
        expect(describeWorkbookAgentStructuralCommand(command).trim()).not.toBe('')

        const ranges = deriveWorkbookAgentStructuralCommandPreviewRanges(command)
        for (const range of ranges) {
          expect(range.sheetName).toBeTruthy()
          expect(range.startAddress).toMatch(/^[A-Z]+[1-9]\d*$/u)
          expect(range.endAddress).toMatch(/^[A-Z]+[1-9]\d*$/u)
          expect(range.role).toBe('target')
        }

        const affectedCells = estimateWorkbookAgentStructuralCommandAffectedCells(command)
        expect(affectedCells === null || Number.isSafeInteger(affectedCells)).toBe(true)
        expect(affectedCells === null || affectedCells > 0).toBe(true)

        if (isWorkbookScopeStructuralCommand(command)) {
          expect(
            command.kind === 'createSheet' ||
              command.kind === 'renameSheet' ||
              command.kind === 'deleteSheet' ||
              command.kind === 'insertRows' ||
              command.kind === 'deleteRows' ||
              command.kind === 'insertColumns' ||
              command.kind === 'deleteColumns',
          ).toBe(true)
        }
        if (isHighRiskWorkbookAgentStructuralCommand(command)) {
          expect(
            command.kind === 'createSheet' ||
              command.kind === 'renameSheet' ||
              command.kind === 'deleteSheet' ||
              command.kind === 'insertRows' ||
              command.kind === 'deleteRows' ||
              command.kind === 'insertColumns' ||
              command.kind === 'deleteColumns',
          ).toBe(true)
        }
      },
      parameters: { numRuns: 120 },
    })
  })

  it('should reject structural command corruptions at required fields', async () => {
    await runProperty({
      suite: 'agent-api/structural-commands/reject-corruption',
      arbitrary: structuralCommandArbitrary,
      predicate: async (command) => {
        expect(isWorkbookAgentStructuralCommandValue({ ...command, kind: 'unsupported' })).toBe(false)
        expect(isWorkbookAgentStructuralCommandValue(corruptCommand(command))).toBe(false)
      },
      parameters: { numRuns: 120 },
    })
  })
})

// Helpers

const structuralCommandArbitrary: fc.Arbitrary<WorkbookAgentStructuralCommand> = fc.oneof(
  fc.record({ name: fc.constantFrom('Sheet1', 'Revenue', 'Inputs') }).map(({ name }) => ({ kind: 'createSheet' as const, name })),
  fc
    .record({
      currentName: fc.constantFrom('Sheet1', 'Revenue', 'Inputs'),
      nextName: fc.constantFrom('Next', 'Forecast', 'Archive'),
    })
    .filter(({ currentName, nextName }) => currentName !== nextName)
    .map(({ currentName, nextName }) => ({ kind: 'renameSheet' as const, currentName, nextName })),
  fc.record({ name: fc.constantFrom('Sheet1', 'Revenue', 'Inputs') }).map(({ name }) => ({ kind: 'deleteSheet' as const, name })),
  axisSpanCommandArbitrary('insertRows'),
  axisSpanCommandArbitrary('deleteRows'),
  axisSpanCommandArbitrary('insertColumns'),
  axisSpanCommandArbitrary('deleteColumns'),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Revenue', 'Inputs'),
      rows: fc.integer({ min: 0, max: 10 }),
      cols: fc.integer({ min: 0, max: 10 }),
    })
    .map((command) => ({
      kind: 'setFreezePane' as const,
      sheetName: command.sheetName,
      rows: command.rows,
      cols: command.cols,
    })),
  fuzzCellRangeRefArbitrary.map((range) => ({ kind: 'setFilter' as const, range })),
  fuzzCellRangeRefArbitrary.map((range) => ({ kind: 'clearFilter' as const, range })),
  fuzzCellRangeRefArbitrary.map((range) => ({
    kind: 'setSort' as const,
    range,
    keys: [{ keyAddress: range.startAddress, direction: 'asc' as const }],
  })),
  fuzzCellRangeRefArbitrary.map((range) => ({ kind: 'clearSort' as const, range })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Revenue', 'Inputs'),
      startRow: fc.integer({ min: 0, max: 24 }),
      count: fc.integer({ min: 1, max: 8 }),
      height: fc.option(fc.double({ noNaN: true, noDefaultInfinity: true, min: 1, max: 200 }), { nil: null }),
      hidden: fc.option(fc.boolean(), { nil: undefined }),
    })
    .map((command) => createUpdateRowMetadataCommand(command)),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Revenue', 'Inputs'),
      startCol: fc.integer({ min: 0, max: 24 }),
      count: fc.integer({ min: 1, max: 8 }),
      width: fc.option(fc.double({ noNaN: true, noDefaultInfinity: true, min: 1, max: 300 }), { nil: null }),
      hidden: fc.option(fc.boolean(), { nil: undefined }),
    })
    .map((command) => createUpdateColumnMetadataCommand(command)),
)

function axisSpanCommandArbitrary<TKind extends 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns'>(
  kind: TKind,
): fc.Arbitrary<WorkbookAgentStructuralCommand> {
  return fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Revenue', 'Inputs'),
      start: fc.integer({ min: 0, max: 24 }),
      count: fc.integer({ min: 1, max: 8 }),
    })
    .map((command) => createAxisSpanCommand(kind, command))
}

function createAxisSpanCommand(
  kind: 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns',
  command: {
    readonly sheetName: string
    readonly start: number
    readonly count: number
  },
): WorkbookAgentStructuralCommand {
  switch (kind) {
    case 'insertRows':
      return { kind, sheetName: command.sheetName, start: command.start, count: command.count }
    case 'deleteRows':
      return { kind, sheetName: command.sheetName, start: command.start, count: command.count }
    case 'insertColumns':
      return { kind, sheetName: command.sheetName, start: command.start, count: command.count }
    case 'deleteColumns':
      return { kind, sheetName: command.sheetName, start: command.start, count: command.count }
  }
}

function createUpdateRowMetadataCommand(command: {
  readonly sheetName: string
  readonly startRow: number
  readonly count: number
  readonly height: number | null
  readonly hidden?: boolean | undefined
}): WorkbookAgentStructuralCommand {
  const result: Extract<WorkbookAgentStructuralCommand, { kind: 'updateRowMetadata' }> = {
    kind: 'updateRowMetadata',
    sheetName: command.sheetName,
    startRow: command.startRow,
    count: command.count,
    height: command.height,
  }
  if (command.hidden !== undefined) {
    result.hidden = command.hidden
  }
  return result
}

function createUpdateColumnMetadataCommand(command: {
  readonly sheetName: string
  readonly startCol: number
  readonly count: number
  readonly width: number | null
  readonly hidden?: boolean | undefined
}): WorkbookAgentStructuralCommand {
  const result: Extract<WorkbookAgentStructuralCommand, { kind: 'updateColumnMetadata' }> = {
    kind: 'updateColumnMetadata',
    sheetName: command.sheetName,
    startCol: command.startCol,
    count: command.count,
    width: command.width,
  }
  if (command.hidden !== undefined) {
    result.hidden = command.hidden
  }
  return result
}

function corruptCommand(command: WorkbookAgentStructuralCommand): unknown {
  switch (command.kind) {
    case 'createSheet':
    case 'deleteSheet':
      return { ...command, name: '' }
    case 'renameSheet':
      return { ...command, nextName: '' }
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns':
      return { ...command, count: 0 }
    case 'setFreezePane':
      return { ...command, rows: -1 }
    case 'setFilter':
    case 'clearFilter':
    case 'setSort':
    case 'clearSort':
      return { ...command, range: { ...command.range, sheetName: null } }
    case 'updateRowMetadata':
      return { ...command, count: 0 }
    case 'updateColumnMetadata':
      return { ...command, count: 0 }
  }
}
