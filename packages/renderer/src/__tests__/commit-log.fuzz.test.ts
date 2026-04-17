import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import type { CommitOp } from '@bilig/core'
import { runProperty } from '@bilig/test-fuzz'
import type { CellDescriptor, SheetDescriptor, WorkbookDescriptor } from '../descriptors.js'
import { normalizeCommitOps } from '../commit-log.js'
import { validateDescriptorTree } from '../validation.js'

type CommitSubset = Extract<
  CommitOp,
  { kind: 'upsertWorkbook' } | { kind: 'upsertSheet' } | { kind: 'deleteSheet' } | { kind: 'upsertCell' } | { kind: 'deleteCell' }
>

describe('renderer commit log fuzz', () => {
  it('should keep commit normalization idempotent and last-write-wins by semantic key', async () => {
    await runProperty({
      suite: 'renderer/commit-log/idempotent-normalization',
      arbitrary: fc.array(commitOpArbitrary, { minLength: 1, maxLength: 24 }),
      predicate: async (ops) => {
        const normalized = normalizeCommitOps(ops)
        expect(normalizeCommitOps(normalized)).toEqual(normalized)
        expect(new Set(normalized.map(commitKey)).size).toBe(normalized.length)
        expect(normalized).toEqual(expectedNormalizedOps(ops))
      },
    })
  })

  it('should accept generated valid workbook descriptor trees', async () => {
    await runProperty({
      suite: 'renderer/validation/generated-workbooks',
      arbitrary: workbookSpecArbitrary,
      predicate: async (spec) => {
        expect(() => validateDescriptorTree(createWorkbookDescriptor(spec))).not.toThrow()
      },
    })
  })
})

// Helpers

type WorkbookSpec = ReadonlyArray<{
  name: string
  cells: ReadonlyArray<{ addr: string; value?: number | string | boolean; formula?: string }>
}>

const commitOpArbitrary = fc.oneof<CommitSubset>(
  fc.constantFrom('book-a', 'book-b').map((name) => ({ kind: 'upsertWorkbook' as const, name })),
  fc
    .record({
      name: fc.constantFrom('Sheet1', 'Sheet2', 'Sheet3'),
      order: fc.integer({ min: 0, max: 3 }),
    })
    .map(({ name, order }) => ({ kind: 'upsertSheet' as const, name, order })),
  fc.constantFrom('Sheet1', 'Sheet2', 'Sheet3').map((name) => ({ kind: 'deleteSheet' as const, name })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      addr: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
      value: fc.oneof(fc.integer({ min: -20, max: 20 }), fc.boolean(), fc.constantFrom('north', 'south')),
    })
    .map(({ sheetName, addr, value }) => ({ kind: 'upsertCell' as const, sheetName, addr, value })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      addr: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
      formula: fc.constantFrom('A1+1', 'B2*2', '1+2'),
    })
    .map(({ sheetName, addr, formula }) => ({ kind: 'upsertCell' as const, sheetName, addr, formula })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      addr: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
    })
    .map(({ sheetName, addr }) => ({ kind: 'deleteCell' as const, sheetName, addr })),
)

const workbookSpecArbitrary = fc.uniqueArray(
  fc.record({
    name: fc.constantFrom('Sheet1', 'Sheet2', 'Sheet3'),
    cells: fc.uniqueArray(
      fc.oneof(
        fc.record({
          addr: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
          value: fc.oneof(fc.integer({ min: -20, max: 20 }), fc.boolean(), fc.constantFrom('north', 'south')),
        }),
        fc.record({
          addr: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
          formula: fc.constantFrom('A1+1', 'B2*2', '1+2'),
        }),
      ),
      {
        minLength: 0,
        maxLength: 4,
        selector: (cell) => cell.addr,
      },
    ),
  }),
  {
    minLength: 1,
    maxLength: 3,
    selector: (sheet) => sheet.name,
  },
)

function commitKey(op: CommitSubset): string {
  switch (op.kind) {
    case 'upsertWorkbook':
      return 'workbook'
    case 'renameSheet':
      return `sheet:${op.newName}`
    case 'upsertSheet':
    case 'deleteSheet':
      return `sheet:${op.name}`
    case 'upsertCell':
    case 'deleteCell':
      return `cell:${op.sheetName}!${op.addr}`
  }
}

function expectedNormalizedOps(ops: readonly CommitSubset[]): CommitSubset[] {
  const orderedKeys: string[] = []
  const latest = new Map<string, CommitSubset>()
  ops.forEach((op) => {
    const key = commitKey(op)
    if (!latest.has(key)) {
      orderedKeys.push(key)
    }
    latest.set(key, op)
  })
  return orderedKeys.flatMap((key) => {
    const op = latest.get(key)
    return op ? [op] : []
  })
}

function createWorkbookDescriptor(spec: WorkbookSpec): WorkbookDescriptor {
  const workbook: WorkbookDescriptor = {
    kind: 'Workbook',
    props: { name: 'fuzz-book' },
    children: [],
    parent: null,
    container: null,
  }
  workbook.children = spec.map((sheetSpec) => {
    const sheet: SheetDescriptor = {
      kind: 'Sheet',
      props: { name: sheetSpec.name },
      children: [],
      parent: workbook,
      container: null,
    }
    sheet.children = sheetSpec.cells.map((cellSpec) => {
      const cell: CellDescriptor = {
        kind: 'Cell',
        props: {
          addr: cellSpec.addr,
          ...(cellSpec.formula !== undefined ? { formula: cellSpec.formula } : {}),
          ...(cellSpec.value !== undefined ? { value: cellSpec.value } : {}),
        },
        parent: sheet,
        container: null,
      }
      return cell
    })
    return sheet
  })
  return workbook
}
