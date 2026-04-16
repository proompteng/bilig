import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import { SpreadsheetEngine } from '@bilig/core'
import type { LiteralInput } from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import { buildWorkbookSourceProjection, buildWorkbookSourceProjectionFromEngine, type WorkbookProjectionOptions } from '../projection.js'

type MigrationParityAction =
  | { kind: 'value'; address: string; value: LiteralInput }
  | { kind: 'formula'; address: string; formula: string }
  | { kind: 'insertRows'; start: number; count: number }
  | { kind: 'deleteColumns'; start: number; count: number }

const options: WorkbookProjectionOptions = {
  revision: 5,
  calculatedRevision: 5,
  ownerUserId: 'owner-1',
  updatedBy: 'user-1',
  updatedAt: '2026-04-13T08:15:00.000Z',
}

describe('migration parity fuzz', () => {
  it('preserves source projection meaning when a snapshot is re-imported into a fresh engine', async () => {
    await runProperty({
      suite: 'bilig/migration/snapshot-import-parity',
      arbitrary: fc.array(migrationParityActionArbitrary, { minLength: 4, maxLength: 18 }),
      predicate: async (actions) => {
        const source = new SpreadsheetEngine({
          workbookName: 'migration-fuzz-source',
          replicaId: 'migration-fuzz-source',
        })
        await source.ready()

        actions.forEach((action) => applyMigrationAction(source, action))

        const snapshot = source.exportSnapshot()
        const restored = new SpreadsheetEngine({
          workbookName: 'migration-fuzz-restored',
          replicaId: 'migration-fuzz-restored',
        })
        await restored.ready()
        restored.importSnapshot(snapshot)

        expect(buildWorkbookSourceProjectionFromEngine('doc-1', restored, options)).toEqual(
          buildWorkbookSourceProjection('doc-1', snapshot, options),
        )
      },
    })
  })
})

// Helpers

const migrationParityActionArbitrary = fc.oneof<MigrationParityAction>(
  fc
    .record({
      address: fc.constantFrom('A1', 'B1', 'A2', 'B2', 'C3'),
      value: fc.oneof<LiteralInput>(fc.integer({ min: -20, max: 20 }), fc.boolean(), fc.constantFrom('draft', 'final'), fc.constant(null)),
    })
    .map((action) => Object.assign({ kind: 'value' as const }, action)),
  fc
    .record({
      address: fc.constantFrom('A1', 'B1', 'A2', 'B2', 'C3'),
      formula: fc.constantFrom('1', 'A1+1', 'SUM(A1:B2)'),
    })
    .map((action) => Object.assign({ kind: 'formula' as const }, action)),
  fc
    .record({ start: fc.integer({ min: 0, max: 2 }), count: fc.integer({ min: 1, max: 1 }) })
    .map((action) => Object.assign({ kind: 'insertRows' as const }, action)),
  fc
    .record({ start: fc.integer({ min: 0, max: 2 }), count: fc.integer({ min: 1, max: 1 }) })
    .map((action) => Object.assign({ kind: 'deleteColumns' as const }, action)),
)

function applyMigrationAction(engine: SpreadsheetEngine, action: MigrationParityAction): void {
  switch (action.kind) {
    case 'value':
      engine.setCellValue('Sheet1', action.address, action.value)
      return
    case 'formula':
      engine.setCellFormula('Sheet1', action.address, action.formula)
      return
    case 'insertRows':
      engine.insertRows('Sheet1', action.start, action.count)
      return
    case 'deleteColumns':
      engine.deleteColumns('Sheet1', action.start, action.count)
      return
  }
}
