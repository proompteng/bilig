import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { runProperty } from '@bilig/test-fuzz'
import {
  applyActionAndCaptureResult,
  applyCoreAction,
  assertSnapshotInvariants,
  corePreparationActionArbitrary,
  coreStructuralActionArbitrary,
  createEngineSeedSnapshot,
  engineSeedNameArbitrary,
  normalizeSnapshotForSemanticComparison,
} from './engine-fuzz-helpers.js'

function undoAll(engine: SpreadsheetEngine): number {
  let count = 0
  while (engine.undo()) {
    count += 1
  }
  return count
}

function expectSemanticSnapshot(actual: WorkbookSnapshot, expected: WorkbookSnapshot): void {
  expect(normalizeSnapshotForSemanticComparison(actual)).toEqual(normalizeSnapshotForSemanticComparison(expected))
}

describe('engine structural fuzz', () => {
  it('preserves structural inverse semantics from seeded workbook states', async () => {
    await runProperty({
      suite: 'core/structure/inverse-replay',
      arbitrary: fc.record({
        seedName: engineSeedNameArbitrary,
        setupActions: fc.array(corePreparationActionArbitrary, { minLength: 0, maxLength: 3 }),
        structuralActions: fc.array(coreStructuralActionArbitrary, { minLength: 1, maxLength: 4 }),
      }),
      predicate: async ({ seedName, setupActions, structuralActions }) => {
        const seedSnapshot = await createEngineSeedSnapshot(seedName, `fuzz-core-structure-${seedName}`)
        const engine = new SpreadsheetEngine({
          workbookName: seedSnapshot.workbook.name,
          replicaId: `fuzz-core-structure-${seedName}`,
        })
        const replay = new SpreadsheetEngine({
          workbookName: seedSnapshot.workbook.name,
          replicaId: `fuzz-core-structure-replay-${seedName}`,
        })
        await engine.ready()
        await replay.ready()
        engine.importSnapshot(structuredClone(seedSnapshot))
        replay.importSnapshot(structuredClone(seedSnapshot))

        let acceptedCount = 0
        const restoreChecks: Array<Promise<void>> = []
        const allActions = [...setupActions, ...structuralActions]
        for (const action of allActions) {
          const result = applyActionAndCaptureResult(engine, action)
          if (result.accepted) {
            acceptedCount += 1
            applyCoreAction(replay, action)
          }
          assertSnapshotInvariants(result.after)
          expectSemanticSnapshot(result.after, replay.exportSnapshot())

          restoreChecks.push(
            (async () => {
              const restored = new SpreadsheetEngine({
                workbookName: result.after.workbook.name,
                replicaId: `fuzz-core-structure-restore-${seedName}`,
              })
              await restored.ready()
              restored.importSnapshot(result.after)
              expect(restored.exportSnapshot()).toEqual(result.after)
            })(),
          )
        }

        await Promise.all(restoreChecks)
        expect(undoAll(engine)).toBe(acceptedCount)
        expectSemanticSnapshot(engine.exportSnapshot(), seedSnapshot)
      },
    })
  })
})
