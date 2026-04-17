import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { SpreadsheetEngine } from '../engine.js'
import { runProperty } from '@bilig/test-fuzz'
import {
  applyActionAndCaptureResult,
  applyCoreAction,
  assertSnapshotInvariants,
  coreReplicaActionArbitrary,
  createEngineSeedSnapshot,
  engineSeedNameArbitrary,
  normalizeSnapshotForSemanticComparison,
} from './engine-fuzz-helpers.js'

function expectSemanticSnapshot(
  actual: import('@bilig/protocol').WorkbookSnapshot,
  expected: import('@bilig/protocol').WorkbookSnapshot,
): void {
  expect(normalizeSnapshotForSemanticComparison(actual)).toEqual(normalizeSnapshotForSemanticComparison(expected))
}

describe('engine replica fuzz', () => {
  it('keeps local batch replay, snapshot restore, and semantic replay aligned', async () => {
    await runProperty({
      suite: 'core/replica/local-batch-replay-parity',
      arbitrary: fc.record({
        seedName: engineSeedNameArbitrary,
        actions: fc.array(coreReplicaActionArbitrary, { minLength: 4, maxLength: 18 }),
      }),
      predicate: async ({ seedName, actions }) => {
        const seedSnapshot = await createEngineSeedSnapshot(seedName, `fuzz-core-replica-${seedName}`)
        const primary = new SpreadsheetEngine({
          workbookName: seedSnapshot.workbook.name,
          replicaId: `primary-${seedName}`,
        })
        const replica = new SpreadsheetEngine({
          workbookName: seedSnapshot.workbook.name,
          replicaId: `replica-${seedName}`,
        })
        await Promise.all([primary.ready(), replica.ready()])

        const outbound: EngineOpBatch[] = []
        primary.subscribeBatches((batch) => outbound.push(batch))

        primary.importSnapshot(structuredClone(seedSnapshot))
        replica.importSnapshot(structuredClone(seedSnapshot))

        const replay = new SpreadsheetEngine({
          workbookName: seedSnapshot.workbook.name,
          replicaId: `semantic-replay-${seedName}`,
        })
        const restored = new SpreadsheetEngine({
          workbookName: seedSnapshot.workbook.name,
          replicaId: `restore-${seedName}`,
        })
        let appliedBatches = 0
        await Promise.all([replay.ready(), restored.ready()])
        expect(replica.exportSnapshot()).toEqual(primary.exportSnapshot())
        replay.importSnapshot(structuredClone(seedSnapshot))

        for (const action of actions) {
          const result = applyActionAndCaptureResult(primary, action)
          if (result.accepted) {
            applyCoreAction(replay, action)
          }
          while (appliedBatches < outbound.length) {
            const nextBatch = outbound[appliedBatches]
            if (!nextBatch) {
              throw new Error(`Missing outbound batch at index ${appliedBatches}`)
            }
            replica.applyRemoteBatch(nextBatch)
            appliedBatches += 1
          }

          const primarySnapshot = result.after
          assertSnapshotInvariants(primarySnapshot)
          expect(replica.exportSnapshot()).toEqual(primarySnapshot)
          expectSemanticSnapshot(primarySnapshot, replay.exportSnapshot())
          restored.importSnapshot(primarySnapshot)
          expect(restored.exportSnapshot()).toEqual(primarySnapshot)
        }
      },
    })
  })
})
