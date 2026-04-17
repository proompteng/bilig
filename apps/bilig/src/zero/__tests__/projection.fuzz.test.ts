import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { SpreadsheetEngine } from '@bilig/core'
import { runProperty } from '@bilig/test-fuzz'
import {
  applyProjectionAction,
  projectionActionArbitrary,
  projectProjectionFromEngine,
  projectProjectionFromSnapshot,
} from './projection-fuzz-helpers.js'

describe('projection fuzz', () => {
  it('preserves projection parity between live engine state and exported snapshots', async () => {
    await runProperty({
      suite: 'bilig/projection/live-snapshot-parity',
      arbitrary: fc.array(projectionActionArbitrary, { minLength: 4, maxLength: 18 }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({
          workbookName: 'projection-fuzz',
          replicaId: 'projection-fuzz',
        })
        await engine.ready()

        actions.forEach((action) => {
          applyProjectionAction(engine, action)
        })

        expect(projectProjectionFromEngine(engine)).toEqual(projectProjectionFromSnapshot(engine))
      },
    })
  })
})
