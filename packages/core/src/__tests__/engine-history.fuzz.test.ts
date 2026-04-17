import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import type { AsyncCommand } from 'fast-check'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { runModelProperty } from '@bilig/test-fuzz'
import {
  applyActionAndCaptureResult,
  assertSnapshotInvariants,
  clearActionArbitrary,
  copyActionArbitrary,
  createEngineSeedSnapshot,
  deleteColumnsActionArbitrary,
  deleteRowsActionArbitrary,
  engineSeedNames,
  exportReplaySnapshot,
  fillActionArbitrary,
  formatActionArbitrary,
  formulaActionArbitrary,
  insertColumnsActionArbitrary,
  insertRowsActionArbitrary,
  moveActionArbitrary,
  normalizeSnapshotForSemanticComparison,
  styleActionArbitrary,
  type CoreAction,
  valuesActionArbitrary,
} from './engine-fuzz-helpers.js'

interface EngineHistoryModel {
  initialSnapshot: WorkbookSnapshot
  applied: CoreAction[]
  undone: CoreAction[]
  history: string[]
}

async function applyAndExpectSnapshot(engine: SpreadsheetEngine, model: EngineHistoryModel): Promise<void> {
  const snapshot = engine.exportSnapshot()
  assertSnapshotInvariants(snapshot)
  try {
    const expectedSnapshot = await exportReplaySnapshot(model.initialSnapshot, model.applied)
    expect(normalizeSnapshotForSemanticComparison(snapshot)).toEqual(normalizeSnapshotForSemanticComparison(expectedSnapshot))
  } catch (error) {
    const commandHistory = model.history.join(' -> ')
    const wrapped =
      error instanceof Error
        ? new Error(`${error.message}\ncommandHistory=${commandHistory}`)
        : new Error(`commandHistory=${commandHistory}`)
    if (error instanceof Error) {
      wrapped.stack = error.stack
    }
    throw wrapped
  }
}

function applyActionCommandArbitrary(
  actionArbitrary: fc.Arbitrary<CoreAction>,
): fc.Arbitrary<AsyncCommand<EngineHistoryModel, SpreadsheetEngine>> {
  return actionArbitrary.map((action) => ({
    check: () => true,
    run: async (model, real) => {
      const result = applyActionAndCaptureResult(real, action)
      if (result.accepted) {
        model.applied.push(action)
        model.undone = []
      }
      model.history.push(`apply(${JSON.stringify(action)})`)
      await applyAndExpectSnapshot(real, model)
    },
    toString: () => `apply(${JSON.stringify(action)})`,
  }))
}

const undoCommandArbitrary: fc.Arbitrary<AsyncCommand<EngineHistoryModel, SpreadsheetEngine>> = fc.constant({
  check: (model) => model.applied.length > 0,
  run: async (model, real) => {
    if (!real.undo()) {
      throw new Error('Undo command failed despite a non-empty applied history')
    }
    const action = model.applied.pop()
    if (!action) {
      throw new Error('Undo command ran with no applied action')
    }
    model.undone.push(action)
    model.history.push('undo()')
    await applyAndExpectSnapshot(real, model)
  },
  toString: () => 'undo()',
})

const redoCommandArbitrary: fc.Arbitrary<AsyncCommand<EngineHistoryModel, SpreadsheetEngine>> = fc.constant({
  check: (model) => model.undone.length > 0,
  run: async (model, real) => {
    if (!real.redo()) {
      throw new Error('Redo command failed despite a non-empty undone history')
    }
    const action = model.undone.pop()
    if (!action) {
      throw new Error('Redo command ran with no undone action')
    }
    model.applied.push(action)
    model.history.push('redo()')
    await applyAndExpectSnapshot(real, model)
  },
  toString: () => 'redo()',
})

const engineHistoryCommandArbitraries: Array<fc.Arbitrary<AsyncCommand<EngineHistoryModel, SpreadsheetEngine>>> = [
  applyActionCommandArbitrary(valuesActionArbitrary),
  applyActionCommandArbitrary(formulaActionArbitrary),
  applyActionCommandArbitrary(styleActionArbitrary),
  applyActionCommandArbitrary(formatActionArbitrary),
  applyActionCommandArbitrary(clearActionArbitrary),
  applyActionCommandArbitrary(fillActionArbitrary),
  applyActionCommandArbitrary(copyActionArbitrary),
  applyActionCommandArbitrary(moveActionArbitrary),
  applyActionCommandArbitrary(insertRowsActionArbitrary),
  applyActionCommandArbitrary(deleteRowsActionArbitrary),
  applyActionCommandArbitrary(insertColumnsActionArbitrary),
  applyActionCommandArbitrary(deleteColumnsActionArbitrary),
  undoCommandArbitrary,
  redoCommandArbitrary,
]

describe('engine history fuzz', () => {
  for (const seedName of engineSeedNames) {
    it(`keeps model-based history semantics aligned for ${seedName} seeded workbooks`, async () => {
      const seedSnapshot = await createEngineSeedSnapshot(seedName, `fuzz-core-history-${seedName}`)
      const ran = await runModelProperty({
        suite: `core/history/${seedName}`,
        commands: (replayPath) =>
          fc.commands(engineHistoryCommandArbitraries, {
            maxCommands: 18,
            ...(replayPath ? { replayPath } : {}),
          }),
        createModel: () => ({
          initialSnapshot: structuredClone(seedSnapshot),
          applied: [],
          undone: [],
          history: [],
        }),
        createReal: async () => {
          const engine = new SpreadsheetEngine({
            workbookName: seedSnapshot.workbook.name,
            replicaId: `fuzz-core-model-${seedName}`,
          })
          await engine.ready()
          engine.importSnapshot(structuredClone(seedSnapshot))
          await applyAndExpectSnapshot(engine, {
            initialSnapshot: structuredClone(seedSnapshot),
            applied: [],
            undone: [],
            history: [],
          })
          return engine
        },
      })
      expect(ran).toBe(true)
    })
  }
})
