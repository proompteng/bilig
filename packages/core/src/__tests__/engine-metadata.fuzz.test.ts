import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import type { AsyncCommand } from 'fast-check'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { runModelProperty } from '@bilig/test-fuzz'
import { createEngineSeedSnapshot, normalizeSnapshotForSemanticComparison } from './engine-fuzz-helpers.js'
import {
  applyEngineSemanticActionAndCaptureResult,
  exportEngineSemanticReplaySnapshot,
  metadataSeedNames,
  metadataSemanticActionArbitrary,
  projectMetadataSnapshot,
  type EngineSemanticAction,
} from './engine-fuzz-metadata-helpers.js'

interface EngineMetadataModel {
  initialSnapshot: WorkbookSnapshot
  applied: EngineSemanticAction[]
  undone: EngineSemanticAction[]
  history: string[]
}

async function applyAndExpectMetadataSemantics(engine: SpreadsheetEngine, model: EngineMetadataModel): Promise<void> {
  const snapshot = engine.exportSnapshot()
  try {
    const expectedSnapshot = await exportEngineSemanticReplaySnapshot(model.initialSnapshot, model.applied)
    expect(normalizeSnapshotForSemanticComparison(snapshot)).toEqual(normalizeSnapshotForSemanticComparison(expectedSnapshot))
    expect(projectMetadataSnapshot(snapshot)).toEqual(projectMetadataSnapshot(expectedSnapshot))
  } catch (error) {
    const wrapped =
      error instanceof Error
        ? new Error(`${error.message}\ncommandHistory=${model.history.join(' -> ')}`)
        : new Error(`commandHistory=${model.history.join(' -> ')}`)
    if (error instanceof Error) {
      wrapped.stack = error.stack
    }
    throw wrapped
  }
}

function applySemanticCommandArbitrary(
  actionArbitrary: fc.Arbitrary<EngineSemanticAction>,
): fc.Arbitrary<AsyncCommand<EngineMetadataModel, SpreadsheetEngine>> {
  return actionArbitrary.map((action) => ({
    check: () => true,
    run: async (model, real) => {
      const result = applyEngineSemanticActionAndCaptureResult(real, action)
      if (result.accepted) {
        model.applied.push(action)
        model.undone = []
      }
      model.history.push(`apply(${JSON.stringify(action)})`)
      await applyAndExpectMetadataSemantics(real, model)
    },
    toString: () => `apply(${JSON.stringify(action)})`,
  }))
}

const undoCommandArbitrary: fc.Arbitrary<AsyncCommand<EngineMetadataModel, SpreadsheetEngine>> = fc.constant({
  check: (model) => model.applied.length > 0,
  run: async (model, real) => {
    if (!real.undo()) {
      throw new Error(`Undo command failed despite a non-empty metadata history\ncommandHistory=${model.history.join(' -> ')}`)
    }
    const action = model.applied.pop()
    if (!action) {
      throw new Error('Undo ran with no applied metadata action')
    }
    model.undone.push(action)
    model.history.push('undo()')
    await applyAndExpectMetadataSemantics(real, model)
  },
  toString: () => 'undo()',
})

const redoCommandArbitrary: fc.Arbitrary<AsyncCommand<EngineMetadataModel, SpreadsheetEngine>> = fc.constant({
  check: (model) => model.undone.length > 0,
  run: async (model, real) => {
    if (!real.redo()) {
      throw new Error(`Redo command failed despite a non-empty metadata undo stack\ncommandHistory=${model.history.join(' -> ')}`)
    }
    const action = model.undone.pop()
    if (!action) {
      throw new Error('Redo ran with no undone metadata action')
    }
    model.applied.push(action)
    model.history.push('redo()')
    await applyAndExpectMetadataSemantics(real, model)
  },
  toString: () => 'redo()',
})

describe('engine metadata fuzz', () => {
  for (const seedName of metadataSeedNames) {
    it(`keeps metadata semantics aligned for ${seedName} seeded workbooks`, async () => {
      const seedSnapshot = await createEngineSeedSnapshot(seedName, `fuzz-core-metadata-${seedName}`)
      const ran = await runModelProperty({
        suite: `core/metadata/${seedName}`,
        commands: (replayPath) =>
          fc.commands([applySemanticCommandArbitrary(metadataSemanticActionArbitrary), undoCommandArbitrary, redoCommandArbitrary], {
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
            replicaId: `metadata-${seedName}`,
          })
          await engine.ready()
          engine.importSnapshot(structuredClone(seedSnapshot))
          await applyAndExpectMetadataSemantics(engine, {
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
