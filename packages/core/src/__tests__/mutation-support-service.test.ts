/* eslint-disable typescript-eslint/no-unsafe-type-assertion -- support-service error-path tests intentionally inject partial collaborators */
import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { CellFlags } from '../cell-store.js'
import { SpreadsheetEngine } from '../engine.js'
import { createEngineMutationSupportService, type EngineMutationSupportService } from '../engine/services/mutation-support-service.js'

interface MutationSupportScratch {
  changedInputEpoch: number
  changedFormulaEpoch: number
  changedUnionEpoch: number
  explicitChangedEpoch: number
  impactedFormulaEpoch: number
  materializedCellCount: number
  changedInputSeen: Uint32Array
  changedInputBuffer: Uint32Array
  changedFormulaSeen: Uint32Array
  changedFormulaBuffer: Uint32Array
  changedUnionSeen: Uint32Array
  changedUnion: Uint32Array
  mutationRoots: Uint32Array
  materializedCells: Uint32Array
  explicitChangedSeen: Uint32Array
  explicitChangedBuffer: Uint32Array
  impactedFormulaSeen: Uint32Array
  impactedFormulaBuffer: Uint32Array
}

type StubMutationSupportService = EngineMutationSupportService & {
  readonly __scratch: MutationSupportScratch
}

function createStubMutationSupportService(
  overrides: Partial<Parameters<typeof createEngineMutationSupportService>[0]> = {},
): StubMutationSupportService {
  const scratch: MutationSupportScratch = {
    changedInputEpoch: 1,
    changedFormulaEpoch: 1,
    changedUnionEpoch: 1,
    explicitChangedEpoch: 1,
    impactedFormulaEpoch: 1,
    materializedCellCount: 0,
    changedInputSeen: new Uint32Array(16),
    changedInputBuffer: new Uint32Array(16),
    changedFormulaSeen: new Uint32Array(16),
    changedFormulaBuffer: new Uint32Array(16),
    changedUnionSeen: new Uint32Array(16),
    changedUnion: new Uint32Array(16),
    mutationRoots: new Uint32Array(16),
    materializedCells: new Uint32Array(4),
    explicitChangedSeen: new Uint32Array(16),
    explicitChangedBuffer: new Uint32Array(16),
    impactedFormulaSeen: new Uint32Array(16),
    impactedFormulaBuffer: new Uint32Array(16),
  }

  const defaults: Parameters<typeof createEngineMutationSupportService>[0] = {
    state: {
      workbook: {
        cellStore: {
          size: 0,
          sheetIds: [],
          rows: [],
          cols: [],
          flags: [],
          getValue: () => ({ tag: ValueTag.Empty }),
          setValue: () => undefined,
        },
        ensureCellRecord: () => ({ created: false, cellIndex: 0 }),
        ensureCellAt: () => ({ created: false, cellIndex: 0 }),
        getSheetNameById: () => 'Sheet1',
        getAddress: () => 'A1',
        getSpill: () => undefined,
        getSheet: () => undefined,
        deleteSheet: () => undefined,
        sheetsByName: new Map(),
      },
      strings: { get: () => '' },
      formulas: new Map(),
      ranges: { addDynamicMember: () => [] },
    } as never,
    edgeArena: {
      empty: () => ({ ptr: -1, len: 0 }),
      appendUnique: () => ({ ptr: -1, len: 0 }),
      readView: () => new Uint32Array(),
    } as never,
    reverseState: {
      reverseCellEdges: [],
      reverseRangeEdges: [],
    },
    removeFormula: () => false,
    rebindFormulasForSheet: () => 0,
    getSelectionState: () =>
      ({
        sheetName: 'Sheet1',
        address: 'A1',
        anchorAddress: 'A1',
        range: { startAddress: 'A1', endAddress: 'A1' },
        editMode: 'idle',
      }) as never,
    setSelection: () => undefined,
    applyDerivedOp: () => [],
    scheduleWasmProgramSync: () => undefined,
    ensureRecalcScratchCapacity: () => undefined,
    collectFormulaDependents: () => new Uint32Array(),
    getChangedInputEpoch: () => scratch.changedInputEpoch,
    setChangedInputEpoch: (next) => {
      scratch.changedInputEpoch = next
    },
    getChangedInputSeen: () => scratch.changedInputSeen,
    setChangedInputSeen: (next) => {
      scratch.changedInputSeen = next
    },
    getChangedInputBuffer: () => scratch.changedInputBuffer,
    setChangedInputBuffer: (next) => {
      scratch.changedInputBuffer = next
    },
    getChangedFormulaEpoch: () => scratch.changedFormulaEpoch,
    setChangedFormulaEpoch: (next) => {
      scratch.changedFormulaEpoch = next
    },
    getChangedFormulaSeen: () => scratch.changedFormulaSeen,
    setChangedFormulaSeen: (next) => {
      scratch.changedFormulaSeen = next
    },
    getChangedFormulaBuffer: () => scratch.changedFormulaBuffer,
    setChangedFormulaBuffer: (next) => {
      scratch.changedFormulaBuffer = next
    },
    getChangedUnionEpoch: () => scratch.changedUnionEpoch,
    setChangedUnionEpoch: (next) => {
      scratch.changedUnionEpoch = next
    },
    getChangedUnionSeen: () => scratch.changedUnionSeen,
    setChangedUnionSeen: (next) => {
      scratch.changedUnionSeen = next
    },
    getChangedUnion: () => scratch.changedUnion,
    setChangedUnion: (next) => {
      scratch.changedUnion = next
    },
    getMutationRoots: () => scratch.mutationRoots,
    setMutationRoots: (next) => {
      scratch.mutationRoots = next
    },
    getMaterializedCellCount: () => scratch.materializedCellCount,
    setMaterializedCellCount: (next) => {
      scratch.materializedCellCount = next
    },
    getMaterializedCells: () => scratch.materializedCells,
    setMaterializedCells: (next) => {
      scratch.materializedCells = next
    },
    getExplicitChangedEpoch: () => scratch.explicitChangedEpoch,
    setExplicitChangedEpoch: (next) => {
      scratch.explicitChangedEpoch = next
    },
    getExplicitChangedSeen: () => scratch.explicitChangedSeen,
    setExplicitChangedSeen: (next) => {
      scratch.explicitChangedSeen = next
    },
    getExplicitChangedBuffer: () => scratch.explicitChangedBuffer,
    setExplicitChangedBuffer: (next) => {
      scratch.explicitChangedBuffer = next
    },
    getImpactedFormulaEpoch: () => scratch.impactedFormulaEpoch,
    setImpactedFormulaEpoch: (next) => {
      scratch.impactedFormulaEpoch = next
    },
    getImpactedFormulaSeen: () => scratch.impactedFormulaSeen,
    setImpactedFormulaSeen: (next) => {
      scratch.impactedFormulaSeen = next
    },
    getImpactedFormulaBuffer: () => scratch.impactedFormulaBuffer,
    setImpactedFormulaBuffer: (next) => {
      scratch.impactedFormulaBuffer = next
    },
  }

  return Object.assign(createEngineMutationSupportService({ ...defaults, ...overrides }), {
    __scratch: scratch,
  })
}

function isEngineMutationSupportService(value: unknown): value is EngineMutationSupportService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'materializeSpill') === 'function' &&
    typeof Reflect.get(value, 'clearOwnedSpill') === 'function' &&
    typeof Reflect.get(value, 'removeSheetRuntime') === 'function'
  )
}

function getMutationSupportService(engine: SpreadsheetEngine): EngineMutationSupportService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const support = Reflect.get(runtime, 'support')
  if (!isEngineMutationSupportService(support)) {
    throw new TypeError('Expected engine mutation support service')
  }
  return support
}

describe('EngineMutationSupportService', () => {
  it('tracks changed roots and unions through the public wrapper methods', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'support-wrapper-roots' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const support = getMutationSupportService(engine)
    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(a1Index).toBeDefined()
    expect(b1Index).toBeDefined()

    Effect.runSync(support.beginMutationCollection())
    const changedInputCount = Effect.runSync(support.markInputChanged(a1Index!, 0))
    const changedFormulaCount = Effect.runSync(support.markFormulaChanged(b1Index!, 0))
    const explicitChangedCount = Effect.runSync(support.markExplicitChanged(a1Index!, 0))

    expect(changedInputCount).toBe(1)
    expect(changedFormulaCount).toBe(1)
    expect(explicitChangedCount).toBe(1)
    expect(Effect.runSync(support.getChangedInputBuffer())[0]).toBe(a1Index)

    const roots = Effect.runSync(support.composeMutationRoots(changedInputCount, changedFormulaCount))
    expect(Array.from(roots)).toEqual([a1Index, b1Index])

    const eventChanges = Effect.runSync(support.composeEventChanges(Uint32Array.of(b1Index!), explicitChangedCount))
    expect(Array.from(eventChanges)).toEqual([a1Index, b1Index])

    const eventChangesWithRootEcho = Effect.runSync(support.composeEventChanges(Uint32Array.of(a1Index!, b1Index!), explicitChangedCount))
    expect(Array.from(eventChangesWithRootEcho)).toEqual([a1Index, b1Index])

    const union = Effect.runSync(support.unionChangedSets(Uint32Array.of(a1Index!), Uint32Array.of(a1Index!, b1Index!)))
    expect(Array.from(union)).toEqual([a1Index, b1Index])

    const ordered = Effect.runSync(support.composeChangedRootsAndOrdered(Uint32Array.of(a1Index!), Uint32Array.of(a1Index!, b1Index!), 2))
    expect(Array.from(ordered)).toEqual([a1Index, b1Index])

    Effect.runSync(support.beginMutationCollection())
    expect(Effect.runSync(support.markSpillRootsChanged([a1Index!], 0))).toBe(1)
    Effect.runSync(support.beginMutationCollection())
    expect(Effect.runSync(support.markPivotRootsChanged([a1Index!], 0))).toBe(1)

    const ensuredByName = Effect.runSync(support.ensureCellTracked('Sheet1', 'C1'))
    const ensuredByCoords = Effect.runSync(support.ensureCellTrackedByCoords(engine.workbook.getSheet('Sheet1')!.id, 0, 2))
    expect(ensuredByCoords).toBe(ensuredByName)

    Effect.runSync(support.resetMaterializedCellScratch(8))
    expect(Effect.runSync(support.syncDynamicRanges(0))).toBe(0)
  })

  it('covers direct changed-set composition branches for tiny and deduplicated unions', () => {
    const support = createStubMutationSupportService()

    Effect.runSync(support.beginMutationCollection())
    expect(Effect.runSync(support.markExplicitChanged(1, 0))).toBe(1)

    const explicitEchoSecond = support.composeEventChangesNow(Uint32Array.of(2, 1), 1)
    expect(Array.from(explicitEchoSecond)).toEqual([1, 2])

    const singleDistinct = support.composeEventChangesNow(Uint32Array.of(2), 1)
    expect(Array.from(singleDistinct)).toEqual([1, 2])

    const doubleExplicit = support.composeEventChangesNow(Uint32Array.of(1, 1), 1)
    expect(Array.from(doubleExplicit)).toEqual([1])

    const explicitDistinct = support.composeEventChangesNow(Uint32Array.of(2, 3), 1)
    expect(Array.from(explicitDistinct)).toEqual([1, 2, 3])

    const disjoint = support.composeDisjointEventChangesNow(Uint32Array.of(4, 5), 1)
    expect(Array.from(disjoint)).toEqual([1, 4, 5])

    const duplicateRecalculated = support.composeEventChangesNow(Uint32Array.of(2, 2), 1)
    expect(Array.from(duplicateRecalculated)).toEqual([1, 2])

    const composed = support.composeChangedRootsAndOrderedNow(Uint32Array.of(4, 4, 5), Uint32Array.of(5, 6, 4), 3)
    expect(Array.from(composed)).toEqual([4, 5, 6])

    Effect.runSync(support.beginMutationCollection())
    support.__scratch.explicitChangedBuffer[0] = 7
    support.__scratch.explicitChangedBuffer[1] = 7
    const duplicateExplicit = support.composeEventChangesNow(Uint32Array.of(), 2)
    expect(Array.from(duplicateExplicit)).toEqual([7])
  })

  it('materializes and clears spill children through the support service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'support-spill' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    expect(a1Index).toBeDefined()

    const materialized = Effect.runSync(
      getMutationSupportService(engine).materializeSpill(a1Index!, {
        rows: 2,
        cols: 2,
        values: [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 4 },
        ],
      }),
    )

    expect(materialized.ownerValue).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'A1', rows: 2, cols: 2 }])

    Effect.runSync(getMutationSupportService(engine).clearOwnedSpill(a1Index!))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toBeUndefined()
  })

  it('reports blocked spills and missing sheet removals through the wrappers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'support-spill-blocked' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 9)

    const support = getMutationSupportService(engine)
    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    expect(a1Index).toBeDefined()

    const blocked = Effect.runSync(
      support.materializeSpill(a1Index!, {
        rows: 1,
        cols: 2,
        values: [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
        ],
      }),
    )
    expect(blocked.ownerValue).toMatchObject({
      tag: ValueTag.Error,
    })
    expect(blocked.changedCellIndices).toEqual([])
    expect(Effect.runSync(support.clearOwnedSpill(a1Index!))).toEqual([])
    expect(Effect.runSync(support.removeSheetRuntime('Missing', 0))).toEqual({
      changedInputCount: 0,
      formulaChangedCount: 0,
      explicitChangedCount: 0,
    })
  })

  it('removes sheet runtime through the service and moves selection to the next sheet', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'support-delete-sheet' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setSelection('Sheet1', 'B2')

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    expect(a1Index).toBeDefined()

    const removal = Effect.runSync(getMutationSupportService(engine).removeSheetRuntime('Sheet1', 0))

    expect(removal.changedInputCount).toBeGreaterThan(0)
    expect(removal.explicitChangedCount).toBeGreaterThan(0)
    expect(engine.workbook.getSheet('Sheet1')).toBeUndefined()
    expect(engine.getSelectionState()).toMatchObject({
      sheetName: 'Sheet2',
      address: 'A1',
      anchorAddress: 'A1',
    })
    expect((engine.workbook.cellStore.flags[a1Index!] & CellFlags.PendingDelete) !== 0).toBe(true)
  })

  it('covers direct mutation helpers for volatile formulas and deduplicated unions', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'support-direct-helpers' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A2', 7)
    engine.setCellFormula('Sheet1', 'A1', 'RAND()')
    engine.setCellFormula('Sheet1', 'B1', 'A2*2')

    const support = getMutationSupportService(engine)
    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const a2Index = engine.workbook.getCellIndex('Sheet1', 'A2')
    expect(a1Index).toBeDefined()
    expect(b1Index).toBeDefined()
    expect(a2Index).toBeDefined()

    support.beginMutationCollectionNow()
    expect(support.markVolatileFormulasChangedNow(0)).toBe(1)
    expect(Array.from(support.composeMutationRootsNow(0, 1))).toEqual([a1Index])

    const explicitChangedCount = support.markExplicitChangedNow(a2Index!, 0)
    expect(explicitChangedCount).toBe(1)
    expect(support.markPivotRootsChangedNow([a2Index!], 0)).toBe(1)

    const eventChanges = support.composeEventChangesNow(Uint32Array.of(a2Index!, a1Index!), explicitChangedCount)
    expect(Array.from(eventChanges)).toEqual([a2Index, a1Index])

    const union = support.unionChangedSetsNow(Uint32Array.of(a2Index!, a1Index!), Uint32Array.of(a2Index!, b1Index!))
    expect(Array.from(union)).toEqual([a2Index, a1Index, b1Index])

    const changedRootsAndOrdered = support.composeChangedRootsAndOrderedNow(
      Uint32Array.of(a2Index!, a1Index!),
      Uint32Array.of(a1Index!, b1Index!),
      2,
    )
    expect(Array.from(changedRootsAndOrdered)).toEqual([a2Index, a1Index, b1Index])

    support.resetMaterializedCellScratchNow(128)
  })

  it('returns a spill error when a materialized array would overflow sheet bounds', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'support-spill-overflow' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'XFD1048576', 1)

    const lastCellIndex = engine.workbook.getCellIndex('Sheet1', 'XFD1048576')
    expect(lastCellIndex).toBeDefined()

    const materialized = Effect.runSync(
      getMutationSupportService(engine).materializeSpill(lastCellIndex!, {
        rows: 1,
        cols: 2,
        values: [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
        ],
      }),
    )

    expect(materialized.ownerValue).toMatchObject({
      tag: ValueTag.Error,
      code: ErrorCode.Spill,
    })
    expect(materialized.changedCellIndices).toEqual([])
  })

  it('wraps mutation support callback failures with engine mutation errors', () => {
    const poisonedCellIndices = {
      get length() {
        throw new Error('roots boom')
      },
    } as unknown as readonly number[]

    const unionErrorService = createStubMutationSupportService({
      getChangedUnionEpoch: () => {
        throw new Error('union boom')
      },
    })
    expect(() => Effect.runSync(unionErrorService.composeEventChanges(Uint32Array.of(1), 1))).toThrow('union boom')
    expect(() => Effect.runSync(unionErrorService.unionChangedSets(Uint32Array.of(1), Uint32Array.of(2)))).toThrow('union boom')
    expect(() => Effect.runSync(unionErrorService.composeChangedRootsAndOrdered(Uint32Array.of(1), Uint32Array.of(2), 1))).toThrow(
      'union boom',
    )

    const workbookErrorService = createStubMutationSupportService({
      state: {
        workbook: {
          cellStore: {
            size: 0,
            sheetIds: [],
            rows: [],
            cols: [],
            flags: [],
            getValue: () => ({ tag: ValueTag.Empty }),
            setValue: () => undefined,
          },
          ensureCellRecord: () => {
            throw new Error('ensure by name boom')
          },
          ensureCellAt: () => {
            throw new Error('ensure coords boom')
          },
          getSheetNameById: () => {
            throw new Error('spill boom')
          },
          getAddress: () => 'A1',
          getSpill: () => undefined,
          getSheet: () => {
            throw new Error('remove sheet boom')
          },
          deleteSheet: () => undefined,
          sheetsByName: new Map(),
        },
        strings: { get: () => '' },
        formulas: {
          forEach: () => {
            throw new Error('volatile boom')
          },
        },
        ranges: { addDynamicMember: () => [] },
      } as never,
      getChangedInputBuffer: () => {
        throw new Error('buffer boom')
      },
      getMaterializedCellCount: () => {
        throw new Error('sync boom')
      },
      setMaterializedCellCount: () => {
        throw new Error('scratch boom')
      },
      ensureRecalcScratchCapacity: () => {
        throw new Error('compose roots boom')
      },
    })

    expect(() => Effect.runSync(workbookErrorService.markVolatileFormulasChanged(0))).toThrow('volatile boom')
    expect(() => Effect.runSync(workbookErrorService.markSpillRootsChanged(poisonedCellIndices, 0))).toThrow('roots boom')
    expect(() => Effect.runSync(workbookErrorService.markPivotRootsChanged(poisonedCellIndices, 0))).toThrow('roots boom')
    expect(() => Effect.runSync(workbookErrorService.composeMutationRoots(1, 1))).toThrow('compose roots boom')
    expect(() => Effect.runSync(workbookErrorService.getChangedInputBuffer())).toThrow('buffer boom')
    expect(() => Effect.runSync(workbookErrorService.ensureCellTracked('Sheet1', 'A1'))).toThrow('ensure by name boom')
    expect(() => Effect.runSync(workbookErrorService.ensureCellTrackedByCoords(1, 0, 0))).toThrow('ensure coords boom')
    expect(() => Effect.runSync(workbookErrorService.clearOwnedSpill(1))).toThrow('spill boom')
    expect(() =>
      Effect.runSync(
        workbookErrorService.materializeSpill(1, {
          rows: 1,
          cols: 1,
          values: [{ tag: ValueTag.Number, value: 1 }],
        }),
      ),
    ).toThrow('spill boom')
    expect(() => Effect.runSync(workbookErrorService.removeSheetRuntime('Sheet1', 0))).toThrow('remove sheet boom')
    expect(() => Effect.runSync(workbookErrorService.syncDynamicRanges(0))).toThrow('sync boom')
    expect(() => Effect.runSync(workbookErrorService.resetMaterializedCellScratch(4))).toThrow('scratch boom')
  })

  it('resets union epochs and grows scratch buffers in direct helper overflow paths', () => {
    let changedUnionEpoch = 0xffff_fffe

    const support = createStubMutationSupportService({
      getChangedUnionEpoch: () => changedUnionEpoch,
      setChangedUnionEpoch: (next) => {
        changedUnionEpoch = next
      },
    })

    expect(Array.from(support.composeEventChangesNow(Uint32Array.of(2), 0))).toEqual([2])
    expect(changedUnionEpoch).toBe(1)

    changedUnionEpoch = 0xffff_fffe
    expect(Array.from(support.unionChangedSetsNow(Uint32Array.of(3)))).toEqual([3])
    expect(changedUnionEpoch).toBe(1)

    changedUnionEpoch = 0xffff_fffe
    expect(Array.from(support.composeChangedRootsAndOrderedNow(Uint32Array.of(4), Uint32Array.of(4), 1))).toEqual([4])
    expect(changedUnionEpoch).toBe(1)

    support.resetMaterializedCellScratchNow(8)
  })
})
