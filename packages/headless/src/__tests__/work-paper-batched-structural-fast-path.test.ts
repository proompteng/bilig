import { ValueTag } from '@bilig/protocol'
import { describe, expect, it, vi } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress, type WorkPaperChange } from '../index.js'
import { hasDeferredTrackedIndexChanges } from '../tracked-cell-index-changes.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function hasCaptureVisibilitySnapshot(value: unknown): value is WorkPaper & { captureVisibilitySnapshot: () => unknown } {
  return typeof Reflect.get(value, 'captureVisibilitySnapshot') === 'function'
}

interface TestSheetDimensionCache {
  updateAfterCellMutationRefs(...args: unknown[]): unknown
  updateAfterMatrixMutationImpact(...args: unknown[]): unknown
}

interface TestEngineRuntimeSupport {
  clearOwnedSpillNow(...args: unknown[]): unknown
}

interface TestEngineFormulaBindingSupport {
  getFormulaFamilyStatsNow(): { familyCount: number; runCount: number; memberCount: number }
  upsertFreshFormulaInstancesNow(...args: unknown[]): unknown
}

interface TestEngineCellMutationApplySupport {
  applyCellMutationsAtWithOptions(...args: unknown[]): unknown
}

interface TestFormulaFamilyStore {
  registerFormulaRun(...args: unknown[]): unknown
  upsertFormula(...args: unknown[]): unknown
}

interface TestCoreWorkbookSheetLookup {
  getSheetById(sheetId: number): unknown
}

interface TestCoreWorkbookFormulaLookup {
  getCellIndex(sheetName: string, address: string): number | undefined
}

interface TestRuntimeFormula {
  readonly dependencyIndices: Uint32Array
  readonly rangeDependencies: Uint32Array
  readonly directAggregate: unknown
  readonly directScalar: unknown
}

interface TestRuntimeFormulaTable {
  get(cellIndex: number): unknown
}

interface TestFreshDenseAttachSheet {
  readonly grid: {
    createRowMajorSetter(...args: unknown[]): unknown
    setDenseRowMajor(...args: unknown[]): unknown
  }
  readonly logical: {
    setFreshVisibleCellIdentityWithAxisIdsDeferred(...args: unknown[]): unknown
    setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred(...args: unknown[]): unknown
  }
}

function hasSheetDimensionCacheUpdater(value: unknown): value is TestSheetDimensionCache {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof Reflect.get(value, 'updateAfterCellMutationRefs') === 'function' &&
    typeof Reflect.get(value, 'updateAfterMatrixMutationImpact') === 'function'
  )
}

function hasClearOwnedSpill(value: unknown): value is TestEngineRuntimeSupport {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'clearOwnedSpillNow') === 'function'
}

function hasFormulaBindingSupport(value: unknown): value is TestEngineFormulaBindingSupport {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof Reflect.get(value, 'getFormulaFamilyStatsNow') === 'function' &&
    typeof Reflect.get(value, 'upsertFreshFormulaInstancesNow') === 'function'
  )
}

function hasCellMutationApplySupport(value: unknown): value is TestEngineCellMutationApplySupport {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'applyCellMutationsAtWithOptions') === 'function'
}

function hasFormulaFamilyStore(value: unknown): value is TestFormulaFamilyStore {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof Reflect.get(value, 'registerFormulaRun') === 'function' &&
    typeof Reflect.get(value, 'upsertFormula') === 'function'
  )
}

function hasCoreWorkbookSheetLookup(value: unknown): value is TestCoreWorkbookSheetLookup {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'getSheetById') === 'function'
}

function hasCoreWorkbookFormulaLookup(value: unknown): value is TestCoreWorkbookFormulaLookup {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'getCellIndex') === 'function'
}

function hasRuntimeFormulaTable(value: unknown): value is TestRuntimeFormulaTable {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'get') === 'function'
}

function isRuntimeFormula(value: unknown): value is TestRuntimeFormula {
  return (
    typeof value === 'object' &&
    value !== null &&
    Reflect.get(value, 'dependencyIndices') instanceof Uint32Array &&
    Reflect.get(value, 'rangeDependencies') instanceof Uint32Array &&
    (Reflect.get(value, 'directAggregate') !== undefined || Reflect.get(value, 'directScalar') !== undefined)
  )
}

function hasFreshDenseAttachSheet(value: unknown): value is TestFreshDenseAttachSheet {
  const logical = typeof value === 'object' && value !== null ? Reflect.get(value, 'logical') : undefined
  const grid = typeof value === 'object' && value !== null ? Reflect.get(value, 'grid') : undefined
  return (
    typeof logical === 'object' &&
    logical !== null &&
    typeof Reflect.get(logical, 'setFreshVisibleCellIdentityWithAxisIdsDeferred') === 'function' &&
    typeof Reflect.get(logical, 'setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred') === 'function' &&
    typeof grid === 'object' &&
    grid !== null &&
    typeof Reflect.get(grid, 'createRowMajorSetter') === 'function' &&
    typeof Reflect.get(grid, 'setDenseRowMajor') === 'function'
  )
}

function isCellChangeArray(changes: WorkPaperChange[]): changes is Extract<WorkPaperChange, { kind: 'cell' }>[] {
  return changes.every((change) => change.kind === 'cell')
}

function trackSheetDimensionCacheUpdates(workbook: WorkPaper): {
  readonly matrixImpactCount: number
  readonly refScanCount: number
  restore: () => void
} {
  const cache: unknown = Reflect.get(workbook, 'sheetDimensionCache')
  if (!hasSheetDimensionCacheUpdater(cache)) {
    throw new Error('Expected WorkPaper to expose a sheet dimension cache in tests')
  }
  const refScanSpy = vi.spyOn(cache, 'updateAfterCellMutationRefs')
  const matrixImpactSpy = vi.spyOn(cache, 'updateAfterMatrixMutationImpact')
  return {
    get matrixImpactCount() {
      return matrixImpactSpy.mock.calls.length
    },
    get refScanCount() {
      return refScanSpy.mock.calls.length
    },
    restore: () => {
      matrixImpactSpy.mockRestore()
      refScanSpy.mockRestore()
    },
  }
}

function trackCoreSpillOwnerClears(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const engine: unknown = Reflect.get(workbook, 'engine')
  const runtime: unknown = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'runtime') : undefined
  const support: unknown = typeof runtime === 'object' && runtime !== null ? Reflect.get(runtime, 'support') : undefined
  if (!hasClearOwnedSpill(support)) {
    throw new Error('Expected WorkPaper to expose core spill-owner cleanup in tests')
  }
  const spy = vi.spyOn(support, 'clearOwnedSpillNow')
  return {
    get count() {
      return spy.mock.calls.length
    },
    restore: () => {
      spy.mockRestore()
    },
  }
}

function getCoreFormulaBindingSupport(workbook: WorkPaper): TestEngineFormulaBindingSupport {
  const engine: unknown = Reflect.get(workbook, 'engine')
  const runtime: unknown = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'runtime') : undefined
  const binding: unknown = typeof runtime === 'object' && runtime !== null ? Reflect.get(runtime, 'binding') : undefined
  if (!hasFormulaBindingSupport(binding)) {
    throw new Error('Expected WorkPaper to expose formula binding support in tests')
  }
  return binding
}

function getCoreFormulaFamilyStore(workbook: WorkPaper): TestFormulaFamilyStore {
  const engine: unknown = Reflect.get(workbook, 'engine')
  const runtime: unknown = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'runtime') : undefined
  const formulaFamilies: unknown = typeof runtime === 'object' && runtime !== null ? Reflect.get(runtime, 'formulaFamilies') : undefined
  if (!hasFormulaFamilyStore(formulaFamilies)) {
    throw new Error('Expected WorkPaper to expose formula family store in tests')
  }
  return formulaFamilies
}

function getCoreRuntimeFormula(workbook: WorkPaper, sheetName: string, address: string): TestRuntimeFormula {
  const engine: unknown = Reflect.get(workbook, 'engine')
  const coreWorkbook: unknown = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'workbook') : undefined
  if (!hasCoreWorkbookFormulaLookup(coreWorkbook)) {
    throw new Error('Expected WorkPaper to expose core workbook cell lookup in tests')
  }
  const cellIndex = coreWorkbook.getCellIndex(sheetName, address)
  if (cellIndex === undefined) {
    throw new Error(`Expected runtime cell at ${sheetName}!${address}`)
  }
  const formulas: unknown = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'formulas') : undefined
  if (!hasRuntimeFormulaTable(formulas)) {
    throw new Error('Expected WorkPaper to expose core runtime formulas in tests')
  }
  const formula = formulas.get(cellIndex)
  if (!isRuntimeFormula(formula)) {
    throw new Error(`Expected runtime direct aggregate formula at ${sheetName}!${address}`)
  }
  return formula
}

describe('work paper batched structural fast path', () => {
  it('keeps appended formula rows on the tracked batch path', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, 2, '=SUM(A1:B1)'],
        [3, 4, '=SUM(A2:B2)'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    expect(hasCaptureVisibilitySnapshot(workbook)).toBe(true)
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected WorkPaper to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('batched append formulas should not rebuild visibility snapshots')
    })
    const dimensionUpdates = trackSheetDimensionCacheUpdates(workbook)
    const spillOwnerClears = trackCoreSpillOwnerClears(workbook)

    let changes: WorkPaperChange[] = []
    try {
      changes = workbook.batch(() => {
        expect(workbook.addRows(sheetId, 2, 2)).toEqual([])
        expect(
          workbook.setCellContents(cell(sheetId, 2, 0), [
            [5, 6, '=SUM(A3:B3)'],
            [7, 8, '=SUM(A4:B4)'],
          ]),
        ).toEqual([])
      })
      expect(dimensionUpdates.matrixImpactCount).toBe(1)
      expect(dimensionUpdates.refScanCount).toBe(0)
      expect(spillOwnerClears.count).toBe(0)
    } finally {
      spillOwnerClears.restore()
      dimensionUpdates.restore()
      captureVisibilitySnapshot.mockRestore()
    }

    expect(changes.length).toBeGreaterThan(0)
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(workbook.getCellValue(cell(sheetId, 3, 2))).toEqual({ tag: ValueTag.Number, value: 15 })

    workbook.setCellContents(cell(sheetId, 2, 0), 10)
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 16 })

    workbook.undo()
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
    workbook.undo()
    expect(workbook.getSheetDimensions(sheetId).height).toBe(2)
  })

  it('keeps large appended formula-row changes lazy after structural edits', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, 2, '=SUM(A1:B1)'],
        [3, 4, '=SUM(A2:B2)'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    const appendCount = 96
    const engine: unknown = Reflect.get(workbook, 'engine')
    if (!hasCellMutationApplySupport(engine)) {
      throw new Error('Expected WorkPaper to expose cell mutation application in tests')
    }
    const coreWorkbook: unknown = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'workbook') : undefined
    if (!hasCoreWorkbookSheetLookup(coreWorkbook)) {
      throw new Error('Expected WorkPaper to expose the core workbook in tests')
    }
    const coreSheet = coreWorkbook.getSheetById(sheetId)
    if (!hasFreshDenseAttachSheet(coreSheet)) {
      throw new Error('Expected WorkPaper to expose fresh dense attach internals in tests')
    }
    const applyCellMutations = vi.spyOn(engine, 'applyCellMutationsAtWithOptions')
    const denseIdentityAttach = vi.spyOn(coreSheet.logical, 'setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred')
    const singleIdentityAttach = vi.spyOn(coreSheet.logical, 'setFreshVisibleCellIdentityWithAxisIdsDeferred')
    const denseGridAttach = vi.spyOn(coreSheet.grid, 'setDenseRowMajor')
    const rowMajorSetter = vi.spyOn(coreSheet.grid, 'createRowMajorSetter')
    const binding = getCoreFormulaBindingSupport(workbook)
    expect(binding.getFormulaFamilyStatsNow()).toEqual({ familyCount: 1, runCount: 1, memberCount: 2 })
    const formulaInstanceBulkUpsert = vi.spyOn(binding, 'upsertFreshFormulaInstancesNow')
    const formulaFamilies = getCoreFormulaFamilyStore(workbook)
    const formulaFamilyRunRegistration = vi.spyOn(formulaFamilies, 'registerFormulaRun')
    const perFormulaFamilyRegistration = vi.spyOn(formulaFamilies, 'upsertFormula')
    const dimensionUpdates = trackSheetDimensionCacheUpdates(workbook)
    const rows = Array.from({ length: appendCount }, (_, index) => {
      const rowNumber = index + 3
      return [rowNumber, rowNumber * 2, `=SUM(A${rowNumber}:B${rowNumber})`]
    })

    let changes: WorkPaperChange[] = []
    try {
      workbook.resetPerformanceCounters()
      changes = workbook.batch(() => {
        workbook.addRows(sheetId, 2, appendCount)
        workbook.setCellContents(cell(sheetId, 2, 0), rows)
      })

      expect(changes).toHaveLength(appendCount * 3)
      expect(isCellChangeArray(changes)).toBe(true)
      if (!isCellChangeArray(changes)) {
        throw new Error('Expected appended formula rows to emit only cell changes')
      }
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(changes[0]).toMatchObject({
        a1: 'A3',
        newValue: { tag: ValueTag.Number, value: 3 },
      })
      expect(changes.at(-1)).toMatchObject({
        a1: `C${appendCount + 2}`,
        newValue: { tag: ValueTag.Number, value: (appendCount + 2) * 3 },
      })
      expect(applyCellMutations).toHaveBeenCalledTimes(1)
      expect(applyCellMutations.mock.calls[0]?.[0]).toHaveLength(appendCount * 3)
      expect(applyCellMutations.mock.calls[0]?.[1]).toMatchObject({
        potentialNewCells: appendCount * 3,
      })
      expect(denseIdentityAttach).toHaveBeenCalledTimes(1)
      expect(denseIdentityAttach.mock.calls[0]?.[1]).toHaveLength(appendCount)
      expect(denseIdentityAttach.mock.calls[0]?.[2]).toHaveLength(3)
      expect(singleIdentityAttach).not.toHaveBeenCalled()
      expect(denseGridAttach).toHaveBeenCalledWith(2, 0, appendCount, 3, expect.any(Number))
      expect(rowMajorSetter).not.toHaveBeenCalled()
      expect(formulaFamilyRunRegistration).toHaveBeenCalledTimes(1)
      expect(perFormulaFamilyRegistration).not.toHaveBeenCalled()
      expect(formulaInstanceBulkUpsert).toHaveBeenCalledTimes(1)
      expect(formulaInstanceBulkUpsert.mock.calls[0]?.[0]).toHaveLength(appendCount)
      expect(binding.getFormulaFamilyStatsNow()).toEqual({ familyCount: 1, runCount: 1, memberCount: appendCount + 2 })
      expect(dimensionUpdates.matrixImpactCount).toBe(1)
      expect(dimensionUpdates.refScanCount).toBe(0)
      expect(workbook.getPerformanceCounters()).toMatchObject({
        calcChainFullScans: 0,
        directAggregateScanEvaluations: 0,
        directFormulaKernelSyncOnlyRecalcSkips: 1,
        kernelSyncOnlyRecalcSkips: 1,
        regionQueryIndexBuilds: 0,
        topoRepairs: 0,
      })
    } finally {
      applyCellMutations.mockRestore()
      denseIdentityAttach.mockRestore()
      singleIdentityAttach.mockRestore()
      denseGridAttach.mockRestore()
      rowMajorSetter.mockRestore()
      formulaInstanceBulkUpsert.mockRestore()
      formulaFamilyRunRegistration.mockRestore()
      perFormulaFamilyRegistration.mockRestore()
      dimensionUpdates.restore()
    }
  })

  it('bulk-binds fresh appended direct-scalar formula matrices after structural edits', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, 2, '=A1+B1'],
        [3, 4, '=A2+B2'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    const appendCount = 96
    const engine: unknown = Reflect.get(workbook, 'engine')
    if (!hasCellMutationApplySupport(engine)) {
      throw new Error('Expected WorkPaper to expose cell mutation application in tests')
    }
    const coreWorkbook: unknown = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'workbook') : undefined
    if (!hasCoreWorkbookSheetLookup(coreWorkbook)) {
      throw new Error('Expected WorkPaper to expose the core workbook in tests')
    }
    const coreSheet = coreWorkbook.getSheetById(sheetId)
    if (!hasFreshDenseAttachSheet(coreSheet)) {
      throw new Error('Expected WorkPaper to expose fresh dense attach internals in tests')
    }
    const applyCellMutations = vi.spyOn(engine, 'applyCellMutationsAtWithOptions')
    const denseIdentityAttach = vi.spyOn(coreSheet.logical, 'setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred')
    const singleIdentityAttach = vi.spyOn(coreSheet.logical, 'setFreshVisibleCellIdentityWithAxisIdsDeferred')
    const denseGridAttach = vi.spyOn(coreSheet.grid, 'setDenseRowMajor')
    const rowMajorSetter = vi.spyOn(coreSheet.grid, 'createRowMajorSetter')
    const binding = getCoreFormulaBindingSupport(workbook)
    expect(binding.getFormulaFamilyStatsNow()).toEqual({ familyCount: 1, runCount: 1, memberCount: 2 })
    const formulaInstanceBulkUpsert = vi.spyOn(binding, 'upsertFreshFormulaInstancesNow')
    const formulaFamilies = getCoreFormulaFamilyStore(workbook)
    const formulaFamilyRunRegistration = vi.spyOn(formulaFamilies, 'registerFormulaRun')
    const perFormulaFamilyRegistration = vi.spyOn(formulaFamilies, 'upsertFormula')
    const dimensionUpdates = trackSheetDimensionCacheUpdates(workbook)
    const rows = Array.from({ length: appendCount }, (_, index) => {
      const rowNumber = index + 3
      return [rowNumber, rowNumber * 2, `=A${rowNumber}+B${rowNumber}`]
    })

    try {
      workbook.resetPerformanceCounters()
      const changes = workbook.batch(() => {
        workbook.addRows(sheetId, 2, appendCount)
        workbook.setCellContents(cell(sheetId, 2, 0), rows)
      })

      expect(changes).toHaveLength(appendCount * 3)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(applyCellMutations).toHaveBeenCalledTimes(1)
      expect(applyCellMutations.mock.calls[0]?.[0]).toHaveLength(appendCount * 3)
      expect(denseIdentityAttach).toHaveBeenCalledTimes(1)
      expect(denseIdentityAttach.mock.calls[0]?.[1]).toHaveLength(appendCount)
      expect(denseIdentityAttach.mock.calls[0]?.[2]).toHaveLength(3)
      expect(singleIdentityAttach).not.toHaveBeenCalled()
      expect(denseGridAttach).toHaveBeenCalledWith(2, 0, appendCount, 3, expect.any(Number))
      expect(rowMajorSetter).not.toHaveBeenCalled()
      expect(formulaFamilyRunRegistration).toHaveBeenCalledTimes(1)
      expect(perFormulaFamilyRegistration).not.toHaveBeenCalled()
      expect(formulaInstanceBulkUpsert).toHaveBeenCalledTimes(1)
      expect(formulaInstanceBulkUpsert.mock.calls[0]?.[0]).toHaveLength(appendCount)
      expect(binding.getFormulaFamilyStatsNow()).toEqual({ familyCount: 1, runCount: 1, memberCount: appendCount + 2 })
      expect(dimensionUpdates.matrixImpactCount).toBe(1)
      expect(dimensionUpdates.refScanCount).toBe(0)
      expect(workbook.getPerformanceCounters()).toMatchObject({
        calcChainFullScans: 0,
        directFormulaKernelSyncOnlyRecalcSkips: 1,
        kernelSyncOnlyRecalcSkips: 1,
        regionQueryIndexBuilds: 0,
        topoRebuilds: 0,
        topoRepairs: 0,
      })
      expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 9 })
      expect(workbook.getCellValue(cell(sheetId, appendCount + 1, 2))).toEqual({
        tag: ValueTag.Number,
        value: (appendCount + 2) * 3,
      })

      const runtimeFormula = getCoreRuntimeFormula(workbook, 'Data', 'C3')
      expect(runtimeFormula.dependencyIndices).toHaveLength(2)
      expect(runtimeFormula.rangeDependencies).toHaveLength(0)
      expect(runtimeFormula.directScalar).toMatchObject({ kind: 'binary', operator: '+' })

      workbook.setCellContents(cell(sheetId, 2, 0), 10)
      expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 16 })
    } finally {
      applyCellMutations.mockRestore()
      denseIdentityAttach.mockRestore()
      singleIdentityAttach.mockRestore()
      denseGridAttach.mockRestore()
      rowMajorSetter.mockRestore()
      formulaInstanceBulkUpsert.mockRestore()
      formulaFamilyRunRegistration.mockRestore()
      perFormulaFamilyRegistration.mockRestore()
      dimensionUpdates.restore()
    }
  })

  it('keeps fresh appended aggregate matrices dependency-light when input columns already contain formulas', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        ['=B1+1', 2, '=SUM(A1:B1)'],
        [3, 4, '=SUM(A2:B2)'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    const appendCount = 40
    const rows = Array.from({ length: appendCount }, (_, index) => {
      const rowNumber = index + 3
      return [rowNumber, rowNumber * 2, `=SUM(A${rowNumber}:B${rowNumber})`]
    })

    workbook.batch(() => {
      workbook.addRows(sheetId, 2, appendCount)
      workbook.setCellContents(cell(sheetId, 2, 0), rows)
    })

    const runtimeFormula = getCoreRuntimeFormula(workbook, 'Data', 'C3')
    expect(runtimeFormula.dependencyIndices).toHaveLength(0)
    expect(runtimeFormula.rangeDependencies).toHaveLength(0)
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 9 })

    workbook.setCellContents(cell(sheetId, 2, 0), 10)
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 16 })
  })

  it('trusts physical tracked changes for fresh appended formula matrices after structural edits', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, 2, '=SUM(A1:B1)'],
        [3, 4, '=SUM(A2:B2)'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    const appendCount = 96
    const rows = Array.from({ length: appendCount }, (_, index) => {
      const rowNumber = index + 3
      return [rowNumber, rowNumber * 2, `=SUM(A${rowNumber}:B${rowNumber})`]
    })

    const engine: unknown = Reflect.get(workbook, 'engine')
    const coreWorkbook: unknown = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'workbook') : undefined
    if (typeof coreWorkbook !== 'object' || coreWorkbook === null || typeof Reflect.get(coreWorkbook, 'getCellPosition') !== 'function') {
      throw new Error('Expected WorkPaper to expose the core workbook in tests')
    }
    const positionSpy = vi.spyOn(coreWorkbook, 'getCellPosition')

    try {
      const changes = workbook.batch(() => {
        workbook.addRows(sheetId, 2, appendCount)
        workbook.setCellContents(cell(sheetId, 2, 0), rows)
      })

      expect(changes).toHaveLength(appendCount * 3)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(positionSpy).not.toHaveBeenCalled()
      expect(changes[0]).toMatchObject({
        a1: 'A3',
        newValue: { tag: ValueTag.Number, value: 3 },
      })
      expect(changes.at(-1)).toMatchObject({
        a1: `C${appendCount + 2}`,
        newValue: { tag: ValueTag.Number, value: (appendCount + 2) * 3 },
      })
      expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 9 })
      expect(workbook.getCellValue(cell(sheetId, appendCount + 1, 2))).toEqual({
        tag: ValueTag.Number,
        value: (appendCount + 2) * 3,
      })
    } finally {
      positionSpy.mockRestore()
    }
  })

  it('recalculates existing direct ranges that overlap appended formula rows', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, 2, '=SUM(A1:B1)', undefined, undefined, '=SUM(A1:A4)'],
        [3, 4, '=SUM(A2:B2)'],
        [5, 6, '=SUM(A3:B3)'],
        [7, 8, '=SUM(A4:B4)'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!

    const changes = workbook.batch(() => {
      expect(workbook.addRows(sheetId, 2, 2)).toEqual([])
      expect(
        workbook.setCellContents(cell(sheetId, 2, 0), [
          [5, 6, '=SUM(A3:B3)'],
          [7, 8, '=SUM(A4:B4)'],
        ]),
      ).toEqual([])
    })

    expect(changes.length).toBeGreaterThan(0)
    expect(workbook.getCellFormula(cell(sheetId, 0, 5))).toBe('=SUM(A1:A6)')
    expect(workbook.getCellValue(cell(sheetId, 0, 5))).toEqual({ tag: ValueTag.Number, value: 28 })
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(workbook.getCellValue(cell(sheetId, 3, 2))).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('skips tracked cell change reduction when structural inserts change no values', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, 2, '=A1+B1'],
        [3, 4, '=A2+B2'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    const originalComputeTrackedChanges = Reflect.get(workbook, 'computeTrackedChangesWithoutVisibilityCache')
    if (typeof originalComputeTrackedChanges !== 'function') {
      throw new Error('Expected WorkPaper to expose tracked change reduction in tests')
    }
    let computeTrackedChangeCalls = 0
    Reflect.set(workbook, 'computeTrackedChangesWithoutVisibilityCache', (...args: unknown[]) => {
      computeTrackedChangeCalls += 1
      return Reflect.apply(originalComputeTrackedChanges, workbook, args)
    })

    try {
      const changes = workbook.addColumns(sheetId, 1, 1)

      expect(changes).toEqual([])
      expect(computeTrackedChangeCalls).toBe(0)
      expect(workbook.getCellFormula(cell(sheetId, 0, 3))).toBe('=A1+C1')
      expect(workbook.getCellValue(cell(sheetId, 0, 3))).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 4, height: 2 })
    } finally {
      Reflect.set(workbook, 'computeTrackedChangesWithoutVisibilityCache', originalComputeTrackedChanges)
    }
  })
})
