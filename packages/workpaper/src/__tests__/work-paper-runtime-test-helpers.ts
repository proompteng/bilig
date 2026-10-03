import { afterEach, expect, vi } from 'vitest'
export { ErrorCode, ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
export { afterEach, describe, expect, it, vi } from 'vitest'
export {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  WorkPaper,
  WorkPaperEvaluationSuspendedError,
} from '../index.js'
export type { WorkPaperCellAddress, WorkPaperCellChange, WorkPaperChange } from '../index.js'
export { hasDeferredTrackedIndexChanges } from '../tracked-cell-index-changes.js'

import type { WorkPaperCellAddress, WorkPaperCellChange, WorkPaperChange } from '../index.js'
import { WorkPaper } from '../index.js'

export const TEST_LANGUAGE_CODE = 'xHF'

export function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

export function columnLabel(index: number): string {
  let value = index + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

export function hasCaptureVisibilitySnapshot(value: unknown): value is WorkPaper & { captureVisibilitySnapshot: () => unknown } {
  return typeof Reflect.get(value, 'captureVisibilitySnapshot') === 'function'
}

export function trackPrivateMethod(workbook: WorkPaper, methodName: string): { readonly count: number; restore: () => void } {
  const original = Reflect.get(workbook, methodName)
  if (typeof original !== 'function') {
    throw new Error(`Expected WorkPaper to expose ${methodName} in tests`)
  }
  let count = 0
  Reflect.set(workbook, methodName, (...args: unknown[]) => {
    count += 1
    return Reflect.apply(original, workbook, args)
  })
  return {
    get count() {
      return count
    },
    restore: () => {
      Reflect.set(workbook, methodName, original)
    },
  }
}

export interface TestSheetDimensionCache {
  updateAfterCellMutationRefs(...args: unknown[]): unknown
}

export function hasSheetDimensionCacheUpdater(value: unknown): value is TestSheetDimensionCache {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'updateAfterCellMutationRefs') === 'function'
}

export function trackSheetDimensionCacheUpdates(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const cache: unknown = Reflect.get(workbook, 'sheetDimensionCache')
  if (!hasSheetDimensionCacheUpdater(cache)) {
    throw new Error('Expected WorkPaper to expose a sheet dimension cache in tests')
  }
  const spy = vi.spyOn(cache, 'updateAfterCellMutationRefs')
  return {
    get count() {
      return spy.mock.calls.length
    },
    restore: () => {
      spy.mockRestore()
    },
  }
}

export function readEngineUseColumnIndexEnabled(workbook: WorkPaper): boolean {
  const engine = Reflect.get(workbook, 'engine')
  if (typeof engine !== 'object' || engine === null) {
    throw new Error('Expected WorkPaper to expose an engine object in tests')
  }
  return Reflect.get(engine, 'useColumnIndexEnabled') === true
}

export interface WorkPaperFormulaBindingTestSurface {
  collectFormulaCellsReferencingSheetNow(sheetName: string): readonly number[]
}

export function hasFormulaBindingTestSurface(value: unknown): value is WorkPaperFormulaBindingTestSurface {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'collectFormulaCellsReferencingSheetNow') === 'function'
}

export function readFormulaBindingTestSurface(workbook: WorkPaper): WorkPaperFormulaBindingTestSurface {
  const engine = Reflect.get(workbook, 'engine')
  const runtime = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'runtime') : undefined
  const binding = typeof runtime === 'object' && runtime !== null ? Reflect.get(runtime, 'binding') : undefined
  if (!hasFormulaBindingTestSurface(binding)) {
    throw new Error('Expected WorkPaper to expose formula binding internals in tests')
  }
  return binding
}

export function expectOnlyCellChanges(changes: WorkPaperChange[]): asserts changes is WorkPaperCellChange[] {
  expect(changes.every((change) => change.kind === 'cell')).toBe(true)
}

export function trackComputeCellChangesFromTrackedEvents(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const original = Reflect.get(workbook, 'computeCellChangesFromTrackedEvents')
  if (typeof original !== 'function') {
    throw new Error('Expected WorkPaper to expose computeCellChangesFromTrackedEvents in tests')
  }
  let count = 0
  Reflect.set(workbook, 'computeCellChangesFromTrackedEvents', (...args: unknown[]) => {
    count += 1
    return Reflect.apply(original, workbook, args)
  })
  return {
    get count() {
      return count
    },
    restore: () => {
      Reflect.set(workbook, 'computeCellChangesFromTrackedEvents', original)
    },
  }
}

export function rejectSingleTrackedCellReader(workbook: WorkPaper): { restore: () => void } {
  const original = Reflect.get(workbook, 'readSingleTrackedCellChange')
  if (typeof original !== 'function') {
    throw new Error('Expected WorkPaper to expose readSingleTrackedCellChange in tests')
  }
  Reflect.set(workbook, 'readSingleTrackedCellChange', () => {
    throw new Error('Expected tiny sorted physical changes to avoid generic single-cell reading')
  })
  return {
    restore: () => {
      Reflect.set(workbook, 'readSingleTrackedCellChange', original)
    },
  }
}

export function trackCaptureTrackedChangesWithoutVisibilityCache(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const original = Reflect.get(workbook, 'captureTrackedChangesWithoutVisibilityCache')
  if (typeof original !== 'function') {
    throw new Error('Expected WorkPaper to expose captureTrackedChangesWithoutVisibilityCache in tests')
  }
  let count = 0
  Reflect.set(workbook, 'captureTrackedChangesWithoutVisibilityCache', (...args: unknown[]) => {
    count += 1
    return Reflect.apply(original, workbook, args)
  })
  return {
    get count() {
      return count
    },
    restore: () => {
      Reflect.set(workbook, 'captureTrackedChangesWithoutVisibilityCache', original)
    },
  }
}

export interface EngineApplyCellMutationsTarget {
  applyCellMutationsAtWithOptions: (...args: unknown[]) => unknown
}

export interface EngineExistingNumericCellMutationsTarget {
  tryApplyExistingNumericCellMutationsAt: (...args: unknown[]) => unknown
}

export interface SheetGridEntryTarget {
  forEachCellEntry: (fn: (cellIndex: number, row: number, col: number) => void) => void
}

export interface SheetRecordTarget {
  grid: SheetGridEntryTarget
}

export interface EngineWorkbookTarget {
  workbook: {
    getSheetById(sheetId: number): SheetRecordTarget | undefined
  }
}

export interface EngineFormulaBindingTarget {
  runtime: {
    binding: {
      forEachFormulaFamilyNow: (fn: (...args: unknown[]) => void) => void
      isFormulaFamilyIndexReadyNow: () => boolean
      rebindFormulaCellsNow: (inputs: readonly unknown[], formulaChangedCount?: number) => unknown
    }
  }
}

export function isEngineApplyCellMutationsTarget(value: unknown): value is EngineApplyCellMutationsTarget {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'applyCellMutationsAtWithOptions') === 'function'
}

export function isEngineExistingNumericCellMutationsTarget(value: unknown): value is EngineExistingNumericCellMutationsTarget {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'tryApplyExistingNumericCellMutationsAt') === 'function'
}

export function isEngineWorkbookTarget(value: unknown): value is EngineWorkbookTarget {
  const workbook = typeof value === 'object' && value !== null ? Reflect.get(value, 'workbook') : undefined
  return typeof workbook === 'object' && workbook !== null && typeof Reflect.get(workbook, 'getSheetById') === 'function'
}

export function isEngineFormulaBindingTarget(value: unknown): value is EngineFormulaBindingTarget {
  const runtime = typeof value === 'object' && value !== null ? Reflect.get(value, 'runtime') : undefined
  const binding = typeof runtime === 'object' && runtime !== null ? Reflect.get(runtime, 'binding') : undefined
  return (
    typeof binding === 'object' &&
    binding !== null &&
    typeof Reflect.get(binding, 'forEachFormulaFamilyNow') === 'function' &&
    typeof Reflect.get(binding, 'isFormulaFamilyIndexReadyNow') === 'function' &&
    typeof Reflect.get(binding, 'rebindFormulaCellsNow') === 'function'
  )
}

export function engineApplyCellMutationsTarget(workbook: WorkPaper): EngineApplyCellMutationsTarget {
  const engine = Reflect.get(workbook, 'engine')
  if (!isEngineApplyCellMutationsTarget(engine)) {
    throw new Error('Expected WorkPaper to expose applyCellMutationsAtWithOptions in tests')
  }
  return engine
}

export function engineExistingNumericCellMutationsTarget(workbook: WorkPaper): EngineExistingNumericCellMutationsTarget {
  const engine = Reflect.get(workbook, 'engine')
  if (!isEngineExistingNumericCellMutationsTarget(engine)) {
    throw new Error('Expected WorkPaper to expose tryApplyExistingNumericCellMutationsAt in tests')
  }
  return engine
}

export function sheetGridEntryTarget(workbook: WorkPaper, sheetId: number): SheetGridEntryTarget {
  const engine = Reflect.get(workbook, 'engine')
  if (!isEngineWorkbookTarget(engine)) {
    throw new Error('Expected WorkPaper to expose workbook in tests')
  }
  const sheet = engine?.workbook?.getSheetById(sheetId)
  if (!sheet) {
    throw new Error('Expected WorkPaper to expose sheet grid in tests')
  }
  return sheet.grid
}

export function engineFormulaBindingTarget(workbook: WorkPaper): EngineFormulaBindingTarget['runtime']['binding'] {
  const engine = Reflect.get(workbook, 'engine')
  if (!isEngineFormulaBindingTarget(engine)) {
    throw new Error('Expected WorkPaper to expose formula binding service in tests')
  }
  return engine.runtime.binding
}

export function readUndoStack(value: unknown): unknown[] | null {
  const engine = Reflect.get(value, 'engine')
  if (!engine || typeof engine !== 'object') {
    return null
  }
  const undoStack = Reflect.get(engine, 'undoStack')
  return Array.isArray(undoStack) ? undoStack : null
}

afterEach(() => {
  WorkPaper.unregisterAllFunctions()
  if (WorkPaper.getRegisteredLanguagesCodes().includes(TEST_LANGUAGE_CODE)) {
    WorkPaper.unregisterLanguage(TEST_LANGUAGE_CODE)
  }
})
