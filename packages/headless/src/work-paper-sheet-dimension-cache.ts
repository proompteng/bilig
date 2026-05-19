import type { EngineCellMutationRef } from '@bilig/core/headless-runtime'
import { workPaperFormulaMayResizeDynamically } from './work-paper-sheet-inspection.js'
import type { WorkPaperSheetDimensions } from './work-paper-types.js'
import type { WorkPaperAxisIntervalEditMode, WorkPaperAxisKind } from './work-paper-axis-helpers.js'
import type { MatrixMutationDimensionImpact } from './matrix-mutation-plan.js'

export interface WorkPaperSheetDimensionEngine {
  readonly workbook: {
    listSpills(): readonly { readonly sheetName: string }[]
    getSheet(sheetName: string): { readonly id: number } | undefined
  }
}

export interface WorkPaperSheetDimensionRecord {
  readonly grid: {
    forEachCellEntry(callback: (cellIndex: number, row: number, col: number) => void): void
  }
}

export class WorkPaperSheetDimensionCache {
  private readonly dimensions = new Map<number, WorkPaperSheetDimensions>()
  private spillSheetIds: Set<number> | null = null

  constructor(private readonly engine: WorkPaperSheetDimensionEngine) {}

  get(sheetId: number): WorkPaperSheetDimensions | undefined {
    return this.dimensions.get(sheetId)
  }

  cache(sheetId: number, dimensions: WorkPaperSheetDimensions): void {
    this.dimensions.set(sheetId, { width: dimensions.width, height: dimensions.height })
  }

  cacheInitialized(sheetId: number, dimensions: WorkPaperSheetDimensions, options: { readonly mayResizeDynamically?: boolean } = {}): void {
    if (options.mayResizeDynamically === true || this.sheetHasSpills(sheetId)) {
      this.invalidate(sheetId)
      return
    }
    this.cache(sheetId, dimensions)
  }

  cacheScanned(sheetId: number, dimensions: WorkPaperSheetDimensions, options: { readonly mayResizeDynamically?: boolean } = {}): void {
    if (options.mayResizeDynamically === true || this.sheetHasSpills(sheetId)) {
      this.invalidate(sheetId)
      return
    }
    this.cache(sheetId, dimensions)
  }

  scan(sheet: WorkPaperSheetDimensionRecord): WorkPaperSheetDimensions {
    let width = 0
    let height = 0
    sheet.grid.forEachCellEntry((_cellIndex: number, row: number, col: number) => {
      height = Math.max(height, row + 1)
      width = Math.max(width, col + 1)
    })
    return { width, height }
  }

  invalidate(sheetId: number): void {
    this.dimensions.delete(sheetId)
  }

  invalidateAll(): void {
    this.dimensions.clear()
    this.spillSheetIds = null
  }

  updateAfterCellMutationRefs(refs: readonly EngineCellMutationRef[]): void {
    if (this.dimensions.size === 0) {
      return
    }
    if (refs.length === 1) {
      const ref = refs[0]
      const mutation = ref?.mutation
      if (ref && mutation) {
        const cached = this.dimensions.get(ref.sheetId)
        if (!cached) {
          return
        }
        if (mutation.kind === 'setCellFormula' && workPaperFormulaMayResizeDynamically(mutation.formula)) {
          this.spillSheetIds = null
          this.invalidate(ref.sheetId)
          return
        }
        const noKnownSpills = this.spillSheetIds !== null && !this.spillSheetIds.has(ref.sheetId)
        if (
          noKnownSpills &&
          (mutation.kind === 'setCellValue'
            ? mutation.row + 1 <= cached.height && mutation.col + 1 <= cached.width
            : mutation.row + 1 < cached.height && mutation.col + 1 < cached.width)
        ) {
          return
        }
      }
    }
    for (let index = 0; index < refs.length; index += 1) {
      const ref = refs[index]
      if (!ref) {
        continue
      }
      const mutation = ref.mutation
      if (mutation.kind === 'setCellFormula' && workPaperFormulaMayResizeDynamically(mutation.formula)) {
        this.spillSheetIds = null
        this.invalidate(ref.sheetId)
        continue
      }
      if (this.sheetHasSpills(ref.sheetId)) {
        this.invalidate(ref.sheetId)
        continue
      }
      if (mutation.kind === 'clearCell') {
        this.invalidateIfEdge(ref.sheetId, mutation.row, mutation.col)
        continue
      }
      this.expand(ref.sheetId, mutation.row, mutation.col)
    }
  }

  updateAfterMatrixMutationImpact(impact: MatrixMutationDimensionImpact): void {
    const cached = this.dimensions.get(impact.sheetId)
    if (!cached) {
      return
    }
    if (impact.hasDynamicFormula || this.sheetHasSpills(impact.sheetId)) {
      this.invalidate(impact.sheetId)
      return
    }
    if (
      (impact.maxClearRow >= 0 && impact.maxClearRow + 1 >= cached.height) ||
      (impact.maxClearCol >= 0 && impact.maxClearCol + 1 >= cached.width)
    ) {
      this.invalidate(impact.sheetId)
      return
    }
    if (impact.maxSetRow >= 0) {
      cached.height = Math.max(cached.height, impact.maxSetRow + 1)
      cached.width = Math.max(cached.width, impact.maxSetCol + 1)
    }
  }

  updateAfterAxisIntervalEdit(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheetId: number,
    start: number,
    amount: number,
  ): void {
    const cached = this.dimensions.get(sheetId)
    if (!cached) {
      return
    }
    if (this.sheetHasSpills(sheetId)) {
      this.invalidate(sheetId)
      return
    }
    const dimension = axis === 'row' ? 'height' : 'width'
    const current = cached[dimension]
    if (mode === 'add') {
      if (start < current) {
        cached[dimension] = current + amount
      }
      return
    }
    if (start >= current) {
      return
    }
    if (start + amount < current) {
      cached[dimension] = current - amount
      return
    }
    this.invalidate(sheetId)
  }

  updateAfterAxisMove(axis: WorkPaperAxisKind, sheetId: number, start: number, count: number, target: number): void {
    const cached = this.dimensions.get(sheetId)
    if (!cached) {
      return
    }
    if (this.sheetHasSpills(sheetId)) {
      this.invalidate(sheetId)
      return
    }
    const current = axis === 'row' ? cached.height : cached.width
    const sourceInsideBounds = start < current
    const targetInsideOrAtBounds = target <= current
    const sourceEndsInsideBounds = start + count <= current
    if (!sourceInsideBounds && target >= current) {
      return
    }
    if (sourceInsideBounds && sourceEndsInsideBounds && targetInsideOrAtBounds) {
      return
    }
    this.invalidate(sheetId)
  }

  private expand(sheetId: number, row: number, col: number): void {
    const cached = this.dimensions.get(sheetId)
    if (!cached) {
      return
    }
    cached.height = Math.max(cached.height, row + 1)
    cached.width = Math.max(cached.width, col + 1)
  }

  private invalidateIfEdge(sheetId: number, row: number, col: number): void {
    const cached = this.dimensions.get(sheetId)
    if (!cached) {
      return
    }
    if (row + 1 >= cached.height || col + 1 >= cached.width) {
      this.invalidate(sheetId)
    }
  }

  private sheetHasSpills(sheetId: number): boolean {
    if (this.spillSheetIds === null) {
      const spillSheetIds = new Set<number>()
      this.engine.workbook.listSpills().forEach((spill) => {
        const spillSheet = this.engine.workbook.getSheet(spill.sheetName)
        if (spillSheet) {
          spillSheetIds.add(spillSheet.id)
        }
      })
      this.spillSheetIds = spillSheetIds
    }
    return this.spillSheetIds.has(sheetId)
  }
}
