import type { EngineEvent } from '@bilig/protocol'
import type { WorkbookDeltaBatchV3, WorkbookDeltaSourceV3 } from '@bilig/worker-transport'
import { DirtyMaskV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { collectSheetViewportImpacts, type SheetViewportImpact, type WorkerEngine, type WorkerSheet } from './worker-runtime-support.js'

export interface WorkbookDeltaSheetIdentityV3 {
  readonly sheetId: number
  readonly sheetOrdinal: number
}

export interface WorkerRuntimeDeltaPublisherBuildInput {
  readonly engine: WorkerEngine
  readonly event: EngineEvent
  readonly source?: WorkbookDeltaSourceV3 | undefined
}

export interface WorkbookDeltaBatchBuildInput extends WorkerRuntimeDeltaPublisherBuildInput {
  readonly allocateSeq: () => number
  readonly resolveSheetIdentity?: ((sheetName: string) => WorkbookDeltaSheetIdentityV3 | null) | undefined
}

const CHANGED_CELL_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.Rect
const INVALIDATED_RANGE_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border
const AXIS_X_DIRTY_MASK = DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect
const AXIS_Y_DIRTY_MASK = DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect
const FULL_SHEET_DIRTY_MASK = INVALIDATED_RANGE_DIRTY_MASK | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze

export class WorkerRuntimeDeltaPublisher {
  private nextSeq = 1

  buildFromEngineEvent(input: WorkerRuntimeDeltaPublisherBuildInput): WorkbookDeltaBatchV3[] {
    return buildWorkbookDeltaBatchesFromEngineEventV3({
      ...input,
      allocateSeq: () => this.nextSeq++,
    })
  }

  reset(): void {
    this.nextSeq = 1
  }
}

export function buildWorkbookDeltaBatchesFromEngineEventV3(input: WorkbookDeltaBatchBuildInput): WorkbookDeltaBatchV3[] {
  const source = input.source ?? 'workerAuthoritative'
  const resolveSheetIdentity = input.resolveSheetIdentity ?? ((sheetName) => resolveDefaultSheetIdentity(input.engine, sheetName))
  if (input.event.invalidation === 'full') {
    return buildFullSheetDeltaBatches(input, source, resolveSheetIdentity)
  }

  const impactsBySheet = collectSheetViewportImpacts(input.engine, input.event)
  if (!impactsBySheet) {
    return []
  }

  return [...impactsBySheet.entries()]
    .toSorted(([leftName], [rightName]) => compareSheetNames(input.engine, leftName, rightName))
    .flatMap(([sheetName, impact]) => {
      const identity = resolveSheetIdentity(sheetName)
      if (!identity) {
        return []
      }
      const seq = input.allocateSeq()
      return [
        createWorkbookDeltaBatch({
          dirty: buildDirtyRangesFromImpact(impact),
          event: input.event,
          identity,
          seq,
          source,
        }),
      ]
    })
}

function buildFullSheetDeltaBatches(
  input: WorkbookDeltaBatchBuildInput,
  source: WorkbookDeltaSourceV3,
  resolveSheetIdentity: (sheetName: string) => WorkbookDeltaSheetIdentityV3 | null,
): WorkbookDeltaBatchV3[] {
  return [...input.engine.workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .flatMap((sheet) => {
      const identity = resolveSheetIdentity(sheet.name)
      if (!identity) {
        return []
      }
      const seq = input.allocateSeq()
      return [
        createWorkbookDeltaBatch({
          dirty: {
            axisX: new Uint32Array(),
            axisY: new Uint32Array(),
            cellRanges: new Uint32Array(),
            sheets: Uint32Array.from([identity.sheetOrdinal, FULL_SHEET_DIRTY_MASK]),
          },
          event: input.event,
          identity,
          seq,
          source,
        }),
      ]
    })
}

function buildDirtyRangesFromImpact(impact: SheetViewportImpact): WorkbookDeltaBatchV3['dirty'] {
  const cellRanges: number[] = []
  const axisX: number[] = []
  const axisY: number[] = []

  impact.changedCells?.positions.forEach((cell) => {
    appendCellRange(cellRanges, cell.row, cell.row, cell.col, cell.col, CHANGED_CELL_DIRTY_MASK)
  })

  impact.invalidatedRanges.forEach((range) => {
    appendCellRange(cellRanges, range.rowStart, range.rowEnd, range.colStart, range.colEnd, INVALIDATED_RANGE_DIRTY_MASK)
  })

  impact.invalidatedColumns.forEach((column) => {
    appendAxisRange(axisX, column.startIndex, column.endIndex, AXIS_X_DIRTY_MASK)
  })

  impact.invalidatedRows.forEach((row) => {
    appendAxisRange(axisY, row.startIndex, row.endIndex, AXIS_Y_DIRTY_MASK)
  })

  return {
    axisX: Uint32Array.from(axisX),
    axisY: Uint32Array.from(axisY),
    cellRanges: Uint32Array.from(cellRanges),
  }
}

function createWorkbookDeltaBatch(input: {
  readonly dirty: WorkbookDeltaBatchV3['dirty']
  readonly event: EngineEvent
  readonly identity: WorkbookDeltaSheetIdentityV3
  readonly seq: number
  readonly source: WorkbookDeltaSourceV3
}): WorkbookDeltaBatchV3 {
  const revision = resolveRevision(input.event.metrics.batchId, input.seq)
  return {
    magic: 'bilig.workbook.delta.v3',
    version: 1,
    seq: input.seq,
    source: input.source,
    sheetId: input.identity.sheetId,
    sheetOrdinal: input.identity.sheetOrdinal,
    valueSeq: revision,
    styleSeq: revision,
    axisSeqX: revision,
    axisSeqY: revision,
    freezeSeq: revision,
    calcSeq: revision,
    dirty: input.dirty,
  }
}

function appendCellRange(ranges: number[], rowStart: number, rowEnd: number, colStart: number, colEnd: number, mask: number): void {
  ranges.push(rowStart, rowEnd, colStart, colEnd, mask)
}

function appendAxisRange(ranges: number[], start: number, end: number, mask: number): void {
  ranges.push(start, end, mask)
}

function resolveRevision(batchId: number, fallback: number): number {
  return Number.isInteger(batchId) && batchId >= 0 ? batchId : fallback
}

function resolveDefaultSheetIdentity(engine: WorkerEngine, sheetName: string): WorkbookDeltaSheetIdentityV3 | null {
  const sheet = engine.workbook.getSheet(sheetName)
  if (!sheet) {
    return null
  }
  const sheetId = readSheetId(sheet) ?? sheet.order
  return {
    sheetId,
    sheetOrdinal: sheet.order,
  }
}

function readSheetId(sheet: WorkerSheet): number | null {
  const id = sheet.id
  return typeof id === 'number' && Number.isInteger(id) && id >= 0 ? id : null
}

function compareSheetNames(engine: WorkerEngine, leftName: string, rightName: string): number {
  const left = engine.workbook.getSheet(leftName)
  const right = engine.workbook.getSheet(rightName)
  if (left && right) {
    return left.order - right.order
  }
  if (left) {
    return -1
  }
  if (right) {
    return 1
  }
  return leftName.localeCompare(rightName)
}
