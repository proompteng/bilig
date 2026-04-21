import { addEngineCounter, type EngineCounters } from './perf/engine-counters.js'

export const BLOCK_ROWS = 128
export const BLOCK_COLS = 32

function blockKey(row: number, col: number): number {
  return Math.floor(row / BLOCK_ROWS) * 1_000_000 + Math.floor(col / BLOCK_COLS)
}

export interface SheetGridAxisRemapScope {
  readonly start: number
  readonly end?: number
}

export interface SheetGridLogicalLookup {
  get(row: number, col: number): number | undefined
  forEachCellEntry(fn: (cellIndex: number, row: number, col: number) => void): void
}

function blockIntersectsScope(axis: 'row' | 'column', key: number, scope: SheetGridAxisRemapScope | undefined): boolean {
  if (!scope) {
    return true
  }
  const blockStart = axis === 'row' ? Math.floor(key / 1_000_000) * BLOCK_ROWS : (key % 1_000_000) * BLOCK_COLS
  const blockEnd = blockStart + (axis === 'row' ? BLOCK_ROWS : BLOCK_COLS)
  if (scope.end === undefined) {
    return blockEnd > scope.start
  }
  return blockEnd > scope.start && blockStart < scope.end
}

function localAxisRangeForBlock(
  axis: 'row' | 'column',
  key: number,
  scope: SheetGridAxisRemapScope | undefined,
): { start: number; end: number } {
  const blockSize = axis === 'row' ? BLOCK_ROWS : BLOCK_COLS
  if (!scope) {
    return { start: 0, end: blockSize }
  }
  const blockStart = axis === 'row' ? Math.floor(key / 1_000_000) * BLOCK_ROWS : (key % 1_000_000) * BLOCK_COLS
  const localStart = Math.max(scope.start - blockStart, 0)
  const localEnd = scope.end === undefined ? blockSize : Math.min(scope.end - blockStart, blockSize)
  return {
    start: Math.min(localStart, blockSize),
    end: Math.max(Math.min(localEnd, blockSize), 0),
  }
}

function blockIsEmpty(block: Uint32Array): boolean {
  for (let index = 0; index < block.length; index += 1) {
    if (block[index] !== 0) {
      return false
    }
  }
  return true
}

export class SheetGrid {
  readonly blocks = new Map<number, Uint32Array>()

  constructor(
    private readonly counters?: EngineCounters,
    private readonly logicalLookup?: SheetGridLogicalLookup,
  ) {}

  private setInBlocks(blocks: Map<number, Uint32Array>, row: number, col: number, cellIndex: number): void {
    const key = blockKey(row, col)
    let block = blocks.get(key)
    if (!block) {
      block = new Uint32Array(BLOCK_ROWS * BLOCK_COLS)
      blocks.set(key, block)
    }
    const offset = (row % BLOCK_ROWS) * BLOCK_COLS + (col % BLOCK_COLS)
    block[offset] = cellIndex + 1
  }

  get(row: number, col: number): number {
    if (this.logicalLookup) {
      return this.logicalLookup.get(row, col) ?? -1
    }
    return this.getPhysical(row, col)
  }

  getPhysical(row: number, col: number): number {
    const block = this.blocks.get(blockKey(row, col))
    if (!block) return -1
    const offset = (row % BLOCK_ROWS) * BLOCK_COLS + (col % BLOCK_COLS)
    const value = block[offset]!
    return value === 0 ? -1 : value - 1
  }

  forEachPhysicalRangeEntry(
    rowStart: number,
    colStart: number,
    rowEnd: number,
    colEnd: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ): void {
    const blockRowStart = Math.floor(rowStart / BLOCK_ROWS)
    const blockRowEnd = Math.floor(rowEnd / BLOCK_ROWS)
    const blockColStart = Math.floor(colStart / BLOCK_COLS)
    const blockColEnd = Math.floor(colEnd / BLOCK_COLS)
    for (let blockRow = blockRowStart; blockRow <= blockRowEnd; blockRow += 1) {
      const absoluteBlockRow = blockRow * BLOCK_ROWS
      const localRowStart = Math.max(rowStart - absoluteBlockRow, 0)
      const localRowEnd = Math.min(rowEnd - absoluteBlockRow, BLOCK_ROWS - 1)
      for (let blockCol = blockColStart; blockCol <= blockColEnd; blockCol += 1) {
        const block = this.blocks.get(blockRow * 1_000_000 + blockCol)
        if (!block) {
          continue
        }
        const absoluteBlockCol = blockCol * BLOCK_COLS
        const localColStart = Math.max(colStart - absoluteBlockCol, 0)
        const localColEnd = Math.min(colEnd - absoluteBlockCol, BLOCK_COLS - 1)
        for (let localRow = localRowStart; localRow <= localRowEnd; localRow += 1) {
          const row = absoluteBlockRow + localRow
          const rowOffset = localRow * BLOCK_COLS
          for (let localCol = localColStart; localCol <= localColEnd; localCol += 1) {
            const value = block[rowOffset + localCol]!
            if (value === 0) {
              continue
            }
            fn(value - 1, row, absoluteBlockCol + localCol)
          }
        }
      }
    }
  }

  set(row: number, col: number, cellIndex: number): void {
    this.setInBlocks(this.blocks, row, col, cellIndex)
  }

  clear(row: number, col: number): void {
    const block = this.blocks.get(blockKey(row, col))
    if (!block) return
    const offset = (row % BLOCK_ROWS) * BLOCK_COLS + (col % BLOCK_COLS)
    block[offset] = 0
  }

  forEachInRange(rowStart: number, colStart: number, rowEnd: number, colEnd: number, fn: (cellIndex: number) => void): void {
    if (this.logicalLookup) {
      for (let row = rowStart; row <= rowEnd; row += 1) {
        for (let col = colStart; col <= colEnd; col += 1) {
          const value = this.logicalLookup.get(row, col)
          if (value !== undefined) {
            fn(value)
          }
        }
      }
      return
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        const value = this.get(row, col)
        if (value !== -1) {
          fn(value)
        }
      }
    }
  }

  forEachCell(fn: (cellIndex: number) => void): void {
    this.forEachCellEntry((cellIndex) => {
      fn(cellIndex)
    })
  }

  forEachCellEntry(fn: (cellIndex: number, row: number, col: number) => void): void {
    if (this.logicalLookup) {
      this.logicalLookup.forEachCellEntry(fn)
      return
    }
    this.blocks.forEach((block, key) => {
      const blockRow = Math.floor(key / 1_000_000)
      const blockCol = key % 1_000_000
      for (let offset = 0; offset < block.length; offset += 1) {
        const value = block[offset]!
        if (value === 0) continue
        const localRow = Math.floor(offset / BLOCK_COLS)
        const localCol = offset % BLOCK_COLS
        const row = blockRow * BLOCK_ROWS + localRow
        const col = blockCol * BLOCK_COLS + localCol
        if (row >= 0 && col >= 0) {
          fn(value - 1, row, col)
        }
      }
    })
  }

  collectAxisRemapEntries(
    axis: 'row' | 'column',
    remapIndex: (index: number) => number | undefined,
    scope?: SheetGridAxisRemapScope,
  ): Array<{
    cellIndex: number
    row: number
    col: number
    nextRow: number | undefined
    nextCol: number | undefined
  }> {
    const changedEntries: Array<{
      cellIndex: number
      row: number
      col: number
      nextRow: number | undefined
      nextCol: number | undefined
    }> = []
    ;[...this.blocks.keys()]
      .filter((key) => blockIntersectsScope(axis, key, scope))
      .forEach((key) => {
        if (this.counters) {
          addEngineCounter(this.counters, 'sheetGridBlockScans')
        }
        const block = this.blocks.get(key)
        if (!block) {
          return
        }
        const blockRow = Math.floor(key / 1_000_000)
        const blockCol = key % 1_000_000
        const localAxisRange = localAxisRangeForBlock(axis, key, scope)
        if (localAxisRange.start >= localAxisRange.end) {
          return
        }
        if (axis === 'row') {
          for (let localRow = localAxisRange.start; localRow < localAxisRange.end; localRow += 1) {
            const row = blockRow * BLOCK_ROWS + localRow
            const nextRow = remapIndex(row)
            if (nextRow === row) {
              continue
            }
            const rowOffset = localRow * BLOCK_COLS
            for (let localCol = 0; localCol < BLOCK_COLS; localCol += 1) {
              const offset = rowOffset + localCol
              const value = block[offset]!
              if (value === 0) {
                continue
              }
              const col = blockCol * BLOCK_COLS + localCol
              changedEntries.push({
                cellIndex: value - 1,
                row,
                col,
                nextRow,
                nextCol: col,
              })
            }
          }
          return
        }
        for (let localCol = localAxisRange.start; localCol < localAxisRange.end; localCol += 1) {
          const col = blockCol * BLOCK_COLS + localCol
          const nextCol = remapIndex(col)
          if (nextCol === col) {
            continue
          }
          for (let localRow = 0; localRow < BLOCK_ROWS; localRow += 1) {
            const offset = localRow * BLOCK_COLS + localCol
            const value = block[offset]!
            if (value === 0) {
              continue
            }
            const row = blockRow * BLOCK_ROWS + localRow
            changedEntries.push({
              cellIndex: value - 1,
              row,
              col,
              nextRow: row,
              nextCol,
            })
          }
        }
      })
    return changedEntries
  }

  someCellInAxisScope(
    axis: 'row' | 'column',
    scope: SheetGridAxisRemapScope,
    predicate: (cellIndex: number, row: number, col: number) => boolean,
  ): boolean {
    if (this.logicalLookup) {
      let found = false
      this.logicalLookup.forEachCellEntry((cellIndex, row, col) => {
        if (found) {
          return
        }
        const axisIndex = axis === 'row' ? row : col
        const inScope = scope.end === undefined ? axisIndex >= scope.start : axisIndex >= scope.start && axisIndex < scope.end
        if (inScope && predicate(cellIndex, row, col)) {
          found = true
        }
      })
      return found
    }
    for (const key of this.blocks.keys()) {
      if (!blockIntersectsScope(axis, key, scope)) {
        continue
      }
      const block = this.blocks.get(key)
      if (!block) {
        continue
      }
      const blockRow = Math.floor(key / 1_000_000)
      const blockCol = key % 1_000_000
      const localAxisRange = localAxisRangeForBlock(axis, key, scope)
      if (localAxisRange.start >= localAxisRange.end) {
        continue
      }
      if (axis === 'row') {
        for (let localRow = localAxisRange.start; localRow < localAxisRange.end; localRow += 1) {
          const row = blockRow * BLOCK_ROWS + localRow
          const rowOffset = localRow * BLOCK_COLS
          for (let localCol = 0; localCol < BLOCK_COLS; localCol += 1) {
            const value = block[rowOffset + localCol]!
            if (value === 0) {
              continue
            }
            const col = blockCol * BLOCK_COLS + localCol
            if (predicate(value - 1, row, col)) {
              return true
            }
          }
        }
        continue
      }
      for (let localCol = localAxisRange.start; localCol < localAxisRange.end; localCol += 1) {
        const col = blockCol * BLOCK_COLS + localCol
        for (let localRow = 0; localRow < BLOCK_ROWS; localRow += 1) {
          const value = block[localRow * BLOCK_COLS + localCol]!
          if (value === 0) {
            continue
          }
          const row = blockRow * BLOCK_ROWS + localRow
          if (predicate(value - 1, row, col)) {
            return true
          }
        }
      }
    }
    return false
  }

  remapAxis(
    axis: 'row' | 'column',
    remapIndex: (index: number) => number | undefined,
    scope?: SheetGridAxisRemapScope,
  ): Array<{
    cellIndex: number
    row: number
    col: number
    nextRow: number | undefined
    nextCol: number | undefined
  }> {
    const changedEntries = this.collectAxisRemapEntries(axis, remapIndex, scope)
    const touchedBlockKeys = new Set<number>()
    changedEntries.forEach(({ row, col, nextRow, nextCol }) => {
      touchedBlockKeys.add(blockKey(row, col))
      if (nextRow !== undefined && nextCol !== undefined) {
        touchedBlockKeys.add(blockKey(nextRow, nextCol))
      }
    })
    changedEntries.forEach(({ row, col }) => {
      this.clear(row, col)
    })
    changedEntries.forEach(({ cellIndex, nextRow, nextCol }) => {
      if (nextRow === undefined || nextCol === undefined) {
        return
      }
      this.set(nextRow, nextCol, cellIndex)
    })
    touchedBlockKeys.forEach((key) => {
      const block = this.blocks.get(key)
      if (block && blockIsEmpty(block)) {
        this.blocks.delete(key)
      }
    })
    return changedEntries
  }
}
