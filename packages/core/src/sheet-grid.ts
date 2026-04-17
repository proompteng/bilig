const BLOCK_ROWS = 128
const BLOCK_COLS = 32

function blockKey(row: number, col: number): number {
  return Math.floor(row / BLOCK_ROWS) * 1_000_000 + Math.floor(col / BLOCK_COLS)
}

export interface SheetGridAxisRemapScope {
  readonly start: number;
  readonly end?: number;
}

function blockIntersectsScope(
  axis: "row" | "column",
  key: number,
  scope: SheetGridAxisRemapScope | undefined,
): boolean {
  if (!scope) {
    return true;
  }
  const blockStart =
    axis === "row" ? Math.floor(key / 1_000_000) * BLOCK_ROWS : (key % 1_000_000) * BLOCK_COLS;
  const blockEnd = blockStart + (axis === "row" ? BLOCK_ROWS : BLOCK_COLS);
  if (scope.end === undefined) {
    return blockEnd > scope.start;
  }
  return blockEnd > scope.start && blockStart < scope.end;
}

function axisIndexInScope(index: number, scope: SheetGridAxisRemapScope | undefined): boolean {
  if (!scope) {
    return true;
  }
  if (index < scope.start) {
    return false;
  }
  return scope.end === undefined || index < scope.end;
}

function blockIsEmpty(block: Uint32Array): boolean {
  for (let index = 0; index < block.length; index += 1) {
    if (block[index] !== 0) {
      return false;
    }
  }
  return true;
}

export class SheetGrid {
  readonly blocks = new Map<number, Uint32Array>()

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
    const block = this.blocks.get(blockKey(row, col))
    if (!block) return -1
    const offset = (row % BLOCK_ROWS) * BLOCK_COLS + (col % BLOCK_COLS)
    const value = block[offset]!
    return value === 0 ? -1 : value - 1
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
    const changedEntries: Array<{
      cellIndex: number;
      row: number;
      col: number;
      nextRow: number | undefined;
      nextCol: number | undefined;
    }> = [];
    const touchedBlockKeys = new Set<number>();
    [...this.blocks.keys()]
      .filter((key) => blockIntersectsScope(axis, key, scope))
      .forEach((key) => {
        const block = this.blocks.get(key);
        if (!block) {
          return;
        }
        const blockRow = Math.floor(key / 1_000_000);
        const blockCol = key % 1_000_000;
        for (let offset = 0; offset < block.length; offset += 1) {
          const value = block[offset]!;
          if (value === 0) {
            continue;
          }
          const localRow = Math.floor(offset / BLOCK_COLS);
          const localCol = offset % BLOCK_COLS;
          const row = blockRow * BLOCK_ROWS + localRow;
          const col = blockCol * BLOCK_COLS + localCol;
          const axisIndex = axis === "row" ? row : col;
          if (!axisIndexInScope(axisIndex, scope)) {
            continue;
          }
          const nextRow = axis === "row" ? remapIndex(row) : row;
          const nextCol = axis === "column" ? remapIndex(col) : col;
          if (nextRow === row && nextCol === col) {
            continue;
          }
          changedEntries.push({
            cellIndex: value - 1,
            row,
            col,
            nextRow,
            nextCol,
          });
          touchedBlockKeys.add(key);
          if (nextRow !== undefined && nextCol !== undefined) {
            touchedBlockKeys.add(blockKey(nextRow, nextCol));
          }
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
