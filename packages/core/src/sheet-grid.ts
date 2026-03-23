const BLOCK_ROWS = 128;
const BLOCK_COLS = 32;

function blockKey(row: number, col: number): number {
  return Math.floor(row / BLOCK_ROWS) * 1_000_000 + Math.floor(col / BLOCK_COLS);
}

export class SheetGrid {
  readonly blocks = new Map<number, Uint32Array>();

  get(row: number, col: number): number {
    const block = this.blocks.get(blockKey(row, col));
    if (!block) return -1;
    const offset = (row % BLOCK_ROWS) * BLOCK_COLS + (col % BLOCK_COLS);
    const value = block[offset]!;
    return value === 0 ? -1 : value - 1;
  }

  set(row: number, col: number, cellIndex: number): void {
    const key = blockKey(row, col);
    let block = this.blocks.get(key);
    if (!block) {
      block = new Uint32Array(BLOCK_ROWS * BLOCK_COLS);
      this.blocks.set(key, block);
    }
    const offset = (row % BLOCK_ROWS) * BLOCK_COLS + (col % BLOCK_COLS);
    block[offset] = cellIndex + 1;
  }

  clear(row: number, col: number): void {
    const block = this.blocks.get(blockKey(row, col));
    if (!block) return;
    const offset = (row % BLOCK_ROWS) * BLOCK_COLS + (col % BLOCK_COLS);
    block[offset] = 0;
  }

  forEachInRange(
    rowStart: number,
    colStart: number,
    rowEnd: number,
    colEnd: number,
    fn: (cellIndex: number) => void,
  ): void {
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        const value = this.get(row, col);
        if (value !== -1) {
          fn(value);
        }
      }
    }
  }

  forEachCell(fn: (cellIndex: number) => void): void {
    this.forEachCellEntry((cellIndex) => {
      fn(cellIndex);
    });
  }

  forEachCellEntry(fn: (cellIndex: number, row: number, col: number) => void): void {
    this.blocks.forEach((block, key) => {
      const blockRow = Math.floor(key / 1_000_000);
      const blockCol = key % 1_000_000;
      for (let offset = 0; offset < block.length; offset += 1) {
        const value = block[offset]!;
        if (value === 0) continue;
        const localRow = Math.floor(offset / BLOCK_COLS);
        const localCol = offset % BLOCK_COLS;
        const row = blockRow * BLOCK_ROWS + localRow;
        const col = blockCol * BLOCK_COLS + localCol;
        if (row >= 0 && col >= 0) {
          fn(value - 1, row, col);
        }
      }
    });
  }
}
