import { MAX_COLS, MAX_ROWS, MAX_WASM_RANGE_CELLS, type RangeIndex } from "@bilig/protocol";
import type { CellAddress, CellRangeAddress, RangeAddress } from "@bilig/formula";
import { EdgeArena, type EdgeSlice } from "./edge-arena.js";

export interface RangeDescriptor {
  index: RangeIndex;
  sheetId: number;
  kind: RangeAddress["kind"];
  row1: number;
  col1: number;
  row2: number;
  col2: number;
  members: EdgeSlice;
  refCount: number;
  dynamic: boolean;
}

export interface RangeMaterializer {
  ensureCell(sheetId: number, row: number, col: number): number;
  listSheetCells(sheetId: number): Array<{ cellIndex: number; row: number; col: number }>;
}

interface DynamicRangeIndex {
  sheetId: number;
  rangeIndex: number;
}

export interface RegisteredCellRange {
  rangeIndex: number;
  cellRange: CellRangeAddress;
  materialized: boolean;
}

export class RangeRegistry {
  private readonly descriptors: RangeDescriptor[] = [];
  private readonly byKey = new Map<string, RangeIndex>();
  private readonly dynamicBySheet = new Map<number, DynamicRangeIndex[]>();
  private readonly members = new EdgeArena();

  get size(): number {
    return this.descriptors.length;
  }

  reset(): void {
    this.descriptors.length = 0;
    this.byKey.clear();
    this.dynamicBySheet.clear();
    this.members.reset();
  }

  intern(sheetId: number, range: RangeAddress, materializer: RangeMaterializer): RegisteredCellRange {
    const descriptorKey = keyForRange(sheetId, range);
    const existingIndex = this.byKey.get(descriptorKey);
    if (existingIndex !== undefined) {
      const existing = this.descriptors[existingIndex]!;
      existing.refCount += 1;
      return {
        rangeIndex: existingIndex,
        cellRange: toCellRange(range),
        materialized: false
      };
    }

    const cellRange = toCellRange(range);
    const dynamic = range.kind !== "cells";
    const descriptor: RangeDescriptor = {
      index: this.descriptors.length,
      sheetId,
      kind: range.kind,
      row1: cellRange.start.row,
      col1: cellRange.start.col,
      row2: cellRange.end.row,
      col2: cellRange.end.col,
      members: this.members.empty(),
      refCount: 1,
      dynamic
    };

    const memberIndices =
      range.kind === "cells"
        ? materializeBoundedMembers(sheetId, cellRange, materializer)
        : materializeDynamicMembers(sheetId, cellRange, range.kind, materializer);
    descriptor.members = this.members.replace(descriptor.members, memberIndices);
    this.descriptors.push(descriptor);
    this.byKey.set(descriptorKey, descriptor.index);

    if (dynamic) {
      const entries = this.dynamicBySheet.get(sheetId) ?? [];
      entries.push({ sheetId, rangeIndex: descriptor.index });
      this.dynamicBySheet.set(sheetId, entries);
    }

    return {
      rangeIndex: descriptor.index,
      cellRange,
      materialized: true
    };
  }

  release(rangeIndex: RangeIndex): { removed: boolean; members: Uint32Array } {
    const descriptor = this.descriptors[rangeIndex];
    if (!descriptor) {
      return { removed: false, members: new Uint32Array() };
    }

    descriptor.refCount -= 1;
    if (descriptor.refCount > 0) {
      return { removed: false, members: this.members.read(descriptor.members) };
    }

    this.byKey.delete(keyForDescriptor(descriptor));
    const members = this.members.read(descriptor.members);
    this.members.free(descriptor.members);
    descriptor.members = this.members.empty();
    if (descriptor.dynamic) {
      const dynamic = this.dynamicBySheet.get(descriptor.sheetId);
      if (dynamic) {
        this.dynamicBySheet.set(
          descriptor.sheetId,
          dynamic.filter((entry) => entry.rangeIndex !== rangeIndex)
        );
      }
    }
    descriptor.refCount = 0;
    return { removed: true, members };
  }

  getDescriptor(rangeIndex: RangeIndex): RangeDescriptor {
    const descriptor = this.descriptors[rangeIndex];
    if (!descriptor) {
      throw new Error(`Unknown range index: ${rangeIndex}`);
    }
    return descriptor;
  }

  getMembers(rangeIndex: RangeIndex): Uint32Array {
    return this.members.read(this.getDescriptor(rangeIndex).members);
  }

  addDynamicMember(sheetId: number, row: number, col: number, cellIndex: number): RangeIndex[] {
    const entries = this.dynamicBySheet.get(sheetId);
    if (!entries || entries.length === 0) {
      return [];
    }

    const matched: RangeIndex[] = [];
    for (let index = 0; index < entries.length; index += 1) {
      const rangeIndex = entries[index]!.rangeIndex;
      const descriptor = this.descriptors[rangeIndex]!;
      if (!matchesDynamicRange(descriptor, row, col)) {
        continue;
      }
      const nextMembers = this.members.appendUnique(descriptor.members, cellIndex);
      if (nextMembers.ptr !== descriptor.members.ptr || nextMembers.len !== descriptor.members.len) {
        descriptor.members = nextMembers;
        matched.push(rangeIndex);
      }
    }
    return matched;
  }

  expandToCells(rangeIndex: RangeIndex): Uint32Array {
    return this.getMembers(rangeIndex);
  }
}

function materializeBoundedMembers(sheetId: number, range: CellRangeAddress, materializer: RangeMaterializer): Uint32Array {
  const rowCount = range.end.row - range.start.row + 1;
  const colCount = range.end.col - range.start.col + 1;
  const memberCount = rowCount * colCount;
  if (memberCount > MAX_WASM_RANGE_CELLS) {
    throw new Error(`Bounded range exceeds fast-path cap: ${memberCount}`);
  }
  const members = new Uint32Array(memberCount);
  let cursor = 0;
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    for (let col = range.start.col; col <= range.end.col; col += 1) {
      members[cursor] = materializer.ensureCell(sheetId, row, col);
      cursor += 1;
    }
  }
  return members;
}

function materializeDynamicMembers(
  sheetId: number,
  range: CellRangeAddress,
  kind: RangeAddress["kind"],
  materializer: RangeMaterializer
): Uint32Array {
  const matches: number[] = [];
  const cells = materializer.listSheetCells(sheetId);
  for (let index = 0; index < cells.length; index += 1) {
    const cell = cells[index]!;
    if (kind === "rows") {
      if (cell.row >= range.start.row && cell.row <= range.end.row) {
        matches.push(cell.cellIndex);
      }
      continue;
    }
    if (cell.col >= range.start.col && cell.col <= range.end.col) {
      matches.push(cell.cellIndex);
    }
  }
  matches.sort((left, right) => left - right);
  return Uint32Array.from(matches);
}

function matchesDynamicRange(descriptor: RangeDescriptor, row: number, col: number): boolean {
  if (descriptor.kind === "rows") {
    return row >= descriptor.row1 && row <= descriptor.row2;
  }
  if (descriptor.kind === "cols") {
    return col >= descriptor.col1 && col <= descriptor.col2;
  }
  return false;
}

function toCellRange(range: RangeAddress): CellRangeAddress {
  if (range.kind === "cells") {
    return range;
  }
  if (range.kind === "rows") {
    const cellRange: CellRangeAddress = {
      kind: "cells",
      start: toCellLikeAddress(range.sheetName, range.start.row, 0),
      end: toCellLikeAddress(range.sheetName, range.end.row, MAX_COLS - 1)
    };
    if (range.sheetName !== undefined) {
      cellRange.sheetName = range.sheetName;
    }
    return cellRange;
  }
  const cellRange: CellRangeAddress = {
    kind: "cells",
    start: toCellLikeAddress(range.sheetName, 0, range.start.col),
    end: toCellLikeAddress(range.sheetName, MAX_ROWS - 1, range.end.col)
  };
  if (range.sheetName !== undefined) {
    cellRange.sheetName = range.sheetName;
  }
  return cellRange;
}

function toCellLikeAddress(sheetName: string | undefined, row: number, col: number): CellAddress {
  const address: CellAddress = {
    row,
    col,
    text: ""
  };
  if (sheetName !== undefined) {
    address.sheetName = sheetName;
  }
  return address;
}

function keyForRange(sheetId: number, range: RangeAddress): string {
  if (range.kind === "cells") {
    return `cells:${sheetId}:${range.start.row}:${range.start.col}:${range.end.row}:${range.end.col}`;
  }
  if (range.kind === "rows") {
    return `rows:${sheetId}:${range.start.row}:${range.end.row}`;
  }
  return `cols:${sheetId}:${range.start.col}:${range.end.col}`;
}

function keyForDescriptor(descriptor: RangeDescriptor): string {
  if (descriptor.kind === "cells") {
    return `cells:${descriptor.sheetId}:${descriptor.row1}:${descriptor.col1}:${descriptor.row2}:${descriptor.col2}`;
  }
  if (descriptor.kind === "rows") {
    return `rows:${descriptor.sheetId}:${descriptor.row1}:${descriptor.row2}`;
  }
  return `cols:${descriptor.sheetId}:${descriptor.col1}:${descriptor.col2}`;
}
