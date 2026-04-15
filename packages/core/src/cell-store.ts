import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";

export const enum CellFlags {
  HasFormula = 1 << 1,
  JsOnly = 1 << 2,
  InCycle = 1 << 3,
  Materialized = 1 << 4,
  PendingDelete = 1 << 5,
  SpillChild = 1 << 6,
  PivotOutput = 1 << 7,
  AuthoredBlank = 1 << 8,
}

export class CellStore {
  size = 0;
  capacity: number;
  onSetValue: ((index: number) => void) | null = null;
  tags: Uint8Array;
  numbers: Float64Array;
  stringIds: Uint32Array;
  errors: Uint16Array;
  formulaIds: Uint32Array;
  versions: Uint32Array;
  flags: Uint32Array;
  sheetIds: Uint16Array;
  rows: Uint32Array;
  cols: Uint16Array;
  topoRanks: Uint32Array;
  cycleGroupIds: Int32Array;

  constructor(initialCapacity = 64) {
    this.capacity = initialCapacity;
    this.tags = new Uint8Array(initialCapacity);
    this.numbers = new Float64Array(initialCapacity);
    this.stringIds = new Uint32Array(initialCapacity);
    this.errors = new Uint16Array(initialCapacity);
    this.formulaIds = new Uint32Array(initialCapacity);
    this.versions = new Uint32Array(initialCapacity);
    this.flags = new Uint32Array(initialCapacity);
    this.sheetIds = new Uint16Array(initialCapacity);
    this.rows = new Uint32Array(initialCapacity);
    this.cols = new Uint16Array(initialCapacity);
    this.topoRanks = new Uint32Array(initialCapacity);
    this.cycleGroupIds = new Int32Array(initialCapacity);
    this.cycleGroupIds.fill(-1);
  }

  ensureCapacity(nextSize: number): void {
    if (nextSize <= this.capacity) return;
    let nextCapacity = this.capacity;
    while (nextCapacity < nextSize) nextCapacity *= 2;
    this.tags = grow(this.tags, nextCapacity);
    this.numbers = grow(this.numbers, nextCapacity);
    this.stringIds = grow(this.stringIds, nextCapacity);
    this.errors = grow(this.errors, nextCapacity);
    this.formulaIds = grow(this.formulaIds, nextCapacity);
    this.versions = grow(this.versions, nextCapacity);
    this.flags = grow(this.flags, nextCapacity);
    this.sheetIds = grow(this.sheetIds, nextCapacity);
    this.rows = grow(this.rows, nextCapacity);
    this.cols = grow(this.cols, nextCapacity);
    this.topoRanks = grow(this.topoRanks, nextCapacity);
    const nextCycle = grow(this.cycleGroupIds, nextCapacity);
    nextCycle.fill(-1, this.capacity);
    this.cycleGroupIds = nextCycle;
    this.capacity = nextCapacity;
  }

  allocate(sheetId: number, row: number, col: number): number {
    this.ensureCapacity(this.size + 1);
    const index = this.size;
    this.size += 1;
    this.sheetIds[index] = sheetId;
    this.rows[index] = row;
    this.cols[index] = col;
    this.tags[index] = ValueTag.Empty;
    this.errors[index] = ErrorCode.None;
    this.flags[index] = CellFlags.Materialized;
    return index;
  }

  reset(): void {
    this.size = 0;
    this.tags.fill(0);
    this.numbers.fill(0);
    this.stringIds.fill(0);
    this.errors.fill(0);
    this.formulaIds.fill(0);
    this.versions.fill(0);
    this.flags.fill(0);
    this.sheetIds.fill(0);
    this.rows.fill(0);
    this.cols.fill(0);
    this.topoRanks.fill(0);
    this.cycleGroupIds.fill(-1);
  }

  setValue(index: number, value: CellValue, stringId = 0): void {
    this.tags[index] = value.tag;
    this.errors[index] = value.tag === ValueTag.Error ? value.code : ErrorCode.None;
    this.stringIds[index] = value.tag === ValueTag.String ? stringId : 0;
    this.numbers[index] =
      value.tag === ValueTag.Number
        ? value.value
        : value.tag === ValueTag.Boolean
          ? value.value
            ? 1
            : 0
          : 0;
    this.versions[index] = (this.versions[index] ?? 0) + 1;
    this.onSetValue?.(index);
  }

  getValue(index: number, stringLookup: (id: number) => string): CellValue {
    const rawTag = this.tags[index];
    const readValue = rawTag === undefined ? undefined : valueReaders[rawTag];
    if (!readValue) {
      return { tag: ValueTag.Empty };
    }
    return readValue(this, index, stringLookup);
  }
}

const valueReaders: Array<
  ((store: CellStore, index: number, stringLookup: (id: number) => string) => CellValue) | undefined
> = [];

valueReaders[ValueTag.Empty] = () => ({ tag: ValueTag.Empty });
valueReaders[ValueTag.Number] = (store, index) => ({
  tag: ValueTag.Number,
  value: store.numbers[index]!,
});
valueReaders[ValueTag.Boolean] = (store, index) => ({
  tag: ValueTag.Boolean,
  value: store.numbers[index]! !== 0,
});
valueReaders[ValueTag.String] = (store, index, stringLookup) => ({
  tag: ValueTag.String,
  value: stringLookup(store.stringIds[index]!),
  stringId: store.stringIds[index]!,
});
valueReaders[ValueTag.Error] = (store, index) => ({
  tag: ValueTag.Error,
  code: store.errors[index]!,
});

function grow(buffer: Uint8Array, capacity: number): Uint8Array;
function grow(buffer: Uint16Array, capacity: number): Uint16Array;
function grow(buffer: Uint32Array, capacity: number): Uint32Array;
function grow(buffer: Int32Array, capacity: number): Int32Array;
function grow(buffer: Float64Array, capacity: number): Float64Array;
function grow(
  buffer: Uint8Array | Uint16Array | Uint32Array | Int32Array | Float64Array,
  capacity: number,
): Uint8Array | Uint16Array | Uint32Array | Int32Array | Float64Array {
  if (buffer instanceof Uint8Array) {
    const next = new Uint8Array(capacity);
    next.set(buffer);
    return next;
  }
  if (buffer instanceof Uint16Array) {
    const next = new Uint16Array(capacity);
    next.set(buffer);
    return next;
  }
  if (buffer instanceof Uint32Array) {
    const next = new Uint32Array(capacity);
    next.set(buffer);
    return next;
  }
  if (buffer instanceof Int32Array) {
    const next = new Int32Array(capacity);
    next.set(buffer);
    return next;
  }
  const next = new Float64Array(capacity);
  next.set(buffer);
  return next;
}
