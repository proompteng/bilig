import { type CellValue, type ErrorCode, ValueTag } from "@bilig/protocol";
import { createKernel, type SpreadsheetKernel } from "@bilig/wasm-kernel";
import type { CellStore } from "./cell-store.js";
import type { StringPool } from "./string-pool.js";

export interface WasmFormulaUploadLayout {
  targets: Uint32Array;
  programs: Uint32Array;
  programOffsets: Uint32Array;
  programLengths: Uint32Array;
  constants: Float64Array;
  constantOffsets: Uint32Array;
  constantLengths: Uint32Array;
}

export interface WasmRangeUploadLayout {
  members: Uint32Array;
  offsets: Uint32Array;
  lengths: Uint32Array;
  rowCounts: Uint32Array;
  colCounts: Uint32Array;
}

export interface WasmSpillResult {
  rows: number;
  cols: number;
  values: CellValue[];
}

const OUTPUT_STRING_BASE = 2147483648;

function decodeSpillValue(
  tag: number,
  rawValue: number,
  strings: StringPool,
  outputStrings: readonly string[],
): CellValue {
  switch (tag) {
    case 1:
      return { tag: ValueTag.Number, value: rawValue };
    case 2:
      return { tag: ValueTag.Boolean, value: rawValue !== 0 };
    case 3: {
      const outputIndex = rawValue >= OUTPUT_STRING_BASE ? rawValue - OUTPUT_STRING_BASE : -1;
      return {
        tag: ValueTag.String,
        value: outputIndex >= 0 ? (outputStrings[outputIndex] ?? "") : strings.get(rawValue),
        stringId: 0,
      };
    }
    case 4:
      return { tag: ValueTag.Error, code: rawValue as ErrorCode };
    default:
      return { tag: ValueTag.Empty };
  }
}

export class WasmKernelFacade {
  private kernel: SpreadsheetKernel | null = null;
  private initPromise: Promise<void> | null = null;
  private uploadedStringPoolSize = 0;

  get ready(): boolean {
    return this.kernel !== null;
  }

  async init(): Promise<void> {
    if (this.initPromise) return this.initPromise;
    this.initPromise = createKernel()
      .then((kernel) => {
        this.kernel = kernel;
        kernel.init(64, 64, 64, 64, 64);
        return undefined;
      })
      .catch(() => {
        this.kernel = null;
        return undefined;
      });
    return this.initPromise;
  }

  ensureCapacity(
    cellCapacity: number,
    formulaCapacity: number,
    constantCapacity: number,
    rangeCapacity = this.kernel?.getRangeCapacity() ?? 64,
    memberCapacity = this.kernel?.getMemberCapacity() ?? 64,
  ): void {
    this.kernel?.ensureCellCapacity(cellCapacity);
    this.kernel?.ensureFormulaCapacity(formulaCapacity);
    this.kernel?.ensureConstantCapacity(constantCapacity);
    this.kernel?.ensureRangeCapacity(rangeCapacity);
    this.kernel?.ensureMemberCapacity(memberCapacity);
  }

  uploadFormulas(layout: WasmFormulaUploadLayout): void {
    if (!this.kernel) return;
    const wasmFormulaCount = layout.targets.length;

    this.ensureCapacity(
      this.kernel.getCellCapacity(),
      Math.max(wasmFormulaCount, 1),
      Math.max(layout.constants.length, 1),
    );
    this.kernel.uploadPrograms(
      layout.programs,
      layout.programOffsets,
      layout.programLengths,
      layout.targets,
    );
    this.kernel.uploadConstants(layout.constants, layout.constantOffsets, layout.constantLengths);
  }

  uploadRanges(layout: WasmRangeUploadLayout): void {
    if (!this.kernel) return;
    const rangeCapacity = Math.max(layout.offsets.length, 1);
    const memberCapacity = Math.max(layout.members.length, 1);

    this.ensureCapacity(
      this.kernel.getCellCapacity(),
      this.kernel.getFormulaCapacity(),
      this.kernel.getConstantCapacity(),
      rangeCapacity,
      memberCapacity,
    );
    this.kernel.uploadRangeMembers(layout.members, layout.offsets, layout.lengths);
    this.kernel.uploadRangeShapes(layout.rowCounts, layout.colCounts);
  }

  uploadVolatileNowSerial(nowSerial: number): void {
    this.kernel?.uploadVolatileNowSerial(nowSerial);
  }

  uploadVolatileRandomValues(values: Float64Array): void {
    this.kernel?.uploadVolatileRandomValues(values);
  }

  syncStringPool(layout: { offsets: Uint32Array; lengths: Uint32Array; data: Uint16Array }): void {
    if (!this.kernel) return;
    if (layout.lengths.length === this.uploadedStringPoolSize) {
      return;
    }
    this.kernel.uploadStrings(layout.offsets, layout.lengths, layout.data);
    this.uploadedStringPoolSize = layout.lengths.length;
  }

  syncFromStore(store: CellStore, changedCellIndices?: readonly number[] | Uint32Array): void {
    if (!this.kernel) return;
    this.ensureCapacity(
      store.size,
      this.kernel.getFormulaCapacity(),
      this.kernel.getConstantCapacity(),
      this.kernel.getRangeCapacity(),
      this.kernel.getMemberCapacity(),
    );
    if (changedCellIndices === undefined) {
      this.kernel.writeCells(
        store.tags.slice(0, store.size),
        store.numbers.slice(0, store.size),
        store.stringIds.slice(0, store.size),
        store.errors.slice(0, store.size),
      );
      return;
    }
    if (changedCellIndices.length === 0) {
      return;
    }

    const tags = this.kernel.readTags();
    const numbers = this.kernel.readNumbers();
    const stringIds = this.kernel.readStringIds();
    const errors = this.kernel.readErrors();
    for (let index = 0; index < changedCellIndices.length; index += 1) {
      const cellIndex = changedCellIndices[index]!;
      if (cellIndex >= store.size) {
        continue;
      }
      tags[cellIndex] = store.tags[cellIndex]!;
      numbers[cellIndex] = store.numbers[cellIndex]!;
      stringIds[cellIndex] = store.stringIds[cellIndex]!;
      errors[cellIndex] = store.errors[cellIndex]!;
    }
  }

  evalBatch(cellIndices: Uint32Array): void {
    this.kernel?.evalBatch(cellIndices);
  }

  readSpill(cellIndex: number, strings: StringPool): WasmSpillResult | undefined {
    if (!this.kernel) {
      return undefined;
    }
    const rows = this.kernel.readSpillRows()[cellIndex] ?? 0;
    const cols = this.kernel.readSpillCols()[cellIndex] ?? 0;
    if (rows === 0 || cols === 0) {
      return undefined;
    }
    const offset = this.kernel.readSpillOffsets()[cellIndex] ?? 0;
    const length = this.kernel.readSpillLengths()[cellIndex] ?? 0;
    const spillTags = this.kernel.readSpillTags();
    const spillValues = this.kernel.readSpillNumbers();
    const outputStrings = this.kernel.readOutputStrings();
    return {
      rows,
      cols,
      values: Array.from({ length }, (_, index) =>
        decodeSpillValue(
          spillTags[offset + index] ?? ValueTag.Empty,
          spillValues[offset + index] ?? 0,
          strings,
          outputStrings,
        ),
      ),
    };
  }

  materializePivotTable(
    sourceRangeIndex: number,
    sourceWidth: number,
    groupByColumnIndices: Uint32Array,
    valueColumnIndices: Uint32Array,
    valueAggregations: Uint8Array,
  ):
    | {
        rows: number;
        cols: number;
        tags: Uint8Array;
        numbers: Float64Array;
        stringIds: Uint32Array;
        errors: Uint16Array;
      }
    | undefined {
    return this.kernel?.materializePivotTable(
      sourceRangeIndex,
      sourceWidth,
      groupByColumnIndices,
      valueColumnIndices,
      valueAggregations,
    );
  }

  syncToStore(store: CellStore, changedCellIndices: Uint32Array, strings: StringPool): void {
    if (!this.kernel) return;

    // Read and intern new output strings from WASM before updating cells
    const newStrings = this.kernel.readOutputStrings();
    const outputStringIdMap = new Map<number, number>();
    for (let index = 0; index < newStrings.length; index += 1) {
      const str = newStrings[index]!;
      const internedId = strings.intern(str);
      outputStringIdMap.set((index | 0x80000000) >>> 0, internedId);
    }

    const tags = this.kernel.readTags();
    const numbers = this.kernel.readNumbers();
    const stringIds = this.kernel.readStringIds();
    const errors = this.kernel.readErrors();
    changedCellIndices.forEach((cellIndex) => {
      if (cellIndex >= store.size) return;
      const previousTag = store.tags[cellIndex]!;
      const previousNumber = store.numbers[cellIndex]!;
      const previousStringId = store.stringIds[cellIndex]!;
      const previousError = store.errors[cellIndex]!;
      store.tags[cellIndex] = tags[cellIndex]!;
      store.numbers[cellIndex] = numbers[cellIndex]!;

      let newStringId = stringIds[cellIndex]!;
      if ((newStringId & 0x80000000) !== 0) {
        newStringId = outputStringIdMap.get(newStringId) ?? 0;
      }
      store.stringIds[cellIndex] = newStringId;

      store.errors[cellIndex] = errors[cellIndex]!;
      if (
        previousTag !== store.tags[cellIndex] ||
        previousNumber !== store.numbers[cellIndex] ||
        previousStringId !== store.stringIds[cellIndex] ||
        previousError !== store.errors[cellIndex]
      ) {
        store.versions[cellIndex] = (store.versions[cellIndex] ?? 0) + 1;
      }
    });
  }

  get tags(): Uint8Array {
    return this.kernel?.readTags() ?? new Uint8Array();
  }

  get numbers(): Float64Array {
    return this.kernel?.readNumbers() ?? new Float64Array();
  }

  get stringIds(): Uint32Array {
    return this.kernel?.readStringIds() ?? new Uint32Array();
  }

  get errors(): Uint16Array {
    return this.kernel?.readErrors() ?? new Uint16Array();
  }

  get programOffsets(): Uint32Array {
    return this.kernel?.readProgramOffsets() ?? new Uint32Array();
  }

  get programLengths(): Uint32Array {
    return this.kernel?.readProgramLengths() ?? new Uint32Array();
  }

  get constantOffsets(): Uint32Array {
    return this.kernel?.readConstantOffsets() ?? new Uint32Array();
  }

  get constantLengths(): Uint32Array {
    return this.kernel?.readConstantLengths() ?? new Uint32Array();
  }

  get constants(): Float64Array {
    return this.kernel?.readConstants() ?? new Float64Array();
  }

  get rangeOffsets(): Uint32Array {
    return this.kernel?.readRangeOffsets() ?? new Uint32Array();
  }

  get rangeLengths(): Uint32Array {
    return this.kernel?.readRangeLengths() ?? new Uint32Array();
  }

  get rangeMembers(): Uint32Array {
    return this.kernel?.readRangeMembers() ?? new Uint32Array();
  }
}
