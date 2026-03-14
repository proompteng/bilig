import { FormulaMode } from "@bilig/protocol";
import { createKernel, type SpreadsheetKernel } from "@bilig/wasm-kernel";
import type { CellStore } from "./cell-store.js";

export interface WasmFormulaBinding {
  cellIndex: number;
  program: Uint32Array;
  constants: number[];
  mode: FormulaMode;
}

export interface WasmRangeBinding {
  rangeIndex: number;
  members: Uint32Array;
}

export class WasmKernelFacade {
  private kernel: SpreadsheetKernel | null = null;
  private initPromise: Promise<void> | null = null;

  get ready(): boolean {
    return this.kernel !== null;
  }

  async init(): Promise<void> {
    if (this.initPromise) return this.initPromise;
    this.initPromise = createKernel()
      .then((kernel) => {
        this.kernel = kernel;
        kernel.init(64, 64, 64, 64, 64);
      })
      .catch(() => {
        this.kernel = null;
      });
    return this.initPromise;
  }

  ensureCapacity(
    cellCapacity: number,
    formulaCapacity: number,
    constantCapacity: number,
    rangeCapacity = this.kernel?.getRangeCapacity() ?? 64,
    memberCapacity = this.kernel?.getMemberCapacity() ?? 64
  ): void {
    this.kernel?.ensureCellCapacity(cellCapacity);
    this.kernel?.ensureFormulaCapacity(formulaCapacity);
    this.kernel?.ensureConstantCapacity(constantCapacity);
    this.kernel?.ensureRangeCapacity(rangeCapacity);
    this.kernel?.ensureMemberCapacity(memberCapacity);
  }

  uploadFormulas(formulas: WasmFormulaBinding[]): void {
    if (!this.kernel) return;
    const wasmFormulas = formulas.filter((formula) => formula.mode === FormulaMode.WasmFastPath);
    const programs: number[] = [];
    const constants: number[] = [];
    const programOffsets = new Uint32Array(wasmFormulas.length);
    const programLengths = new Uint32Array(wasmFormulas.length);
    const constantOffsets = new Uint32Array(wasmFormulas.length);
    const constantLengths = new Uint32Array(wasmFormulas.length);
    const targets = new Uint32Array(wasmFormulas.length);

    let programCursor = 0;
    let constantCursor = 0;
    wasmFormulas.forEach((formula, index) => {
      programOffsets[index] = programCursor;
      programLengths[index] = formula.program.length;
      constantOffsets[index] = constantCursor;
      constantLengths[index] = formula.constants.length;
      targets[index] = formula.cellIndex;
      programs.push(...formula.program);
      constants.push(...formula.constants);
      programCursor += formula.program.length;
      constantCursor += formula.constants.length;
    });

    this.ensureCapacity(
      this.kernel.getCellCapacity(),
      Math.max(wasmFormulas.length, 1),
      Math.max(constants.length, 1)
    );
    this.kernel.uploadPrograms(Uint32Array.from(programs), programOffsets, programLengths, targets);
    this.kernel.uploadConstants(Float64Array.from(constants), constantOffsets, constantLengths);
  }

  uploadRanges(ranges: WasmRangeBinding[]): void {
    if (!this.kernel) return;
    const rangeCapacity = Math.max(ranges.length, 1);
    const memberOffsets = new Uint32Array(rangeCapacity);
    const memberLengths = new Uint32Array(rangeCapacity);
    const members: number[] = [];
    let cursor = 0;

    ranges.forEach((range) => {
      memberOffsets[range.rangeIndex] = cursor;
      memberLengths[range.rangeIndex] = range.members.length;
      members.push(...range.members);
      cursor += range.members.length;
    });

    this.ensureCapacity(
      this.kernel.getCellCapacity(),
      this.kernel.getFormulaCapacity(),
      this.kernel.getConstantCapacity(),
      rangeCapacity,
      Math.max(members.length, 1)
    );
    this.kernel.uploadRangeMembers(Uint32Array.from(members), memberOffsets, memberLengths);
  }

  syncFromStore(store: CellStore, changedCellIndices?: readonly number[] | Uint32Array): void {
    if (!this.kernel) return;
    this.ensureCapacity(
      store.size,
      this.kernel.getFormulaCapacity(),
      this.kernel.getConstantCapacity(),
      this.kernel.getRangeCapacity(),
      this.kernel.getMemberCapacity()
    );
    if (changedCellIndices === undefined) {
      this.kernel.writeCells(
        store.tags.slice(0, store.size),
        store.numbers.slice(0, store.size),
        store.stringIds.slice(0, store.size),
        store.errors.slice(0, store.size)
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

  syncToStore(store: CellStore, changedCellIndices: Uint32Array): void {
    if (!this.kernel) return;
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
      store.stringIds[cellIndex] = stringIds[cellIndex]!;
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
}
