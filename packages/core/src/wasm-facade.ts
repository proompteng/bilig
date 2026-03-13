import { FormulaMode, ValueTag } from "@bilig/protocol";
import { createKernel, type SpreadsheetKernel } from "@bilig/wasm-kernel";
import type { CellStore } from "./cell-store.js";

export interface WasmFormulaBinding {
  cellIndex: number;
  program: Uint32Array;
  constants: number[];
  mode: FormulaMode;
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
        kernel.init(64, 64, 64);
      })
      .catch(() => {
        this.kernel = null;
      });
    return this.initPromise;
  }

  ensureCapacity(cellCapacity: number, formulaCapacity: number, constantCapacity: number): void {
    this.kernel?.ensureCellCapacity(cellCapacity);
    this.kernel?.ensureFormulaCapacity(formulaCapacity);
    this.kernel?.ensureConstantCapacity(constantCapacity);
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

    this.ensureCapacity(this.kernel.getCellCapacity(), Math.max(wasmFormulas.length, 1), Math.max(constants.length, 1));
    this.kernel.uploadPrograms(Uint32Array.from(programs), programOffsets, programLengths, targets);
    this.kernel.uploadConstants(Float64Array.from(constants), constantOffsets, constantLengths);
  }

  syncFromStore(store: CellStore): void {
    if (!this.kernel) return;
    this.ensureCapacity(store.size, this.kernel.getFormulaCapacity(), this.kernel.getConstantCapacity());
    this.kernel.writeCells(store.tags.slice(0, store.size), store.numbers.slice(0, store.size), store.errors.slice(0, store.size));
  }

  evalBatch(cellIndices: Uint32Array): void {
    this.kernel?.evalBatch(cellIndices);
  }

  syncToStore(store: CellStore, changedCellIndices: Uint32Array): void {
    if (!this.kernel) return;
    const tags = this.kernel.readTags();
    const numbers = this.kernel.readNumbers();
    const errors = this.kernel.readErrors();
    changedCellIndices.forEach((cellIndex) => {
      if (cellIndex >= store.size) return;
      store.tags[cellIndex] = tags[cellIndex]!;
      store.numbers[cellIndex] = numbers[cellIndex]!;
      store.errors[cellIndex] = errors[cellIndex]!;
      if (tags[cellIndex] === ValueTag.String) {
        store.stringIds[cellIndex] = 0;
      }
    });
  }
}
