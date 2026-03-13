type TypedArrayValue = Uint8Array | Uint16Array | Uint32Array | Float64Array;

const ARRAY_BUFFER_CLASS_ID = 1;
const UINT8_ARRAY_CLASS_ID = 4;
const FLOAT64_ARRAY_CLASS_ID = 5;
const UINT16_ARRAY_CLASS_ID = 6;
const UINT32_ARRAY_CLASS_ID = 7;

interface RawKernelExports {
  memory: WebAssembly.Memory;
  __new(size: number, id: number): number;
  __pin(pointer: number): number;
  __unpin(pointer: number): void;
  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number): void;
  ensureCellCapacity(nextCapacity: number): void;
  ensureFormulaCapacity(nextCapacity: number): void;
  ensureConstantCapacity(nextCapacity: number): void;
  uploadPrograms(programs: number, offsets: number, lengths: number, targets: number): void;
  uploadConstants(constants: number, offsets: number, lengths: number): void;
  writeCells(tags: number, numbers: number, errors: number): void;
  evalBatch(cellIndices: number): void;
  getTagsPtr(): number;
  getNumbersPtr(): number;
  getErrorsPtr(): number;
  getCellCapacity(): number;
  getFormulaCapacity(): number;
  getConstantCapacity(): number;
}

interface LoweredArraySpec<T extends TypedArrayValue> {
  align: number;
  classId: number;
  ctor: {
    new (buffer: ArrayBufferLike, byteOffset: number, length: number): T;
  };
}

export interface SpreadsheetKernel {
  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number): void;
  ensureCellCapacity(nextCapacity: number): void;
  ensureFormulaCapacity(nextCapacity: number): void;
  ensureConstantCapacity(nextCapacity: number): void;
  uploadPrograms(
    programs: Uint32Array,
    offsets: Uint32Array,
    lengths: Uint32Array,
    targets: Uint32Array
  ): void;
  uploadConstants(constants: Float64Array, offsets: Uint32Array, lengths: Uint32Array): void;
  writeCells(tags: Uint8Array, numbers: Float64Array, errors: Uint16Array): void;
  evalBatch(cellIndices: Uint32Array): void;
  readTags(): Uint8Array;
  readNumbers(): Float64Array;
  readErrors(): Uint16Array;
  getCellCapacity(): number;
  getFormulaCapacity(): number;
  getConstantCapacity(): number;
}

const uint8Spec: LoweredArraySpec<Uint8Array> = {
  align: 0,
  classId: UINT8_ARRAY_CLASS_ID,
  ctor: Uint8Array
};

const uint16Spec: LoweredArraySpec<Uint16Array> = {
  align: 1,
  classId: UINT16_ARRAY_CLASS_ID,
  ctor: Uint16Array
};

const uint32Spec: LoweredArraySpec<Uint32Array> = {
  align: 2,
  classId: UINT32_ARRAY_CLASS_ID,
  ctor: Uint32Array
};

const float64Spec: LoweredArraySpec<Float64Array> = {
  align: 3,
  classId: FLOAT64_ARRAY_CLASS_ID,
  ctor: Float64Array
};

class RawKernelBridge {
  private dataView: DataView;

  constructor(private readonly raw: RawKernelExports) {
    this.dataView = new DataView(raw.memory.buffer);
  }

  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number): void {
    this.raw.init(cellCapacity, formulaCapacity, constantCapacity);
  }

  ensureCellCapacity(nextCapacity: number): void {
    this.raw.ensureCellCapacity(nextCapacity);
  }

  ensureFormulaCapacity(nextCapacity: number): void {
    this.raw.ensureFormulaCapacity(nextCapacity);
  }

  ensureConstantCapacity(nextCapacity: number): void {
    this.raw.ensureConstantCapacity(nextCapacity);
  }

  uploadPrograms(
    programs: Uint32Array,
    offsets: Uint32Array,
    lengths: Uint32Array,
    targets: Uint32Array
  ): void {
    const programsPtr = this.lowerTypedArray(programs, uint32Spec);
    const offsetsPtr = this.lowerTypedArray(offsets, uint32Spec);
    const lengthsPtr = this.lowerTypedArray(lengths, uint32Spec);
    const targetsPtr = this.lowerTypedArray(targets, uint32Spec);
    try {
      this.raw.uploadPrograms(programsPtr, offsetsPtr, lengthsPtr, targetsPtr);
    } finally {
      this.raw.__unpin(programsPtr);
      this.raw.__unpin(offsetsPtr);
      this.raw.__unpin(lengthsPtr);
      this.raw.__unpin(targetsPtr);
    }
  }

  uploadConstants(constants: Float64Array, offsets: Uint32Array, lengths: Uint32Array): void {
    const constantsPtr = this.lowerTypedArray(constants, float64Spec);
    const offsetsPtr = this.lowerTypedArray(offsets, uint32Spec);
    const lengthsPtr = this.lowerTypedArray(lengths, uint32Spec);
    try {
      this.raw.uploadConstants(constantsPtr, offsetsPtr, lengthsPtr);
    } finally {
      this.raw.__unpin(constantsPtr);
      this.raw.__unpin(offsetsPtr);
      this.raw.__unpin(lengthsPtr);
    }
  }

  writeCells(tags: Uint8Array, numbers: Float64Array, errors: Uint16Array): void {
    const tagsPtr = this.lowerTypedArray(tags, uint8Spec);
    const numbersPtr = this.lowerTypedArray(numbers, float64Spec);
    const errorsPtr = this.lowerTypedArray(errors, uint16Spec);
    try {
      this.raw.writeCells(tagsPtr, numbersPtr, errorsPtr);
    } finally {
      this.raw.__unpin(tagsPtr);
      this.raw.__unpin(numbersPtr);
      this.raw.__unpin(errorsPtr);
    }
  }

  evalBatch(cellIndices: Uint32Array): void {
    const cellIndicesPtr = this.lowerTypedArray(cellIndices, uint32Spec);
    try {
      this.raw.evalBatch(cellIndicesPtr);
    } finally {
      this.raw.__unpin(cellIndicesPtr);
    }
  }

  private lowerTypedArray<T extends TypedArrayValue>(values: T, spec: LoweredArraySpec<T>): number {
    const byteLength = values.length << spec.align;
    const bufferPtr = this.raw.__pin(this.raw.__new(byteLength, ARRAY_BUFFER_CLASS_ID));
    const headerPtr = this.raw.__pin(this.raw.__new(12, spec.classId));
    try {
      this.setUint32(headerPtr, bufferPtr);
      this.setUint32(headerPtr + 4, bufferPtr);
      this.setUint32(headerPtr + 8, byteLength);
      new spec.ctor(this.raw.memory.buffer, bufferPtr, values.length).set(values);
      return headerPtr;
    } finally {
      this.raw.__unpin(bufferPtr);
    }
  }

  private setUint32(pointer: number, value: number): void {
    try {
      this.dataView.setUint32(pointer, value, true);
    } catch {
      this.dataView = new DataView(this.raw.memory.buffer);
      this.dataView.setUint32(pointer, value, true);
    }
  }
}

class KernelHandle implements SpreadsheetKernel {
  private readonly bridge: RawKernelBridge;
  private tags = new Uint8Array();
  private numbers = new Float64Array();
  private errors = new Uint16Array();

  constructor(private readonly raw: RawKernelExports) {
    this.bridge = new RawKernelBridge(raw);
    this.refreshViews();
  }

  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number): void {
    this.bridge.init(cellCapacity, formulaCapacity, constantCapacity);
    this.refreshViews();
  }

  ensureCellCapacity(nextCapacity: number): void {
    this.bridge.ensureCellCapacity(nextCapacity);
    this.refreshViews();
  }

  ensureFormulaCapacity(nextCapacity: number): void {
    this.bridge.ensureFormulaCapacity(nextCapacity);
    this.refreshViews();
  }

  ensureConstantCapacity(nextCapacity: number): void {
    this.bridge.ensureConstantCapacity(nextCapacity);
    this.refreshViews();
  }

  uploadPrograms(
    programs: Uint32Array,
    offsets: Uint32Array,
    lengths: Uint32Array,
    targets: Uint32Array
  ): void {
    this.bridge.uploadPrograms(programs, offsets, lengths, targets);
    this.refreshViews();
  }

  uploadConstants(constants: Float64Array, offsets: Uint32Array, lengths: Uint32Array): void {
    this.bridge.uploadConstants(constants, offsets, lengths);
    this.refreshViews();
  }

  writeCells(tags: Uint8Array, numbers: Float64Array, errors: Uint16Array): void {
    this.bridge.writeCells(tags, numbers, errors);
    this.refreshViews();
  }

  evalBatch(cellIndices: Uint32Array): void {
    this.bridge.evalBatch(cellIndices);
    this.refreshViews();
  }

  readTags(): Uint8Array {
    return this.tags;
  }

  readNumbers(): Float64Array {
    return this.numbers;
  }

  readErrors(): Uint16Array {
    return this.errors;
  }

  getCellCapacity(): number {
    return this.raw.getCellCapacity();
  }

  getFormulaCapacity(): number {
    return this.raw.getFormulaCapacity();
  }

  getConstantCapacity(): number {
    return this.raw.getConstantCapacity();
  }

  private refreshViews(): void {
    const memory = this.raw.memory.buffer;
    this.tags = new Uint8Array(memory, this.raw.getTagsPtr(), this.raw.getCellCapacity());
    this.numbers = new Float64Array(memory, this.raw.getNumbersPtr(), this.raw.getCellCapacity());
    this.errors = new Uint16Array(memory, this.raw.getErrorsPtr(), this.raw.getCellCapacity());
  }
}

function isNodeLike(): boolean {
  return typeof process !== "undefined" && process.versions != null && process.versions.node != null;
}

async function loadWasmModule(): Promise<WebAssembly.WebAssemblyInstantiatedSource> {
  const wasmUrl = new URL("../build/release.wasm", import.meta.url);
  const imports = {
    env: {
      abort(_message: number, _fileName: number, lineNumber: number, columnNumber: number) {
        throw new Error(`AssemblyScript abort at ${lineNumber}:${columnNumber}`);
      }
    }
  };

  if (isNodeLike()) {
    const fsPromises = process.getBuiltinModule("fs/promises") as typeof import("node:fs/promises") | undefined;
    if (!fsPromises) {
      throw new Error("Node fs/promises module is unavailable");
    }
    const { readFile } = fsPromises;
    const bytes = await readFile(wasmUrl);
    return WebAssembly.instantiate(bytes, imports);
  }

  const response = await fetch(wasmUrl);
  if (!response.ok) {
    throw new Error(`Failed to load wasm kernel: ${response.status} ${response.statusText}`);
  }
  const bytes = await response.arrayBuffer();
  return WebAssembly.instantiate(bytes, imports);
}

export async function createKernel(): Promise<SpreadsheetKernel> {
  const { instance } = await loadWasmModule();
  return new KernelHandle(instance.exports as unknown as RawKernelExports);
}
