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
  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number, rangeCapacity: number, memberCapacity: number): void;
  ensureCellCapacity(nextCapacity: number): void;
  ensureFormulaCapacity(nextCapacity: number): void;
  ensureConstantCapacity(nextCapacity: number): void;
  ensureRangeCapacity(nextCapacity: number): void;
  ensureMemberCapacity(nextCapacity: number): void;
  uploadPrograms(programs: number, offsets: number, lengths: number, targets: number): void;
  uploadConstants(constants: number, offsets: number, lengths: number): void;
  uploadRangeMembers(members: number, offsets: number, lengths: number): void;
  uploadStringLengths(lengths: number): void;
  writeCells(tags: number, numbers: number, stringIds: number, errors: number): void;
  evalBatch(cellIndices: number): void;
  getTagsPtr(): number;
  getNumbersPtr(): number;
  getStringIdsPtr(): number;
  getErrorsPtr(): number;
  getProgramOffsetsPtr(): number;
  getProgramLengthsPtr(): number;
  getConstantOffsetsPtr(): number;
  getConstantLengthsPtr(): number;
  getConstantArenaPtr(): number;
  getRangeOffsetsPtr(): number;
  getRangeLengthsPtr(): number;
  getRangeMembersPtr(): number;
  getCellCapacity(): number;
  getFormulaCapacity(): number;
  getConstantCapacity(): number;
  getRangeCapacity(): number;
  getMemberCapacity(): number;
}

interface LoweredArraySpec<T extends TypedArrayValue> {
  align: number;
  classId: number;
  ctor: {
    new (buffer: ArrayBufferLike, byteOffset: number, length: number): T;
  };
}

export interface SpreadsheetKernel {
  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number, rangeCapacity: number, memberCapacity: number): void;
  ensureCellCapacity(nextCapacity: number): void;
  ensureFormulaCapacity(nextCapacity: number): void;
  ensureConstantCapacity(nextCapacity: number): void;
  ensureRangeCapacity(nextCapacity: number): void;
  ensureMemberCapacity(nextCapacity: number): void;
  uploadPrograms(
    programs: Uint32Array,
    offsets: Uint32Array,
    lengths: Uint32Array,
    targets: Uint32Array
  ): void;
  uploadConstants(constants: Float64Array, offsets: Uint32Array, lengths: Uint32Array): void;
  uploadRangeMembers(members: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void;
  uploadStringLengths(lengths: Uint32Array): void;
  writeCells(tags: Uint8Array, numbers: Float64Array, stringIds: Uint32Array, errors: Uint16Array): void;
  evalBatch(cellIndices: Uint32Array): void;
  readTags(): Uint8Array;
  readNumbers(): Float64Array;
  readStringIds(): Uint32Array;
  readErrors(): Uint16Array;
  readProgramOffsets(): Uint32Array;
  readProgramLengths(): Uint32Array;
  readConstantOffsets(): Uint32Array;
  readConstantLengths(): Uint32Array;
  readConstants(): Float64Array;
  readRangeOffsets(): Uint32Array;
  readRangeLengths(): Uint32Array;
  readRangeMembers(): Uint32Array;
  getCellCapacity(): number;
  getFormulaCapacity(): number;
  getConstantCapacity(): number;
  getRangeCapacity(): number;
  getMemberCapacity(): number;
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

  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number, rangeCapacity: number, memberCapacity: number): void {
    this.raw.init(cellCapacity, formulaCapacity, constantCapacity, rangeCapacity, memberCapacity);
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

  ensureRangeCapacity(nextCapacity: number): void {
    this.raw.ensureRangeCapacity(nextCapacity);
  }

  ensureMemberCapacity(nextCapacity: number): void {
    this.raw.ensureMemberCapacity(nextCapacity);
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

  uploadRangeMembers(members: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void {
    const membersPtr = this.lowerTypedArray(members, uint32Spec);
    const offsetsPtr = this.lowerTypedArray(offsets, uint32Spec);
    const lengthsPtr = this.lowerTypedArray(lengths, uint32Spec);
    try {
      this.raw.uploadRangeMembers(membersPtr, offsetsPtr, lengthsPtr);
    } finally {
      this.raw.__unpin(membersPtr);
      this.raw.__unpin(offsetsPtr);
      this.raw.__unpin(lengthsPtr);
    }
  }

  uploadStringLengths(lengths: Uint32Array): void {
    const lengthsPtr = this.lowerTypedArray(lengths, uint32Spec);
    try {
      this.raw.uploadStringLengths(lengthsPtr);
    } finally {
      this.raw.__unpin(lengthsPtr);
    }
  }

  writeCells(tags: Uint8Array, numbers: Float64Array, stringIds: Uint32Array, errors: Uint16Array): void {
    const tagsPtr = this.lowerTypedArray(tags, uint8Spec);
    const numbersPtr = this.lowerTypedArray(numbers, float64Spec);
    const stringIdsPtr = this.lowerTypedArray(stringIds, uint32Spec);
    const errorsPtr = this.lowerTypedArray(errors, uint16Spec);
    try {
      this.raw.writeCells(tagsPtr, numbersPtr, stringIdsPtr, errorsPtr);
    } finally {
      this.raw.__unpin(tagsPtr);
      this.raw.__unpin(numbersPtr);
      this.raw.__unpin(stringIdsPtr);
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
  private stringIds = new Uint32Array();
  private errors = new Uint16Array();
  private programOffsets = new Uint32Array();
  private programLengths = new Uint32Array();
  private constantOffsets = new Uint32Array();
  private constantLengths = new Uint32Array();
  private constants = new Float64Array();
  private rangeOffsets = new Uint32Array();
  private rangeLengths = new Uint32Array();
  private rangeMembers = new Uint32Array();

  constructor(private readonly raw: RawKernelExports) {
    this.bridge = new RawKernelBridge(raw);
    this.refreshViews();
  }

  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number, rangeCapacity: number, memberCapacity: number): void {
    this.bridge.init(cellCapacity, formulaCapacity, constantCapacity, rangeCapacity, memberCapacity);
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

  ensureRangeCapacity(nextCapacity: number): void {
    this.bridge.ensureRangeCapacity(nextCapacity);
    this.refreshViews();
  }

  ensureMemberCapacity(nextCapacity: number): void {
    this.bridge.ensureMemberCapacity(nextCapacity);
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

  uploadRangeMembers(members: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void {
    this.bridge.uploadRangeMembers(members, offsets, lengths);
    this.refreshViews();
  }

  uploadStringLengths(lengths: Uint32Array): void {
    this.bridge.uploadStringLengths(lengths);
    this.refreshViews();
  }

  writeCells(tags: Uint8Array, numbers: Float64Array, stringIds: Uint32Array, errors: Uint16Array): void {
    this.bridge.writeCells(tags, numbers, stringIds, errors);
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

  readStringIds(): Uint32Array {
    return this.stringIds;
  }

  readErrors(): Uint16Array {
    return this.errors;
  }

  readProgramOffsets(): Uint32Array {
    return this.programOffsets;
  }

  readProgramLengths(): Uint32Array {
    return this.programLengths;
  }

  readConstantOffsets(): Uint32Array {
    return this.constantOffsets;
  }

  readConstantLengths(): Uint32Array {
    return this.constantLengths;
  }

  readConstants(): Float64Array {
    return this.constants;
  }

  readRangeOffsets(): Uint32Array {
    return this.rangeOffsets;
  }

  readRangeLengths(): Uint32Array {
    return this.rangeLengths;
  }

  readRangeMembers(): Uint32Array {
    return this.rangeMembers;
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

  getRangeCapacity(): number {
    return this.raw.getRangeCapacity();
  }

  getMemberCapacity(): number {
    return this.raw.getMemberCapacity();
  }

  private refreshViews(): void {
    const memory = this.raw.memory.buffer;
    this.tags = new Uint8Array(memory, this.raw.getTagsPtr(), this.raw.getCellCapacity());
    this.numbers = new Float64Array(memory, this.raw.getNumbersPtr(), this.raw.getCellCapacity());
    this.stringIds = new Uint32Array(memory, this.raw.getStringIdsPtr(), this.raw.getCellCapacity());
    this.errors = new Uint16Array(memory, this.raw.getErrorsPtr(), this.raw.getCellCapacity());
    this.programOffsets = new Uint32Array(memory, this.raw.getProgramOffsetsPtr(), this.raw.getFormulaCapacity());
    this.programLengths = new Uint32Array(memory, this.raw.getProgramLengthsPtr(), this.raw.getFormulaCapacity());
    this.constantOffsets = new Uint32Array(memory, this.raw.getConstantOffsetsPtr(), this.raw.getFormulaCapacity());
    this.constantLengths = new Uint32Array(memory, this.raw.getConstantLengthsPtr(), this.raw.getFormulaCapacity());
    this.constants = new Float64Array(memory, this.raw.getConstantArenaPtr(), this.raw.getConstantCapacity());
    this.rangeOffsets = new Uint32Array(memory, this.raw.getRangeOffsetsPtr(), this.raw.getRangeCapacity());
    this.rangeLengths = new Uint32Array(memory, this.raw.getRangeLengthsPtr(), this.raw.getRangeCapacity());
    this.rangeMembers = new Uint32Array(memory, this.raw.getRangeMembersPtr(), this.raw.getMemberCapacity());
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
