import type { RawKernelExports } from './raw-kernel-exports.js'

export type TypedArrayValue = Uint8Array | Uint16Array | Uint32Array | Float64Array

export type RawKernelArrayBridgeRuntime = Pick<RawKernelExports, 'memory' | '__new' | '__pin' | '__unpin'>

export interface LoweredArraySpec<T extends TypedArrayValue> {
  readonly classId: number
  readonly ctor: {
    new (buffer: ArrayBufferLike, byteOffset: number, length: number): T
  }
}

type LoweredArrayInput = readonly [TypedArrayValue, LoweredArraySpec<TypedArrayValue>]
type LoweredArrayPointers<T extends readonly LoweredArrayInput[]> = { [K in keyof T]: number }

const ARRAY_BUFFER_CLASS_ID = 1
const UINT8_ARRAY_CLASS_ID = 4
const FLOAT64_ARRAY_CLASS_ID = 5
const UINT16_ARRAY_CLASS_ID = 6
const UINT32_ARRAY_CLASS_ID = 7

export const uint8ArraySpec: LoweredArraySpec<Uint8Array> = {
  classId: UINT8_ARRAY_CLASS_ID,
  ctor: Uint8Array,
}

export const uint16ArraySpec: LoweredArraySpec<Uint16Array> = {
  classId: UINT16_ARRAY_CLASS_ID,
  ctor: Uint16Array,
}

export const uint32ArraySpec: LoweredArraySpec<Uint32Array> = {
  classId: UINT32_ARRAY_CLASS_ID,
  ctor: Uint32Array,
}

export const float64ArraySpec: LoweredArraySpec<Float64Array> = {
  classId: FLOAT64_ARRAY_CLASS_ID,
  ctor: Float64Array,
}

export class RawKernelArrayBridge {
  private dataView: DataView

  constructor(private readonly raw: RawKernelArrayBridgeRuntime) {
    this.dataView = new DataView(raw.memory.buffer)
  }

  lowerTypedArray<T extends TypedArrayValue>(values: T, spec: LoweredArraySpec<T>): number {
    const byteLength = values.byteLength
    let bufferPtr: number | undefined
    let bufferPinned = false
    let headerPtr: number | undefined
    let headerPinned = false
    let keepHeaderPinned = false

    try {
      const bufferAllocationPtr = this.raw.__new(byteLength, ARRAY_BUFFER_CLASS_ID)
      bufferPtr = this.raw.__pin(bufferAllocationPtr)
      bufferPinned = true

      const headerAllocationPtr = this.raw.__new(12, spec.classId)
      headerPtr = this.raw.__pin(headerAllocationPtr)
      headerPinned = true

      this.setUint32(headerPtr, bufferPtr)
      this.setUint32(headerPtr + 4, bufferPtr)
      this.setUint32(headerPtr + 8, byteLength)
      new spec.ctor(this.raw.memory.buffer, bufferPtr, values.length).set(values)
      keepHeaderPinned = true
      return headerPtr
    } finally {
      if (headerPtr !== undefined && headerPinned && !keepHeaderPinned) {
        this.raw.__unpin(headerPtr)
      }
      if (bufferPtr !== undefined && bufferPinned) {
        this.raw.__unpin(bufferPtr)
      }
    }
  }

  lowerTypedArrays<const T extends readonly LoweredArrayInput[]>(inputs: T): LoweredArrayPointers<T>
  lowerTypedArrays(inputs: readonly LoweredArrayInput[]): readonly number[] {
    const pointers: number[] = []
    try {
      for (const [values, spec] of inputs) {
        pointers.push(this.lowerTypedArray(values, spec))
      }
      return pointers
    } catch (error) {
      for (const pointer of pointers) {
        this.raw.__unpin(pointer)
      }
      throw error
    }
  }

  copyTypedArray<T extends TypedArrayValue>(pointer: number, target: T, spec: LoweredArraySpec<T>): void {
    target.set(new spec.ctor(this.raw.memory.buffer, this.getUint32(pointer + 4), target.length))
  }

  private setUint32(pointer: number, value: number): void {
    try {
      this.dataView.setUint32(pointer, value, true)
    } catch {
      this.dataView = new DataView(this.raw.memory.buffer)
      this.dataView.setUint32(pointer, value, true)
    }
  }

  private getUint32(pointer: number): number {
    try {
      return this.dataView.getUint32(pointer, true)
    } catch {
      this.dataView = new DataView(this.raw.memory.buffer)
      return this.dataView.getUint32(pointer, true)
    }
  }
}
