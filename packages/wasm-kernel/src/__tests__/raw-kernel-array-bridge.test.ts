import { describe, expect, it } from 'vitest'
import { float64ArraySpec, RawKernelArrayBridge, type RawKernelArrayBridgeRuntime, uint8ArraySpec } from '../raw-kernel-array-bridge.js'

describe('raw kernel array bridge', () => {
  it('round-trips typed arrays when allocation grows wasm memory', () => {
    const raw = new FakeRawKernel()
    raw.growOnNextAllocation = true
    const bridge = new RawKernelArrayBridge(raw)
    const source = new Float64Array([1.5, -2, 3.25])
    const target = new Float64Array(source.length)

    const pointer = bridge.lowerTypedArray(source, float64ArraySpec)
    bridge.copyTypedArray(pointer, target, float64ArraySpec)

    expect(Array.from(target)).toEqual(Array.from(source))
    expect(raw.pinned).toEqual([pointer])
    expect(raw.unpinned).toHaveLength(1)
    raw.__unpin(pointer)
    expect(raw.pinned).toEqual([])
  })

  it('releases the backing buffer when header allocation fails', () => {
    const raw = new FakeRawKernel()
    raw.failOnAllocation = 2
    const bridge = new RawKernelArrayBridge(raw)

    expect(() => bridge.lowerTypedArray(new Uint8Array([1, 2]), uint8ArraySpec)).toThrow('allocation failed')
    expect(raw.pinned).toEqual([])
    expect(raw.unpinned).toEqual([16])
  })

  it('releases the backing buffer when header pinning fails', () => {
    const raw = new FakeRawKernel()
    raw.failOnPin = 2
    const bridge = new RawKernelArrayBridge(raw)

    expect(() => bridge.lowerTypedArray(new Uint8Array([1, 2]), uint8ArraySpec)).toThrow('pin failed')
    expect(raw.pinned).toEqual([])
    expect(raw.unpinned).toEqual([16])
  })

  it('releases both pins when header initialization fails', () => {
    const raw = new FakeRawKernel()
    raw.failOnHeaderInitialization = true
    const bridge = new RawKernelArrayBridge(raw)

    expect(() => bridge.lowerTypedArray(new Uint8Array([1, 2]), uint8ArraySpec)).toThrow(RangeError)
    expect(raw.pinned).toEqual([])
    expect(raw.unpinned).toEqual([65532, 16])
  })

  it('releases previously lowered arrays when a later array fails to lower', () => {
    const raw = new FakeRawKernel()
    raw.failOnAllocation = 3
    const bridge = new RawKernelArrayBridge(raw)

    expect(() =>
      bridge.lowerTypedArrays([
        [new Uint8Array([1, 2]), uint8ArraySpec],
        [new Uint8Array([3, 4]), uint8ArraySpec],
      ]),
    ).toThrow('allocation failed')
    expect(raw.pinned).toEqual([])
    expect(raw.unpinned).toEqual([16, 34])
  })
})

class FakeRawKernel implements RawKernelArrayBridgeRuntime {
  readonly memory = new WebAssembly.Memory({ initial: 1, maximum: 2 })
  readonly pinned: number[] = []
  readonly unpinned: number[] = []
  growOnNextAllocation = false
  failOnAllocation: number | undefined
  failOnPin: number | undefined
  failOnHeaderInitialization = false
  private allocationCount = 0
  private pinCount = 0
  private nextPointer = 16

  __new(size: number, _classId: number): number {
    this.allocationCount += 1
    if (this.allocationCount === this.failOnAllocation) {
      throw new Error('allocation failed')
    }

    const pointer = this.failOnHeaderInitialization && this.allocationCount === 2 ? 65532 : this.nextPointer
    this.nextPointer += Math.max(size, 1) + 16
    if (this.growOnNextAllocation) {
      this.memory.grow(1)
      this.growOnNextAllocation = false
    }
    return pointer
  }

  __pin(pointer: number): number {
    this.pinCount += 1
    if (this.pinCount === this.failOnPin) {
      throw new Error('pin failed')
    }

    this.pinned.push(pointer)
    return pointer
  }

  __unpin(pointer: number): void {
    this.unpinned.push(pointer)
    const index = this.pinned.indexOf(pointer)
    if (index >= 0) {
      this.pinned.splice(index, 1)
    }
  }
}
