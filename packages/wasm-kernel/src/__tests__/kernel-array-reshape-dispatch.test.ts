import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushRange(rangeIndex: number): number {
  return (Opcode.PushRange << 24) | rangeIndex
}

function encodeRet(): number {
  return Opcode.Ret << 24
}

function packPrograms(programs: number[][]): {
  programs: Uint32Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0

  for (const program of programs) {
    offsets.push(offset)
    lengths.push(program.length)
    flat.push(...program)
    offset += program.length
  }

  return {
    programs: Uint32Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

function readSpillNumbers(kernel: Awaited<ReturnType<typeof createKernel>>, index: number): number[] {
  const offset = kernel.readSpillOffsets()[index] ?? 0
  const length = kernel.readSpillLengths()[index] ?? 0
  return Array.from(kernel.readSpillNumbers().slice(offset, offset + length))
}

describe('wasm kernel array reshape dispatch', () => {
  it('keeps tocol, torow, wraprows, and wrapcols spill shapes stable', async () => {
    const kernel = await createKernel()
    kernel.init(16, 8, 4, 1, 1)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 3, 4, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(16),
      new Uint16Array(16),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5]), Uint32Array.from([0]), Uint32Array.from([6]))
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([3]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BuiltinId.Tocol, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Torow, 1), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Wraprows, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Wrapcols, 2), encodeRet()],
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([8, 9, 10, 11]))
    kernel.uploadConstants(new Float64Array([2]), new Uint32Array([0, 0, 0, 0]), new Uint32Array([1, 1, 1, 1]))
    kernel.evalBatch(Uint32Array.from([8, 9, 10, 11]))

    expect(kernel.readTags()[8]).toBe(ValueTag.Number)
    expect(kernel.readSpillRows()[8]).toBe(6)
    expect(kernel.readSpillCols()[8]).toBe(1)
    expect(readSpillNumbers(kernel, 8)).toEqual([1, 4, 2, 5, 3, 6])

    expect(kernel.readTags()[9]).toBe(ValueTag.Number)
    expect(kernel.readSpillRows()[9]).toBe(1)
    expect(kernel.readSpillCols()[9]).toBe(6)
    expect(readSpillNumbers(kernel, 9)).toEqual([1, 2, 3, 4, 5, 6])

    expect(kernel.readTags()[10]).toBe(ValueTag.Number)
    expect(kernel.readSpillRows()[10]).toBe(3)
    expect(kernel.readSpillCols()[10]).toBe(2)
    expect(readSpillNumbers(kernel, 10)).toEqual([1, 2, 3, 4, 5, 6])

    expect(kernel.readTags()[11]).toBe(ValueTag.Number)
    expect(kernel.readSpillRows()[11]).toBe(2)
    expect(kernel.readSpillCols()[11]).toBe(3)
    expect(readSpillNumbers(kernel, 11)).toEqual([1, 2, 3, 4, 5, 6])
  })

  it('preserves reshape argument validation errors', async () => {
    const kernel = await createKernel()
    kernel.init(12, 4, 2, 1, 1)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([1, 2, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(12),
      new Uint16Array(12),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3]), Uint32Array.from([0]), Uint32Array.from([4]))
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([2]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Tocol, 3), encodeRet()],
      [encodePushRange(0), encodePushNumber(1), encodeCall(BuiltinId.Wraprows, 2), encodeRet()],
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([6, 7]))
    kernel.uploadConstants(new Float64Array([2, 0]), new Uint32Array([0, 0]), new Uint32Array([1, 1]))
    kernel.evalBatch(Uint32Array.from([6, 7]))

    expect(kernel.readTags()[6]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[6]).toBe(ErrorCode.Value)
    expect(kernel.readTags()[7]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[7]).toBe(ErrorCode.Value)
  })
})
