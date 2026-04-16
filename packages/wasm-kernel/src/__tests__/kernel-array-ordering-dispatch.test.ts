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

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

function readSpillNumbers(kernel: Awaited<ReturnType<typeof createKernel>>, index: number): number[] {
  const offset = kernel.readSpillOffsets()[index] ?? 0
  const length = kernel.readSpillLengths()[index] ?? 0
  return Array.from(kernel.readSpillNumbers().slice(offset, offset + length))
}

describe('wasm kernel array ordering dispatch', () => {
  it('keeps choosecols, chooserows, sort, and sortby spill ordering stable', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(24, 8, 4, 1, 1)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
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
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 3, 4, 5, 6, 4, 3, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]), Uint32Array.from([0, 6]), Uint32Array.from([6, 4]))
    kernel.uploadRangeShapes(Uint32Array.from([2, 4]), Uint32Array.from([3, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Choosecols, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Chooserows, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Sort, 3), encodeRet()],
      [encodePushRange(1), encodePushRange(1), encodeCall(BuiltinId.Sortby, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    kernel.uploadConstants(new Float64Array([2, -1]), Uint32Array.from([0, 0, 0, 0]), Uint32Array.from([2, 2, 2, 2]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readSpillRows()[cellIndex(1, 0, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(1, 0, width)]).toBe(1)
    expect(readSpillNumbers(kernel, cellIndex(1, 0, width))).toEqual([2, 5])

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readSpillRows()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readSpillCols()[cellIndex(1, 1, width)]).toBe(3)
    expect(readSpillNumbers(kernel, cellIndex(1, 1, width))).toEqual([4, 5, 6])

    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readSpillRows()[cellIndex(1, 2, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(1, 2, width)]).toBe(3)
    expect(readSpillNumbers(kernel, cellIndex(1, 2, width))).toEqual([1, 1, 1, 4, 4, 4])

    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readSpillRows()[cellIndex(1, 3, width)]).toBe(4)
    expect(kernel.readSpillCols()[cellIndex(1, 3, width)]).toBe(1)
    expect(readSpillNumbers(kernel, cellIndex(1, 3, width))).toEqual([4, 4, 4, 4])
  })

  it('preserves value errors for invalid array-order arguments', async () => {
    const kernel = await createKernel()
    const width = 6
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
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Choosecols, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Sort, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    )
    kernel.uploadConstants(new Float64Array([3, 0]), Uint32Array.from([0, 0]), Uint32Array.from([2, 2]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 0, width)]).toBe(ErrorCode.Value)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 1, width)]).toBe(ErrorCode.Value)
  })
})
