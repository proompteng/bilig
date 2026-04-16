import { describe, expect, it } from 'vitest'
import { BuiltinId, Opcode, ValueTag, type CellValue } from '@bilig/protocol'
import { createKernel, type KernelInstance } from '../index.js'

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

function packConstants(constantsByProgram: number[][]): {
  constants: Float64Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0

  for (const constants of constantsByProgram) {
    offsets.push(offset)
    lengths.push(constants.length)
    flat.push(...constants)
    offset += constants.length
  }

  return {
    constants: Float64Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

function readSpillValues(kernel: KernelInstance, ownerCellIndex: number): CellValue[] {
  const offset = kernel.readSpillOffsets()[ownerCellIndex] ?? 0
  const length = kernel.readSpillLengths()[ownerCellIndex] ?? 0
  const tags = kernel.readSpillTags()
  const values = kernel.readSpillNumbers()
  return Array.from({ length }, (_, index) => {
    const tag = tags[offset + index] ?? ValueTag.Empty
    const rawValue = values[offset + index] ?? 0
    if (tag === ValueTag.Number) {
      return { tag, value: rawValue }
    }
    if (tag === ValueTag.Empty) {
      return { tag }
    }
    throw new Error(`Unexpected spill tag: ${tag}`)
  })
}

describe('wasm kernel array materialization helpers', () => {
  it('keeps CHOOSE range materialization and UNIQUE spill behavior stable', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(48, 4, 8, 4, 32)

    const cellTags = new Uint8Array(48)
    const cellNumbers = new Float64Array(48)
    ;[10, 20, 30, 40, 1, 2, 1, 2, 3, 4, 5, 5, 6, 6].forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(48), new Uint16Array(48))

    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]),
      Uint32Array.from([0, 4, 10]),
      Uint32Array.from([4, 6, 4]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([2, 3, 2]), Uint32Array.from([2, 2, 2]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Choose, 3), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.Unique, 1), encodeRet()],
      [encodePushRange(2), encodePushNumber(0), encodeCall(BuiltinId.Unique, 2), encodeRet()],
      [encodePushRange(2), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Unique, 3), encodeRet()],
    ])
    const constants = packConstants([[2], [], [1], [1, 1]])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(4, 0, width), cellIndex(4, 1, width), cellIndex(4, 2, width), cellIndex(4, 3, width)]),
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(4, 0, width), cellIndex(4, 1, width), cellIndex(4, 2, width), cellIndex(4, 3, width)]))

    expect(readSpillValues(kernel, cellIndex(4, 0, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
    ])
    expect(readSpillValues(kernel, cellIndex(4, 1, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
    ])
    expect(readSpillValues(kernel, cellIndex(4, 2, width))).toEqual([
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 6 },
    ])
    expect(readSpillValues(kernel, cellIndex(4, 3, width))).toEqual([])
  })
})
