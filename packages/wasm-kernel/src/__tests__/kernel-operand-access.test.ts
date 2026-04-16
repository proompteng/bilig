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

describe('wasm kernel operand access helpers', () => {
  it('keeps scalar, range, and array operand access stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(40, 0, 2, 8, 32)

    const cellTags = new Uint8Array(40)
    const cellNumbers = new Float64Array(40)

    ;[1, 2, 3, 4, 5, 6].forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })

    const trimmedIndexes = [8, 9, 10, 11, 16, 17, 18, 19, 24, 25, 26, 27, 32, 33, 34, 35]
    ;[11, 12, 13, 14].forEach((value, offset) => {
      const index = trimmedIndexes[[5, 6, 9, 10][offset] ?? 0]
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })

    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(40), new Uint16Array(40))
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, ...trimmedIndexes]), Uint32Array.from([0, 6]), Uint32Array.from([6, 16]))
    kernel.uploadRangeShapes(Uint32Array.from([2, 4]), Uint32Array.from([3, 4]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BuiltinId.Rows, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Columns, 1), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Index, 3), encodeRet()],
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(0),
        encodePushNumber(0),
        encodeCall(BuiltinId.Offset, 5),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Take, 3), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(0), encodeCall(BuiltinId.Drop, 3), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Expand, 4), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.Trimrange, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(5, 0, width),
        cellIndex(5, 1, width),
        cellIndex(5, 2, width),
        cellIndex(5, 3, width),
        cellIndex(5, 4, width),
        cellIndex(5, 5, width),
        cellIndex(5, 6, width),
        cellIndex(5, 7, width),
      ]),
    )
    const constants = packConstants([[], [], [2, 2], [1, 1, 1, 1], [1, 2], [1, 1], [3, 4, 9], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(5, 0, width),
        cellIndex(5, 1, width),
        cellIndex(5, 2, width),
        cellIndex(5, 3, width),
        cellIndex(5, 4, width),
        cellIndex(5, 5, width),
        cellIndex(5, 6, width),
        cellIndex(5, 7, width),
      ]),
    )

    expect(kernel.readNumbers()[cellIndex(5, 0, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(5, 1, width)]).toBe(3)
    expect(kernel.readNumbers()[cellIndex(5, 2, width)]).toBe(5)
    expect(kernel.readNumbers()[cellIndex(5, 3, width)]).toBe(5)
    expect(readSpillValues(kernel, cellIndex(5, 4, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
    ])
    expect(readSpillValues(kernel, cellIndex(5, 5, width))).toEqual([
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 6 },
    ])
    expect(readSpillValues(kernel, cellIndex(5, 6, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 9 },
      { tag: ValueTag.Number, value: 4 },
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Number, value: 9 },
      { tag: ValueTag.Number, value: 9 },
      { tag: ValueTag.Number, value: 9 },
      { tag: ValueTag.Number, value: 9 },
      { tag: ValueTag.Number, value: 9 },
    ])
    expect(readSpillValues(kernel, cellIndex(5, 7, width))).toEqual([
      { tag: ValueTag.Number, value: 11 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 13 },
      { tag: ValueTag.Number, value: 14 },
    ])
  })
})
