import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag, type CellValue } from '@bilig/protocol'
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

function decodeErrorCode(rawCode: number): ErrorCode {
  switch (rawCode) {
    case 1:
      return ErrorCode.Null
    case 2:
      return ErrorCode.Div0
    case 3:
      return ErrorCode.Value
    case 4:
      return ErrorCode.Ref
    case 5:
      return ErrorCode.Name
    case 6:
      return ErrorCode.Num
    case 7:
      return ErrorCode.NA
    case 8:
      return ErrorCode.Blocked
    default:
      throw new Error(`Unexpected error code: ${rawCode}`)
  }
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
    if (tag === ValueTag.Error) {
      return { tag, code: decodeErrorCode(rawValue) }
    }
    if (tag === ValueTag.Empty) {
      return { tag }
    }
    throw new Error(`Unexpected spill tag: ${tag}`)
  })
}

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

describe('wasm kernel ordered statistics dispatch slab', () => {
  it('keeps rank, percentiles, trimmean, and probability stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(120, 1, 6, 11, 32)

    const cellTags = new Uint8Array(120)
    const cellNumbers = new Float64Array(120)
    ;[10, 20, 20, 30].forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })
    ;[2, 4, 4, 4, 5, 5, 7, 9].forEach((value, index) => {
      cellTags[10 + index] = ValueTag.Number
      cellNumbers[10 + index] = value
    })
    ;[79, 85, 78, 85, 50, 81].forEach((value, index) => {
      cellTags[20 + index] = ValueTag.Number
      cellNumbers[20 + index] = value
    })
    ;[60, 80, 90].forEach((value, index) => {
      cellTags[30 + index] = ValueTag.Number
      cellNumbers[30 + index] = value
    })
    ;[1, 2, 3].forEach((value, index) => {
      cellTags[40 + index] = ValueTag.Number
      cellNumbers[40 + index] = value
    })
    ;[0.2, 0.3, 0.5].forEach((value, index) => {
      cellTags[50 + index] = ValueTag.Number
      cellNumbers[50 + index] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(120), new Uint16Array(120))

    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 10, 11, 12, 13, 14, 15, 16, 17, 20, 21, 22, 23, 24, 25, 30, 31, 32, 40, 41, 42, 50, 51, 52]),
      Uint32Array.from([0, 4, 12, 18, 21, 24]),
      Uint32Array.from([4, 8, 6, 3, 3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 8, 6, 3, 3, 3]), Uint32Array.from([1, 1, 1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodePushNumber(1), encodeCall(BuiltinId.RankAvg, 3), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.StdevS, 1), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.VarP, 1), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.Median, 1), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodeCall(BuiltinId.Large, 2), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodeCall(BuiltinId.PercentileInc, 2), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodeCall(BuiltinId.QuartileExc, 2), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.PercentrankInc, 3), encodeRet()],
      [encodePushRange(4), encodePushRange(5), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Prob, 4), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodeCall(BuiltinId.Trimmean, 2), encodeRet()],
      [encodePushRange(2), encodePushRange(3), encodeCall(BuiltinId.Frequency, 2), encodeRet()],
    ])
    const constants = packConstants([[20, 0], [], [], [], [2], [0.75], [3], [7, 4], [2, 3], [0.25], []])
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
        cellIndex(5, 8, width),
        cellIndex(5, 9, width),
        cellIndex(7, 0, width),
      ]),
    )
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
        cellIndex(5, 8, width),
        cellIndex(5, 9, width),
        cellIndex(7, 0, width),
      ]),
    )

    expectNumberCell(kernel, cellIndex(5, 0, width), 2.5, 12)
    expectNumberCell(kernel, cellIndex(5, 1, width), 2.138089935299395, 12)
    expectNumberCell(kernel, cellIndex(5, 2, width), 4, 12)
    expectNumberCell(kernel, cellIndex(5, 3, width), 4.5, 12)
    expectNumberCell(kernel, cellIndex(5, 4, width), 7, 12)
    expectNumberCell(kernel, cellIndex(5, 5, width), 5.5, 12)
    expectNumberCell(kernel, cellIndex(5, 6, width), 6.5, 12)
    expectNumberCell(kernel, cellIndex(5, 7, width), 0.8571, 12)
    expectNumberCell(kernel, cellIndex(5, 8, width), 0.8, 12)
    expectNumberCell(kernel, cellIndex(5, 9, width), 29 / 6, 12)
    expect(readSpillValues(kernel, cellIndex(7, 0, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 0 },
    ])
  })
})
