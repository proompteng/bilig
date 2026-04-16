import { describe, expect, it } from 'vitest'
import { BuiltinId, Opcode, ValueTag, type CellValue, ErrorCode } from '@bilig/protocol'
import { createKernel, type KernelInstance } from '../index.js'

const OUTPUT_STRING_BASE = 2147483648

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

function decodeValueTag(rawTag: number): ValueTag {
  switch (rawTag) {
    case 0:
      return ValueTag.Empty
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    default:
      throw new Error(`Unexpected raw tag: ${rawTag}`)
  }
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

function readSpillValues(kernel: KernelInstance, ownerCellIndex: number, pooledStrings: readonly string[]): CellValue[] {
  const offset = kernel.readSpillOffsets()[ownerCellIndex] ?? 0
  const length = kernel.readSpillLengths()[ownerCellIndex] ?? 0
  const tags = kernel.readSpillTags()
  const values = kernel.readSpillNumbers()
  return Array.from({ length }, (_, index) => {
    const tag = decodeValueTag(tags[offset + index] ?? ValueTag.Empty)
    const rawValue = values[offset + index] ?? 0
    switch (tag) {
      case ValueTag.Number:
        return { tag, value: rawValue }
      case ValueTag.Boolean:
        return { tag, value: rawValue !== 0 }
      case ValueTag.Empty:
        return { tag }
      case ValueTag.Error:
        return { tag, code: decodeErrorCode(rawValue) }
      case ValueTag.String: {
        const outputIndex = rawValue >= OUTPUT_STRING_BASE ? rawValue - OUTPUT_STRING_BASE : -1
        return {
          tag,
          value: outputIndex >= 0 ? (kernel.readOutputStrings()[outputIndex] ?? '') : (pooledStrings[rawValue] ?? ''),
          stringId: 0,
        }
      }
    }
    throw new Error('Unexpected decoded spill tag')
  })
}

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

describe('wasm kernel statistics and special-function helpers', () => {
  it('keeps distribution and special-function behavior stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(50, 1, 0, 13, 80)
    kernel.writeCells(new Uint8Array(50), new Float64Array(50), new Uint32Array(50), new Uint16Array(50))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besseli, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besselj, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besselk, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bessely, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Betadist, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.BetaInv, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Fdist, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.FInvRt, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.TDist, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.TInv2T, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.GammaInv, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Chisqdist, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Chisqinv, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(1, 8, width),
        cellIndex(1, 9, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
      ]),
    )
    const constants = packConstants([
      [1.5, 1],
      [1.9, 2],
      [1.5, 1],
      [2.5, 1],
      [2, 8, 10, 1, 3],
      [0.6854705810117458, 8, 10, 1, 3],
      [15.2068649, 6, 4],
      [0.01, 6, 4],
      [1, 1, 1],
      [0.5, 1],
      [0.08030139707139418, 3, 2],
      [18.307, 10],
      [0.050001, 10],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(1, 8, width),
        cellIndex(1, 9, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
      ]),
    )

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.981666428, 8)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.329925728, 8)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.277387804, 7)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.145918138, 7)
    expectNumberCell(kernel, cellIndex(1, 4, width), 0.6854705810117458, 9)
    expectNumberCell(kernel, cellIndex(1, 5, width), 2, 9)
    expectNumberCell(kernel, cellIndex(1, 6, width), 0.01, 9)
    expectNumberCell(kernel, cellIndex(1, 7, width), 15.206864870947697, 7)
    expectNumberCell(kernel, cellIndex(1, 8, width), 0.75, 12)
    expectNumberCell(kernel, cellIndex(1, 9, width), 1, 12)
    expectNumberCell(kernel, cellIndex(2, 0, width), 2, 9)
    expectNumberCell(kernel, cellIndex(2, 1, width), 0.05000058909139826, 9)
    expectNumberCell(kernel, cellIndex(2, 2, width), 18.30697345696106, 8)
  })

  it('keeps ranking and histogram helpers stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(60, 1, 5, 5, 32)
    const cellTags = new Uint8Array(60)
    const cellNumbers = new Float64Array(60)
    const percentRankValues = [1, 2, 4, 7, 8, 9, 10, 12]
    const modeSingleValues = [1, 2, 2, 3, 3, 3]
    const modeMultiValues = [1, 2, 2, 3, 3, 4]
    const frequencyValues = [79, 85, 78, 85, 50, 81]
    const bins = [60, 80, 90]

    percentRankValues.forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })
    modeSingleValues.forEach((value, index) => {
      cellTags[10 + index] = ValueTag.Number
      cellNumbers[10 + index] = value
    })
    modeMultiValues.forEach((value, index) => {
      cellTags[20 + index] = ValueTag.Number
      cellNumbers[20 + index] = value
    })
    frequencyValues.forEach((value, index) => {
      cellTags[30 + index] = ValueTag.Number
      cellNumbers[30 + index] = value
    })
    bins.forEach((value, index) => {
      cellTags[40 + index] = ValueTag.Number
      cellNumbers[40 + index] = value
    })

    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(60), new Uint16Array(60))
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 13, 14, 15, 20, 21, 22, 23, 24, 25, 30, 31, 32, 33, 34, 35, 40, 41, 42]),
      Uint32Array.from([0, 8, 14, 20, 26]),
      Uint32Array.from([8, 6, 6, 6, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([8, 6, 6, 6, 3]), Uint32Array.from([1, 1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Percentrank, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.PercentrankExc, 2), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.ModeSngl, 1), encodeRet()],
      [encodePushRange(2), encodeCall(BuiltinId.ModeMult, 1), encodeRet()],
      [encodePushRange(3), encodePushRange(4), encodeCall(BuiltinId.Frequency, 2), encodeRet()],
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
      ]),
    )
    const constants = packConstants([[8], [8], [], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(5, 0, width),
        cellIndex(5, 1, width),
        cellIndex(5, 2, width),
        cellIndex(5, 3, width),
        cellIndex(5, 4, width),
      ]),
    )

    expectNumberCell(kernel, cellIndex(5, 0, width), 0.571, 12)
    expectNumberCell(kernel, cellIndex(5, 1, width), 0.555, 12)
    expectNumberCell(kernel, cellIndex(5, 2, width), 3, 12)
    expect(readSpillValues(kernel, cellIndex(5, 3, width), [])).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
    ])
    expect(readSpillValues(kernel, cellIndex(5, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 0 },
    ])
  })
})
