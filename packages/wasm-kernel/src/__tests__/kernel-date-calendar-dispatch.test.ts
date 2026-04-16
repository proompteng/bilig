import { describe, expect, it } from 'vitest'
import { BuiltinId, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId
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

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

describe('wasm kernel date/calendar dispatch seams', () => {
  it('keeps date, weekday, and workday dispatch behavior stable', async () => {
    const kernel = await createKernel()
    const width = 20
    kernel.init(80, 18, 10, 4, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 8]),
      Uint32Array.from([8, 1]),
      Uint16Array.from(Array.from('13:02:03Y', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(80), new Float64Array(80), new Uint32Array(80), new Uint16Array(80))

    const packed = packPrograms([
      [encodeCall(BuiltinId.IsBlank, 0), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodeCall(BuiltinId.Year, 1),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodeCall(BuiltinId.Month, 1),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodeCall(BuiltinId.Day, 1),
        encodeRet(),
      ],
      [encodePushString(0), encodeCall(BuiltinId.Timevalue, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Time, 3),
        encodeCall(BuiltinId.Hour, 1),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Time, 3),
        encodeCall(BuiltinId.Minute, 1),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Time, 3),
        encodeCall(BuiltinId.Second, 1),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Weekday, 2),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodeCall(BuiltinId.Isoweeknum, 1),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Edate, 2),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Eomonth, 2),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Date, 3),
        encodePushString(1),
        encodeCall(BuiltinId.Datedif, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Date, 3),
        encodeCall(BuiltinId.Days, 2),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(6),
        encodeCall(BuiltinId.Days360, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(6),
        encodeCall(BuiltinId.Yearfrac, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Weeknum, 2),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.WorkdayIntl, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.NetworkdaysIntl, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: packed.offsets.length }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [],
      [2024, 2, 29],
      [2024, 2, 29],
      [2024, 2, 29],
      [],
      [13, 2, 3],
      [13, 2, 3],
      [13, 2, 3],
      [2024, 1, 1, 2],
      [2024, 1, 1],
      [2024, 1, 31, 1],
      [2024, 1, 31, 1],
      [2023, 1, 1, 2024, 1, 1],
      [2024, 1, 10, 2024, 1, 1],
      [2023, 1, 1, 2023, 3, 1, 0],
      [2023, 1, 1, 2023, 3, 1, 0],
      [2024, 1, 1, 2],
      [46094, 1, 7],
      [46094, 46098, 7],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: packed.offsets.length }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(1)
    expectNumberCell(kernel, cellIndex(1, 1, width), 2024)
    expectNumberCell(kernel, cellIndex(1, 2, width), 2)
    expectNumberCell(kernel, cellIndex(1, 3, width), 29)
    expectNumberCell(kernel, cellIndex(1, 4, width), 46923 / 86400, 12)
    expectNumberCell(kernel, cellIndex(1, 5, width), 13)
    expectNumberCell(kernel, cellIndex(1, 6, width), 2)
    expectNumberCell(kernel, cellIndex(1, 7, width), 3)
    expectNumberCell(kernel, cellIndex(1, 8, width), 1)
    expectNumberCell(kernel, cellIndex(1, 9, width), 1)
    expectNumberCell(kernel, cellIndex(1, 10, width), 45351)
    expectNumberCell(kernel, cellIndex(1, 11, width), 45351)
    expectNumberCell(kernel, cellIndex(1, 12, width), 1)
    expectNumberCell(kernel, cellIndex(1, 13, width), 9)
    expectNumberCell(kernel, cellIndex(1, 14, width), 60)
    expectNumberCell(kernel, cellIndex(1, 15, width), 60 / 360, 12)
    expectNumberCell(kernel, cellIndex(1, 16, width), 1)
    expectNumberCell(kernel, cellIndex(1, 17, width), 46097)
    expectNumberCell(kernel, cellIndex(1, 18, width), 3)
  })
})
