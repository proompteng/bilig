import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag } from '@bilig/protocol'
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

describe('wasm kernel scalar math dispatch', () => {
  it('keeps rounding and core scalar math dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 1, 8, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([3]), Uint16Array.from([98, 97, 100]))
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Abs, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Round, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.FloorMath, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.CeilingPrecise, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mod, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Int, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.RoundUp, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.RoundDown, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Trunc, 2), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Round, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mround, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [-3.5],
      [1.234, 2],
      [5.5, 2],
      [-5.5, 2],
      [-5.5, 2, 0],
      [-5.5, 2],
      [5, 2],
      [-1.2],
      [-1.23, 1],
      [-1.29, 1],
      [-1.29, 1],
      [],
      [10, 3],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(3.5)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(1.23, 12)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(4)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(-6)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(-2)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBeCloseTo(-1.3, 12)
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBeCloseTo(-1.2, 12)
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBeCloseTo(-1.2, 12)
    expect(kernel.readTags()[cellIndex(1, 11, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 11, width)]).toBe(ErrorCode.Value)
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBe(9)
  })

  it('keeps trigonometric and transcendental dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 0, 8, 1, 1)
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Sin, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Cos, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Tan, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Asin, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Acos, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Atan, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Atan2, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Degrees, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Radians, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Exp, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Ln, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Log10, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Log, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Power, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Sqrt, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Seriessum, 6),
        encodeRet(),
      ],
      [encodePushNumber(0), encodeCall(BuiltinId.Sqrtpi, 1), encodeRet()],
      [encodeCall(BuiltinId.Pi, 0), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Cot, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Csc, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Sech, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Sign, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 22 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [Math.PI / 2],
      [0],
      [Math.PI / 4],
      [1],
      [1],
      [1],
      [1, 1],
      [Math.PI],
      [180],
      [1],
      [Math.E],
      [100],
      [100, 10],
      [2, 3],
      [9],
      [1, 0, 1, 1, 2, 3],
      [4],
      [],
      [Math.PI / 4],
      [0],
      [0],
      [-7],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 22 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBeCloseTo(Math.PI / 2, 12)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBeCloseTo(0, 12)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBeCloseTo(Math.PI / 4, 12)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBeCloseTo(Math.PI / 4, 12)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBeCloseTo(180, 12)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBeCloseTo(Math.PI, 12)
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBeCloseTo(Math.E, 12)
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 11, width)]).toBeCloseTo(2, 12)
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBeCloseTo(2, 12)
    expect(kernel.readNumbers()[cellIndex(1, 13, width)]).toBeCloseTo(8, 12)
    expect(kernel.readNumbers()[cellIndex(1, 14, width)]).toBeCloseTo(3, 12)
    expect(kernel.readNumbers()[cellIndex(1, 15, width)]).toBeCloseTo(6, 12)
    expect(kernel.readNumbers()[cellIndex(1, 16, width)]).toBeCloseTo(2 * Math.sqrt(Math.PI), 12)
    expect(kernel.readNumbers()[cellIndex(1, 17, width)]).toBeCloseTo(Math.PI, 12)
    expect(kernel.readNumbers()[cellIndex(1, 18, width)]).toBeCloseTo(1, 12)
    expect(kernel.readTags()[cellIndex(1, 19, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 19, width)]).toBe(ErrorCode.Div0)
    expect(kernel.readNumbers()[cellIndex(1, 20, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 21, width)]).toBe(-1)
  })

  it('keeps bessel and combinatorics dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 0, 8, 1, 1)
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besseli, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besselj, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besselk, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bessely, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Fact, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Factdouble, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combin, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combina, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Quotient, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permut, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permutationa, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Even, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Odd, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[1, 0], [1, 0], [1, 0], [1, 0], [5], [6], [5, 2], [2, 3], [7, 3], [4, 2], [3, 2], [-3.2], [-3.2]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBeCloseTo(1.2660658777520084, 6)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(0.7651976865579666, 6)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBeCloseTo(0.42102443824070834, 6)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBeCloseTo(0.088256964215677, 6)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(120)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(48)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(10)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(4)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBe(12)
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBe(9)
    expect(kernel.readNumbers()[cellIndex(1, 11, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBe(-5)
  })
})
