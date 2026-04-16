import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

const OUTPUT_STRING_BASE = 2147483648

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId
}

function encodePushError(code: ErrorCode): number {
  return (Opcode.PushError << 24) | code
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

function readStringCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, pooledStrings: readonly string[]): string {
  expect(kernel.readTags()[index]).toBe(ValueTag.String)
  const raw = kernel.readStringIds()[index] ?? 0
  const outputIndex = raw >= OUTPUT_STRING_BASE ? raw - OUTPUT_STRING_BASE : -1
  return outputIndex >= 0 ? (kernel.readOutputStrings()[outputIndex] ?? '') : (pooledStrings[raw] ?? '')
}

describe('wasm kernel text mutation dispatch', () => {
  it('keeps replace and substitute builtin dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 16
    const pooledStrings = ['abcdef', 'Z', 'banana', 'na', 'X', 'ab']
    kernel.init(48, pooledStrings.length, 7, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 6, 7, 13, 15, 16]),
      Uint32Array.from([6, 1, 6, 2, 1, 2]),
      Uint16Array.from(Array.from('abcdefZbanananaXab', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(48), new Float64Array(48), new Uint32Array(48), new Uint16Array(48))

    const packed = packPrograms([
      [encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodePushString(1), encodeCall(BuiltinId.Replace, 4), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodePushString(1), encodeCall(BuiltinId.Replaceb, 4), encodeRet()],
      [encodePushString(2), encodePushString(3), encodePushString(4), encodeCall(BuiltinId.Substitute, 3), encodeRet()],
      [
        encodePushString(2),
        encodePushString(3),
        encodePushString(4),
        encodePushNumber(0),
        encodeCall(BuiltinId.Substitute, 4),
        encodeRet(),
      ],
      [encodePushString(5), encodePushNumber(0), encodeCall(BuiltinId.Rept, 2), encodeRet()],
      [
        encodePushError(ErrorCode.Ref),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushString(1),
        encodeCall(BuiltinId.Replace, 4),
        encodeRet(),
      ],
      [
        encodePushError(ErrorCode.Ref),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushString(1),
        encodeCall(BuiltinId.Replaceb, 4),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 7 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[2, 3], [3, 2], [], [2], [3], [2, 3], [3, 2]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 7 }, (_, index) => cellIndex(1, index, width))))

    expect(readStringCell(kernel, cellIndex(1, 0, width), pooledStrings)).toBe('aZef')
    expect(readStringCell(kernel, cellIndex(1, 1, width), pooledStrings)).toBe('abZef')
    expect(readStringCell(kernel, cellIndex(1, 2, width), pooledStrings)).toBe('baXX')
    expect(readStringCell(kernel, cellIndex(1, 3, width), pooledStrings)).toBe('banaX')
    expect(readStringCell(kernel, cellIndex(1, 4, width), pooledStrings)).toBe('ababab')
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 5, width)]).toBe(ErrorCode.Ref)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 6, width)]).toBe(ErrorCode.Ref)
  })
})
