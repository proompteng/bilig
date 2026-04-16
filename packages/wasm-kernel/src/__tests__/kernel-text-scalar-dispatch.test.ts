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

describe('wasm kernel scalar text/search dispatch', () => {
  it('keeps scalar text and search builtin dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 16
    const pooledStrings = ['Hello', 'World', 'Alphabet', '  a  b  ', 'ph', 'HELLO', 'abcdef', 'a*c', 'abc', 'x']
    kernel.init(48, pooledStrings.length, 5, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 5, 10, 18, 26, 28, 33, 39, 42, 45]),
      Uint32Array.from([5, 5, 8, 8, 2, 5, 6, 3, 3, 1]),
      Uint16Array.from(Array.from('HelloWorldAlphabet  a  b  phHELLOabcdefa*cabcx', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(48), new Float64Array(48), new Uint32Array(48), new Uint16Array(48))

    const packed = packPrograms([
      [encodePushString(0), encodePushString(1), encodeCall(BuiltinId.Concat, 2), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Len, 1), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Lenb, 1), encodeRet()],
      [encodePushString(0), encodePushString(5), encodeCall(BuiltinId.Exact, 2), encodeRet()],
      [encodePushString(2), encodePushNumber(0), encodeCall(BuiltinId.Left, 2), encodeRet()],
      [encodePushString(2), encodePushNumber(0), encodeCall(BuiltinId.Right, 2), encodeRet()],
      [encodePushString(2), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mid, 3), encodeRet()],
      [encodePushString(6), encodePushNumber(0), encodeCall(BuiltinId.Leftb, 2), encodeRet()],
      [encodePushString(6), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Midb, 3), encodeRet()],
      [encodePushString(6), encodePushNumber(0), encodeCall(BuiltinId.Rightb, 2), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.Trim, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Upper, 1), encodeRet()],
      [encodePushString(5), encodeCall(BuiltinId.Lower, 1), encodeRet()],
      [encodePushString(4), encodePushString(2), encodeCall(BuiltinId.Find, 2), encodeRet()],
      [encodePushString(7), encodePushString(8), encodeCall(BuiltinId.Search, 2), encodeRet()],
      [encodePushString(4), encodePushString(2), encodeCall(BuiltinId.Findb, 2), encodeRet()],
      [encodePushString(7), encodePushString(8), encodeCall(BuiltinId.Searchb, 2), encodeRet()],
      [encodePushError(ErrorCode.Ref), encodePushString(9), encodeCall(BuiltinId.Concat, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 18 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[], [], [], [], [3], [2], [3, 4], [2], [3, 2], [2], [], [], [], [], [], [], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 18 }, (_, index) => cellIndex(1, index, width))))

    expect(readStringCell(kernel, cellIndex(1, 0, width), pooledStrings)).toBe('HelloWorld')
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(5)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(8)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(0)
    expect(readStringCell(kernel, cellIndex(1, 4, width), pooledStrings)).toBe('Alp')
    expect(readStringCell(kernel, cellIndex(1, 5, width), pooledStrings)).toBe('et')
    expect(readStringCell(kernel, cellIndex(1, 6, width), pooledStrings)).toBe('phab')
    expect(readStringCell(kernel, cellIndex(1, 7, width), pooledStrings)).toBe('ab')
    expect(readStringCell(kernel, cellIndex(1, 8, width), pooledStrings)).toBe('cd')
    expect(readStringCell(kernel, cellIndex(1, 9, width), pooledStrings)).toBe('ef')
    expect(readStringCell(kernel, cellIndex(1, 10, width), pooledStrings)).toBe('a b')
    expect(readStringCell(kernel, cellIndex(1, 11, width), pooledStrings)).toBe('HELLO')
    expect(readStringCell(kernel, cellIndex(1, 12, width), pooledStrings)).toBe('hello')
    expect(kernel.readNumbers()[cellIndex(1, 13, width)]).toBe(3)
    expect(kernel.readNumbers()[cellIndex(1, 14, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 15, width)]).toBe(3)
    expect(kernel.readNumbers()[cellIndex(1, 16, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 17, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 17, width)]).toBe(ErrorCode.Ref)
  })
})
