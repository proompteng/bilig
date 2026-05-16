import { ensureU16, ensureU32 } from './vm-core-helpers'

let outputStringLengths = new Uint32Array(64)
let outputStringOffsets = new Uint32Array(64)
let outputStringData = new Uint16Array(64)
let outputStringCount = 0
let outputStringDataLength = 0

export const OUTPUT_STRING_BASE: f64 = 2147483648.0

export function getOutputStringLengthsPtr(): usize {
  return outputStringLengths.dataStart
}

export function getOutputStringOffsetsPtr(): usize {
  return outputStringOffsets.dataStart
}

export function getOutputStringDataPtr(): usize {
  return outputStringData.dataStart
}

export function getOutputStringCount(): i32 {
  return outputStringCount
}

export function getOutputStringDataLength(): i32 {
  return outputStringDataLength
}

export function outputStringLengthsView(): Uint32Array {
  return outputStringLengths
}

export function outputStringOffsetsView(): Uint32Array {
  return outputStringOffsets
}

export function outputStringDataView(): Uint16Array {
  return outputStringData
}

export function resetOutputStrings(): void {
  outputStringCount = 0
  outputStringDataLength = 0
}

export function allocateOutputString(length: i32): i32 {
  const index = outputStringCount
  outputStringCount += 1
  outputStringLengths = ensureU32(outputStringLengths, outputStringCount)
  outputStringOffsets = ensureU32(outputStringOffsets, outputStringCount)

  outputStringLengths[index] = length
  outputStringOffsets[index] = outputStringDataLength

  outputStringDataLength += length
  outputStringData = ensureU16(outputStringData, outputStringDataLength)

  return index
}

export function writeOutputStringData(index: i32, offset: i32, char: u16): void {
  const dataOffset = outputStringOffsets[index]
  outputStringData[dataOffset + offset] = char
}

export function encodeOutputStringId(index: i32): f64 {
  return OUTPUT_STRING_BASE + <f64>index
}
