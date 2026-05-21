const previewRowCount = 8
const previewColumnCount = 6
export const maxSpreadsheetColumnCount = 16_384
const minInt8 = -0x80
const maxInt8 = 0x7f
const minInt16 = -0x8000
const maxInt16 = 0x7fff
const minInt32 = -0x80000000
const maxInt32 = 0x7fffffff

export const previewCellCount = previewRowCount * previewColumnCount

export function isPreviewCell(row: number, column: number): boolean {
  return row >= 0 && row < previewRowCount && column >= 0 && column < previewColumnCount
}

export function previewIndex(row: number, column: number): number {
  return isPreviewCell(row, column) ? row * previewColumnCount + column : -1
}

export function encodeCellAddress(row: number, column: number): string {
  let value = column + 1
  let columnName = ''
  while (value > 0) {
    value -= 1
    columnName = String.fromCharCode(65 + (value % 26)) + columnName
    value = Math.floor(value / 26)
  }
  return `${columnName}${String(row + 1)}`
}

export function packArenaCellAddress(row: number, column: number): number {
  return row * maxSpreadsheetColumnCount + column
}

export function canStoreLinearCoordinate(width: number, row: number, column: number): boolean {
  if (!Number.isSafeInteger(width) || width <= 0 || row < 0 || column < 0 || column >= width) {
    return false
  }
  const linearCellIndex = row * width + column
  return Number.isSafeInteger(linearCellIndex) && linearCellIndex >= 0 && linearCellIndex <= 0xffffffff
}

export function canStoreInt32Number(value: number): boolean {
  return Number.isInteger(value) && value >= minInt32 && value <= maxInt32 && !Object.is(value, -0)
}

export function canStoreInt16Number(value: number): boolean {
  return Number.isInteger(value) && value >= minInt16 && value <= maxInt16 && !Object.is(value, -0)
}

export function canStoreInt8Number(value: number): boolean {
  return Number.isInteger(value) && value >= minInt8 && value <= maxInt8 && !Object.is(value, -0)
}

export function binarySearchUint32(values: Uint32Array, target: number): number {
  return binarySearchUint32Prefix(values, values.length, target)
}

export function binarySearchUint32Prefix(values: Uint32Array, length: number, target: number): number {
  let low = 0
  let high = length - 1
  while (low <= high) {
    const mid = (low + high) >>> 1
    const value = values[mid] ?? 0
    if (value === target) {
      return mid
    }
    if (value < target) {
      low = mid + 1
    } else {
      high = mid - 1
    }
  }
  return -1
}
