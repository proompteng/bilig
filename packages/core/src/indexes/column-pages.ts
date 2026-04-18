import { ValueTag, type CellValue } from '@bilig/protocol'

export interface RuntimeColumnPage {
  readonly rowStart: number
  readonly tags: Uint8Array
  readonly numbers: Float64Array
  readonly stringIds: Uint32Array
  readonly errors: Uint16Array
}

export function createRuntimeColumnPage(rowStart: number, pageSize: number): RuntimeColumnPage {
  return {
    rowStart,
    tags: new Uint8Array(pageSize),
    numbers: new Float64Array(pageSize),
    stringIds: new Uint32Array(pageSize),
    errors: new Uint16Array(pageSize),
  }
}

export function decodeValueTag(rawTag: number | undefined): ValueTag {
  if (rawTag === undefined) {
    return ValueTag.Empty
  }
  switch (rawTag) {
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    case 0:
    default:
      return ValueTag.Empty
  }
}

export function readRuntimeColumnPageEntry(
  page: RuntimeColumnPage,
  absoluteRow: number,
): {
  rawTag: number
  number: number
  stringId: number
  error: number
} {
  const localRow = absoluteRow - page.rowStart
  return {
    rawTag: page.tags[localRow] ?? ValueTag.Empty,
    number: page.numbers[localRow] ?? 0,
    stringId: page.stringIds[localRow] ?? 0,
    error: page.errors[localRow] ?? 0,
  }
}

export function materializeRuntimeColumnPageValue(
  page: RuntimeColumnPage,
  absoluteRow: number,
  lookupString: (stringId: number) => string | undefined,
): CellValue {
  const entry = readRuntimeColumnPageEntry(page, absoluteRow)
  const tag = decodeValueTag(entry.rawTag)
  switch (tag) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty }
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: entry.number }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: entry.number !== 0 }
    case ValueTag.String:
      return {
        tag: ValueTag.String,
        value: entry.stringId === 0 ? '' : (lookupString(entry.stringId) ?? ''),
        stringId: entry.stringId,
      }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: entry.error }
  }
}

export function patchRuntimeColumnPageValue(args: {
  readonly page: RuntimeColumnPage
  readonly row: number
  readonly value: CellValue
  readonly stringId?: number
}): void {
  const { page, row, value, stringId = 0 } = args
  const localRow = row - page.rowStart
  switch (value.tag) {
    case ValueTag.Empty:
      page.tags[localRow] = ValueTag.Empty
      page.numbers[localRow] = 0
      page.stringIds[localRow] = 0
      page.errors[localRow] = 0
      return
    case ValueTag.Number:
      page.tags[localRow] = ValueTag.Number
      page.numbers[localRow] = Object.is(value.value, -0) ? 0 : value.value
      page.stringIds[localRow] = 0
      page.errors[localRow] = 0
      return
    case ValueTag.Boolean:
      page.tags[localRow] = ValueTag.Boolean
      page.numbers[localRow] = value.value ? 1 : 0
      page.stringIds[localRow] = 0
      page.errors[localRow] = 0
      return
    case ValueTag.String:
      page.tags[localRow] = ValueTag.String
      page.numbers[localRow] = 0
      page.stringIds[localRow] = stringId
      page.errors[localRow] = 0
      return
    case ValueTag.Error:
      page.tags[localRow] = ValueTag.Error
      page.numbers[localRow] = 0
      page.stringIds[localRow] = 0
      page.errors[localRow] = value.code
  }
}
