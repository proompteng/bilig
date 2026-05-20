import type { WorkbookRichTextCellSnapshot } from '@bilig/protocol'

const lazyRichTextCellsBrand = Symbol('bilig.lazyImportedXlsxRichTextCells')

interface ImportedWorkbookLazyRichTextCells extends Array<WorkbookRichTextCellSnapshot> {
  readonly [lazyRichTextCellsBrand]: true
}

export function createLazyWorkbookRichTextCells(
  length: number,
  materialize: (index: number) => WorkbookRichTextCellSnapshot,
): WorkbookRichTextCellSnapshot[] {
  const target: ImportedWorkbookLazyRichTextCells = Object.assign([], { [lazyRichTextCellsBrand]: true as const })
  let proxy: WorkbookRichTextCellSnapshot[]
  const read = (index: number): WorkbookRichTextCellSnapshot | undefined =>
    Number.isInteger(index) && index >= 0 && index < length ? materialize(index) : undefined
  const iterate = function* (): IterableIterator<WorkbookRichTextCellSnapshot> {
    for (let index = 0; index < length; index += 1) {
      yield materialize(index)
    }
  }
  proxy = new Proxy<ImportedWorkbookLazyRichTextCells>(target, {
    get: (_target, property) => {
      if (property === lazyRichTextCellsBrand) {
        return true
      }
      if (property === 'length') {
        return length
      }
      if (property === Symbol.iterator || property === 'values') {
        return iterate
      }
      if (property === 'entries') {
        return function* entries(): IterableIterator<[number, WorkbookRichTextCellSnapshot]> {
          for (let index = 0; index < length; index += 1) {
            yield [index, materialize(index)]
          }
        }
      }
      if (property === 'keys') {
        return function* keys(): IterableIterator<number> {
          for (let index = 0; index < length; index += 1) {
            yield index
          }
        }
      }
      if (property === 'at') {
        return (index: number) => read(index < 0 ? length + index : index)
      }
      if (property === 'forEach') {
        return (
          callback: (cell: WorkbookRichTextCellSnapshot, index: number, cells: WorkbookRichTextCellSnapshot[]) => void,
          thisArg?: unknown,
        ) => {
          for (let index = 0; index < length; index += 1) {
            callback.call(thisArg, materialize(index), index, proxy)
          }
        }
      }
      if (property === 'map') {
        return <T>(
          callback: (cell: WorkbookRichTextCellSnapshot, index: number, cells: WorkbookRichTextCellSnapshot[]) => T,
          thisArg?: unknown,
        ): T[] => {
          const output: T[] = []
          for (let index = 0; index < length; index += 1) {
            output.push(callback.call(thisArg, materialize(index), index, proxy))
          }
          return output
        }
      }
      if (property === 'slice') {
        return (start?: number, end?: number) => {
          const from = normalizeSliceIndex(start ?? 0, length)
          const to = normalizeSliceIndex(end ?? length, length)
          const output: WorkbookRichTextCellSnapshot[] = []
          for (let index = from; index < to; index += 1) {
            output.push(materialize(index))
          }
          return output
        }
      }
      if (property === 'toJSON' || property === 'toArray') {
        return () => Array.from(iterate())
      }
      if (typeof property === 'string' && isArrayIndexProperty(property)) {
        return read(Number(property))
      }
      return Reflect.get(Array.prototype, property)
    },
    has: (_target, property) => property === 'length' || (typeof property === 'string' && isArrayIndexProperty(property)),
    getOwnPropertyDescriptor: (_target, property) => {
      if (property === 'length') {
        return { configurable: true, enumerable: false, value: length }
      }
      if (typeof property === 'string' && isArrayIndexProperty(property)) {
        const value = read(Number(property))
        return value === undefined ? undefined : { configurable: true, enumerable: true, value }
      }
      return undefined
    },
  })
  return proxy
}

export function mergeWorkbookRichTextCells(
  base: WorkbookRichTextCellSnapshot[],
  appended: WorkbookRichTextCellSnapshot[],
): WorkbookRichTextCellSnapshot[] {
  if (base.length === 0) {
    return appended
  }
  if (appended.length === 0) {
    return base
  }
  return createLazyWorkbookRichTextCells(base.length + appended.length, (index) =>
    index < base.length ? base[index]! : appended[index - base.length]!,
  )
}

function normalizeSliceIndex(value: number, length: number): number {
  const integer = Math.trunc(value)
  return integer < 0 ? Math.max(0, length + integer) : Math.min(length, integer)
}

function isArrayIndexProperty(property: string): boolean {
  if (property.length === 0 || !/^(?:0|[1-9][0-9]*)$/u.test(property)) {
    return false
  }
  const index = Number(property)
  return Number.isInteger(index) && index >= 0 && index < 2 ** 32 - 1
}
