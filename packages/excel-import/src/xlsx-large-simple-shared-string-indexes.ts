import { growUint32Array } from './xlsx-large-simple-array-storage.js'
import { binarySearchUint32 } from './xlsx-large-simple-arena-helpers.js'

export interface LargeSimpleSharedStringIndexSet extends Iterable<number> {
  readonly size: number
  has(index: number): boolean
}

export interface LargeSimpleSharedStringIndexSink {
  add(index: number): void
}

class EmptyLargeSimpleSharedStringIndexSet implements LargeSimpleSharedStringIndexSet {
  readonly size = 0

  has(): boolean {
    return false
  }

  *[Symbol.iterator](): IterableIterator<number> {}
}

class Uint32LargeSimpleSharedStringIndexSet implements LargeSimpleSharedStringIndexSet {
  readonly size: number

  constructor(private readonly indexes: Uint32Array<ArrayBuffer>) {
    this.size = indexes.length
  }

  has(index: number): boolean {
    return Number.isSafeInteger(index) && index >= 0 && binarySearchUint32(this.indexes, index) !== -1
  }

  [Symbol.iterator](): IterableIterator<number> {
    return this.indexes[Symbol.iterator]()
  }
}

export const emptyLargeSimpleSharedStringIndexes: LargeSimpleSharedStringIndexSet = new EmptyLargeSimpleSharedStringIndexSet()

export class LargeSimpleSharedStringIndexCollector implements LargeSimpleSharedStringIndexSink {
  private indexes: Uint32Array<ArrayBuffer> = new Uint32Array(1024)
  private count = 0
  private finalized: LargeSimpleSharedStringIndexSet | undefined

  add(index: number): void {
    if (this.finalized) {
      throw new Error('Cannot add shared-string indexes after finalization.')
    }
    if (!Number.isSafeInteger(index) || index < 0 || index > 0xffffffff) {
      throw new Error(`Invalid shared-string index: ${String(index)}`)
    }
    if (this.count >= this.indexes.length) {
      this.indexes = growUint32Array(this.indexes, this.indexes.length * 2)
    }
    this.indexes[this.count] = index
    this.count += 1
  }

  addAll(indexes: LargeSimpleSharedStringIndexSet): void {
    for (const index of indexes) {
      this.add(index)
    }
  }

  finalize(): LargeSimpleSharedStringIndexSet {
    if (this.finalized) {
      return this.finalized
    }
    if (this.count === 0) {
      this.release()
      this.finalized = emptyLargeSimpleSharedStringIndexes
      return this.finalized
    }
    const sorted = this.indexes.slice(0, this.count)
    sorted.sort()
    let uniqueCount = 0
    for (let index = 0; index < sorted.length; index += 1) {
      const value = sorted[index] ?? 0
      if (uniqueCount === 0 || sorted[uniqueCount - 1] !== value) {
        sorted[uniqueCount] = value
        uniqueCount += 1
      }
    }
    this.indexes = new Uint32Array(0)
    this.count = 0
    this.finalized = new Uint32LargeSimpleSharedStringIndexSet(uniqueCount === sorted.length ? sorted : sorted.slice(0, uniqueCount))
    return this.finalized
  }

  release(): void {
    this.indexes = new Uint32Array(0)
    this.count = 0
  }
}
