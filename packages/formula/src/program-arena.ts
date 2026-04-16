export interface ArenaSlice {
  offset: number
  length: number
}

export class Uint32Arena {
  private data = new Uint32Array(64)
  private used = 0

  reset(): void {
    this.used = 0
  }

  append(values: Uint32Array): ArenaSlice {
    const slice = { offset: this.used, length: values.length }
    this.ensureCapacity(this.used + values.length)
    this.data.set(values, this.used)
    this.used += values.length
    return slice
  }

  view(): Uint32Array {
    return this.data.subarray(0, this.used)
  }

  get size(): number {
    return this.used
  }

  private ensureCapacity(required: number): void {
    if (required <= this.data.length) {
      return
    }
    let capacity = this.data.length
    while (capacity < required) {
      capacity *= 2
    }
    const next = new Uint32Array(capacity)
    next.set(this.data)
    this.data = next
  }
}

export class Float64Arena {
  private data = new Float64Array(64)
  private used = 0

  reset(): void {
    this.used = 0
  }

  append(values: ArrayLike<number>): ArenaSlice {
    const slice = { offset: this.used, length: values.length }
    this.ensureCapacity(this.used + values.length)
    for (let index = 0; index < values.length; index += 1) {
      this.data[this.used + index] = values[index]!
    }
    this.used += values.length
    return slice
  }

  view(): Float64Array {
    return this.data.subarray(0, this.used)
  }

  get size(): number {
    return this.used
  }

  private ensureCapacity(required: number): void {
    if (required <= this.data.length) {
      return
    }
    let capacity = this.data.length
    while (capacity < required) {
      capacity *= 2
    }
    const next = new Float64Array(capacity)
    next.set(this.data)
    this.data = next
  }
}
