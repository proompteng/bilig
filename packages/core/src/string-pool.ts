export class StringPool {
  private readonly byValue = new Map<string, number>()
  private readonly values = ['']
  private readonly lengths = [0]

  constructor() {
    this.byValue.set('', 0)
  }

  intern(value: string): number {
    const existing = this.byValue.get(value)
    if (existing !== undefined) {
      return existing
    }
    const id = this.values.length
    this.values.push(value)
    this.lengths.push(value.length)
    this.byValue.set(value, id)
    return id
  }

  get(id: number): string {
    return this.values[id] ?? ''
  }

  get size(): number {
    return this.values.length
  }

  exportLengths(): Uint32Array {
    return Uint32Array.from(this.lengths)
  }

  exportLayout(): {
    offsets: Uint32Array
    lengths: Uint32Array
    data: Uint16Array
  } {
    const offsets = new Uint32Array(this.values.length)
    const lengths = Uint32Array.from(this.lengths)

    let totalUnits = 0
    for (let index = 0; index < this.values.length; index += 1) {
      offsets[index] = totalUnits
      totalUnits += this.values[index]!.length
    }

    const data = new Uint16Array(totalUnits)
    let cursor = 0
    for (const value of this.values) {
      for (let index = 0; index < value.length; index += 1) {
        data[cursor] = value.charCodeAt(index)
        cursor += 1
      }
    }

    return { offsets, lengths, data }
  }
}
