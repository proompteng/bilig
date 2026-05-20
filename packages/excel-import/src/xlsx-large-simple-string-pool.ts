export class ImportedWorkbookStringPool {
  private readonly values = new Map<string, string>()
  private readonly boundedValues = new Map<string, string>()
  private readonly boundedKeys: string[] = []
  private boundedEvictionIndex = 0

  get count(): number {
    return this.values.size
  }

  intern(value: string): string {
    const existing = this.values.get(value)
    if (existing !== undefined) {
      return existing
    }
    this.values.set(value, value)
    return value
  }

  internBounded(value: string, maxEntries: number): string {
    const existing = this.boundedValues.get(value)
    if (existing !== undefined) {
      return existing
    }
    this.boundedValues.set(value, value)
    this.boundedKeys.push(value)
    this.evictBoundedValues(maxEntries)
    return value
  }

  release(): void {
    this.values.clear()
    this.boundedValues.clear()
    this.boundedKeys.length = 0
    this.boundedEvictionIndex = 0
  }

  private evictBoundedValues(maxEntries: number): void {
    const limit = Math.max(0, Math.trunc(maxEntries))
    while (this.boundedKeys.length - this.boundedEvictionIndex > limit) {
      const key = this.boundedKeys[this.boundedEvictionIndex]
      this.boundedEvictionIndex += 1
      if (key !== undefined) {
        this.boundedValues.delete(key)
      }
    }
    if (this.boundedEvictionIndex > limit && this.boundedEvictionIndex * 2 > this.boundedKeys.length) {
      this.boundedKeys.splice(0, this.boundedEvictionIndex)
      this.boundedEvictionIndex = 0
    }
  }
}
