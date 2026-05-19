export class ImportedWorkbookStringPool {
  private readonly values = new Map<string, string>()

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

  release(): void {
    this.values.clear()
  }
}
