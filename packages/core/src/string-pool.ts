export class StringPool {
  private readonly byValue = new Map<string, number>();
  private readonly values = [""];

  constructor() {
    this.byValue.set("", 0);
  }

  intern(value: string): number {
    const existing = this.byValue.get(value);
    if (existing !== undefined) {
      return existing;
    }
    const id = this.values.length;
    this.values.push(value);
    this.byValue.set(value, id);
    return id;
  }

  get(id: number): string {
    return this.values[id] ?? "";
  }
}
