export function parseJsonRecord(serialized: string, context: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(serialized)
  return parseRecordValue(parsed, context)
}

export function parseRecordValue(candidate: unknown, context: string): Record<string, unknown> {
  const parsed = candidate
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error(`Expected ${context} to be a JSON object`)
  }
  const record: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(parsed)) {
    record[key] = value
  }
  return record
}

export function isNumberArray(value: unknown): value is number[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === 'number')
}

export function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === 'string')
}

export function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
