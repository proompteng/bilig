export function parseSourceMaxLines(value: string | undefined): number {
  if (value === undefined) {
    return 1000
  }

  if (!/^(?:[1-9]\d*)$/u.test(value)) {
    throw new Error(`BILIG_SOURCE_MAX_LINES must be a positive integer, got ${value}`)
  }

  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed)) {
    throw new Error(`BILIG_SOURCE_MAX_LINES must be a safe integer, got ${value}`)
  }

  return parsed
}
