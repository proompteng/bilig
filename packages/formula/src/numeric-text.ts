export function parseNumericText(value: string): number | undefined {
  const trimmed = value.trim()
  if (trimmed.length === 0) {
    return undefined
  }

  const direct = Number(trimmed)
  if (Number.isFinite(direct)) {
    return direct
  }

  if (!trimmed.includes(',')) {
    return undefined
  }

  const grouped = /^([+-]?)(\d{1,3}(?:,\d{3})+)(\.\d*)?([eE][+-]?\d+)?$/.exec(trimmed)
  if (!grouped) {
    return undefined
  }

  const normalized = `${grouped[1] ?? ''}${(grouped[2] ?? '').replaceAll(',', '')}${grouped[3] ?? ''}${grouped[4] ?? ''}`
  const numeric = Number(normalized)
  return Number.isFinite(numeric) ? numeric : undefined
}
