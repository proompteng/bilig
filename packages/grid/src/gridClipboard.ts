export function serializeClipboardMatrix(values: readonly (readonly string[])[]): string {
  return values.map((row) => row.join('\u001f')).join('\u001e')
}

export function serializeClipboardPlainText(values: readonly (readonly string[])[]): string {
  return values.map((row) => row.join('\t')).join('\n')
}

function normalizeClipboardLineEndings(rawText: string): string {
  return rawText.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
}

function trimTrailingEmptyRow(values: string[][]): string[][] {
  if (values.length === 0) {
    return values
  }
  const lastRow = values.at(-1)
  if (lastRow && lastRow.every((value) => value.length === 0)) {
    values.pop()
  }
  return values
}

function parseDelimitedText(rawText: string, delimiter: '\t' | ','): string[][] {
  const normalized = normalizeClipboardLineEndings(rawText)
  const values: string[][] = []
  let row: string[] = []
  let cell = ''
  let inQuotes = false

  const pushCell = () => {
    row.push(cell)
    cell = ''
  }

  const pushRow = () => {
    pushCell()
    values.push(row)
    row = []
  }

  for (let index = 0; index < normalized.length; index += 1) {
    const char = normalized[index]!
    if (char === '"') {
      const nextChar = normalized[index + 1]
      if (inQuotes && nextChar === '"') {
        cell += '"'
        index += 1
        continue
      }
      inQuotes = !inQuotes
      continue
    }
    if (!inQuotes && char === delimiter) {
      pushCell()
      continue
    }
    if (!inQuotes && char === '\n') {
      pushRow()
      continue
    }
    cell += char
  }

  if (cell.length > 0 || row.length > 0 || normalized.endsWith(delimiter)) {
    pushRow()
  }

  return trimTrailingEmptyRow(values)
}

function decodeHtmlEntity(entity: string): string {
  switch (entity) {
    case 'amp':
      return '&'
    case 'lt':
      return '<'
    case 'gt':
      return '>'
    case 'quot':
      return '"'
    case 'apos':
    case '#39':
      return "'"
    case 'nbsp':
      return ' '
    default:
      if (entity.startsWith('#x')) {
        return String.fromCodePoint(Number.parseInt(entity.slice(2), 16))
      }
      if (entity.startsWith('#')) {
        return String.fromCodePoint(Number.parseInt(entity.slice(1), 10))
      }
      return `&${entity};`
  }
}

function decodeHtmlEntities(value: string): string {
  return value.replace(/&([a-z0-9#]+);/gi, (_match, entity: string) => decodeHtmlEntity(entity.toLowerCase()))
}

function stripHtmlCellContent(value: string): string {
  return decodeHtmlEntities(
    value
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(div|p|li|tr|h[1-6])>\s*<(div|p|li|tr|h[1-6])[^>]*>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/\u00a0/g, ' '),
  )
}

export function parseClipboardHtml(rawHtml: string): readonly (readonly string[])[] {
  if (!rawHtml || !/<tr[\s>]/i.test(rawHtml)) {
    return []
  }

  const rows: string[][] = []
  const rowPattern = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi

  for (const rowMatch of rawHtml.matchAll(rowPattern)) {
    const rowHtml = rowMatch[1] ?? ''
    const cells: string[] = []
    const cellPattern = /<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi

    for (const cellMatch of rowHtml.matchAll(cellPattern)) {
      cells.push(stripHtmlCellContent(cellMatch[1] ?? ''))
    }

    if (cells.length > 0) {
      rows.push(cells)
    }
  }

  return rows
}

export function parseClipboardContent(plainText: string, rawHtml?: string): readonly (readonly string[])[] {
  const htmlValues = rawHtml ? parseClipboardHtml(rawHtml) : []
  if (htmlValues.length > 0) {
    return htmlValues
  }
  return parseClipboardPlainText(plainText)
}

export function parseClipboardPlainText(rawText: string): readonly (readonly string[])[] {
  if (rawText.length === 0) {
    return []
  }
  const delimiter: '\t' | ',' = rawText.includes('\t') ? '\t' : ','
  return parseDelimitedText(rawText, delimiter)
}
