export function replaceText(text: string, start: i32, count: i32, replacement: string): string {
  const startIndex = start - 1
  if (startIndex >= text.length) {
    return text
  }
  return text.slice(0, startIndex) + replacement + text.slice(startIndex + count)
}

export function substituteText(text: string, oldText: string, newText: string): string {
  if (oldText.length == 0) {
    return text
  }
  let result = ''
  let searchIndex = 0
  let found = text.indexOf(oldText, searchIndex)
  if (found < 0) {
    return text
  }
  while (found >= 0) {
    result += text.slice(searchIndex, found) + newText
    searchIndex = found + oldText.length
    found = text.indexOf(oldText, searchIndex)
  }
  return result + text.slice(searchIndex)
}

export function substituteNthText(text: string, oldText: string, newText: string, instance: i32): string {
  let count = 0
  let searchIndex = 0
  while (searchIndex <= text.length) {
    const found = text.indexOf(oldText, searchIndex)
    if (found < 0) {
      return text
    }
    count += 1
    if (count == instance) {
      return text.slice(0, found) + newText + text.slice(found + oldText.length)
    }
    searchIndex = found + oldText.length
  }
  return text
}

export function repeatText(text: string, count: i32): string {
  let result = ''
  for (let index = 0; index < count; index += 1) {
    result += text
  }
  return result
}

export function indexOfTextWithMode(text: string, delimiter: string, searchFrom: i32, matchMode: i32): i32 {
  const normalizedText = matchMode == 1 ? text.toLowerCase() : text
  const normalizedDelimiter = matchMode == 1 ? delimiter.toLowerCase() : delimiter
  return normalizedText.indexOf(normalizedDelimiter, max<i32>(0, searchFrom))
}

export function lastIndexOfTextWithMode(text: string, delimiter: string, searchFrom: i32, matchMode: i32): i32 {
  const normalizedText = matchMode == 1 ? text.toLowerCase() : text
  const normalizedDelimiter = matchMode == 1 ? delimiter.toLowerCase() : delimiter
  let start = min<i32>(searchFrom, normalizedText.length - normalizedDelimiter.length)
  if (start < 0) {
    return -1
  }
  while (start >= 0) {
    if (normalizedText.slice(start, start + normalizedDelimiter.length) == normalizedDelimiter) {
      return start
    }
    start -= 1
  }
  return -1
}

export function splitTextByDelimiterWithMode(text: string, delimiter: string, matchMode: i32): Array<string> {
  const parts = new Array<string>()
  if (delimiter.length == 0) {
    parts.push(text)
    return parts
  }
  let cursor = 0
  while (cursor <= text.length) {
    const found = indexOfTextWithMode(text, delimiter, cursor, matchMode)
    if (found < 0) {
      parts.push(text.slice(cursor))
      break
    }
    parts.push(text.slice(cursor, found))
    cursor = found + delimiter.length
  }
  return parts
}
