type JsonPrimitive = null | string | number | boolean
type JsonValue = JsonPrimitive | JsonObject | readonly JsonValue[]
type JsonObject = { readonly [key: string]: JsonValue }

// Matches the generated scorecard line width used before this helper stopped shelling out to oxfmt.
const jsonPrintWidth = 139

export function formatJsonForRepo(args: { readonly rootDir: string; readonly serializedJson: string; readonly tempPrefix: string }): string
export function formatJsonForRepo(serializedJson: string): string
export function formatJsonForRepo(
  args:
    | {
        readonly rootDir: string
        readonly serializedJson: string
        readonly tempPrefix: string
      }
    | string,
): string {
  const serializedJson = typeof args === 'string' ? args : args.serializedJson
  if (typeof args !== 'string') {
    void args.rootDir
    void args.tempPrefix
  }
  const parsedValue: unknown = JSON.parse(serializedJson)
  return `${formatJsonValue(parsedValue, 0, 0)}\n`
}

function formatJsonValue(value: unknown, indent: number, inlinePrefixLength: number): string {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return JSON.stringify(value)
  }
  if (Array.isArray(value)) {
    return formatJsonArray(value, indent, inlinePrefixLength)
  }
  if (isJsonObject(value)) {
    return formatJsonObject(value, indent)
  }
  throw new Error(`Unsupported JSON value type: ${typeof value}`)
}

function formatJsonArray(values: readonly unknown[], indent: number, inlinePrefixLength: number): string {
  if (values.length === 0) {
    return '[]'
  }

  const inlineItems = values.map(formatInlineJsonValue)
  if (inlineItems.every((item): item is string => item !== null)) {
    const inline = `[${inlineItems.join(', ')}]`
    if (inlinePrefixLength + inline.length <= jsonPrintWidth) {
      return inline
    }
  }

  const childIndent = indent + 2
  const childPrefixLength = childIndent
  const lines = values.map((item) => indentFirstLine(formatJsonValue(item, childIndent, childPrefixLength), childIndent))
  return `[\n${lines.map((line, index) => `${line}${index === lines.length - 1 ? '' : ','}`).join('\n')}\n${' '.repeat(indent)}]`
}

function formatJsonObject(value: JsonObject, indent: number): string {
  const entries = Object.entries(value)
  if (entries.length === 0) {
    return '{}'
  }

  const childIndent = indent + 2
  const childIndentText = ' '.repeat(childIndent)
  const lines = entries.map(([key, child], index) => {
    const prefix = `${childIndentText}${JSON.stringify(key)}: `
    const formattedChild = formatJsonValue(child, childIndent, prefix.length)
    return `${prefix}${formattedChild}${index === entries.length - 1 ? '' : ','}`
  })
  return `{\n${lines.join('\n')}\n${' '.repeat(indent)}}`
}

function formatInlineJsonValue(value: unknown): string | null {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return JSON.stringify(value)
  }
  return null
}

function isJsonObject(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function indentFirstLine(value: string, indent: number): string {
  const indentText = ' '.repeat(indent)
  return `${indentText}${value}`
}
