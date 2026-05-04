import { Buffer } from 'node:buffer'
import { normalizeWorkbookAgentToolName, type CodexThread, type CodexThreadItem } from '@bilig/agent-api'
import type {
  WorkbookAgentTextEntryKind,
  WorkbookAgentTimelineCitation,
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
} from '@bilig/contracts'
import { z } from 'zod'

const COMMAND_EXECUTION_TOOL_NAME = 'command_execution'

const workbookAgentRenderedRangeSchema = z.object({
  range: z.object({
    sheetName: z.string().min(1),
    startAddress: z.string().min(1),
    endAddress: z.string().min(1),
  }),
  rowCount: z.number().int().nonnegative(),
  columnCount: z.number().int().nonnegative(),
  cellCount: z.number().int().nonnegative(),
  truncated: z.boolean(),
  rows: z.array(
    z.array(
      z.object({
        address: z.string().min(1),
        input: z.unknown(),
        value: z.unknown(),
        formula: z.string().nullable(),
        displayFormat: z.string().nullable(),
        styleId: z.string().nullable(),
        numberFormatId: z.string().nullable(),
        style: z.unknown(),
      }),
    ),
  ),
})

const workbookAgentUiContextSchema: z.ZodType<WorkbookAgentUiContext> = z.object({
  selection: z.object({
    sheetName: z.string().min(1),
    address: z.string().min(1),
    range: z
      .object({
        startAddress: z.string().min(1),
        endAddress: z.string().min(1),
      })
      .optional(),
  }),
  viewport: z.object({
    rowStart: z.number().int().nonnegative(),
    rowEnd: z.number().int().nonnegative(),
    colStart: z.number().int().nonnegative(),
    colEnd: z.number().int().nonnegative(),
  }),
  rendered: z
    .object({
      capturedAtUnixMs: z.number(),
      capturedRevision: z.number().nullable().optional(),
      batchId: z.number().nullable(),
      selection: workbookAgentRenderedRangeSchema.nullable(),
      visibleRange: workbookAgentRenderedRangeSchema.nullable(),
    })
    .optional(),
})

export const createSessionBodySchema = z.object({
  threadId: z.string().min(1).optional(),
  scope: z.enum(['private', 'shared']).optional(),
  executionPolicy: z.enum(['autoApplySafe', 'autoApplyAll', 'ownerReview']).optional(),
  context: workbookAgentUiContextSchema.optional(),
})

export const updateContextBodySchema = z.object({
  context: workbookAgentUiContextSchema,
})

export const startTurnBodySchema = z.object({
  prompt: z.string().trim().min(1),
  context: workbookAgentUiContextSchema.optional(),
})

export const startWorkflowBodySchema = z
  .discriminatedUnion('workflowTemplate', [
    z.object({
      workflowTemplate: z.literal('summarizeWorkbook'),
    }),
    z.object({
      workflowTemplate: z.literal('summarizeCurrentSheet'),
    }),
    z.object({
      workflowTemplate: z.literal('describeRecentChanges'),
    }),
    z.object({
      workflowTemplate: z.literal('findFormulaIssues'),
      sheetName: z.string().min(1).optional(),
      limit: z.number().int().positive().max(200).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('highlightFormulaIssues'),
      sheetName: z.string().min(1).optional(),
      limit: z.number().int().positive().max(200).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('repairFormulaIssues'),
      sheetName: z.string().min(1).optional(),
      limit: z.number().int().positive().max(200).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('highlightCurrentSheetOutliers'),
      sheetName: z.string().min(1).optional(),
      limit: z.number().int().positive().max(100).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('styleCurrentSheetHeaders'),
      sheetName: z.string().min(1).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('normalizeCurrentSheetHeaders'),
      sheetName: z.string().min(1).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('normalizeCurrentSheetNumberFormats'),
      sheetName: z.string().min(1).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('normalizeCurrentSheetWhitespace'),
      sheetName: z.string().min(1).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('fillCurrentSheetFormulasDown'),
      sheetName: z.string().min(1).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('traceSelectionDependencies'),
    }),
    z.object({
      workflowTemplate: z.literal('explainSelectionCell'),
    }),
    z.object({
      workflowTemplate: z.literal('searchWorkbookQuery'),
      query: z.string().trim().min(1),
      sheetName: z.string().min(1).optional(),
      limit: z.number().int().positive().max(50).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('createCurrentSheetRollup'),
      sheetName: z.string().min(1).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('createCurrentSheetReviewTab'),
      sheetName: z.string().min(1).optional(),
    }),
    z.object({
      workflowTemplate: z.literal('createSheet'),
      name: z.string().trim().min(1),
    }),
    z.object({
      workflowTemplate: z.literal('renameCurrentSheet'),
      name: z.string().trim().min(1),
    }),
    z.object({
      workflowTemplate: z.literal('hideCurrentRow'),
    }),
    z.object({
      workflowTemplate: z.literal('hideCurrentColumn'),
    }),
    z.object({
      workflowTemplate: z.literal('unhideCurrentRow'),
    }),
    z.object({
      workflowTemplate: z.literal('unhideCurrentColumn'),
    }),
  ])
  .and(
    z.object({
      context: workbookAgentUiContextSchema.optional(),
    }),
  )

export const reviewReviewItemBodySchema = z.object({
  decision: z.enum(['approved', 'rejected']),
})

export function createWorkbookAgentBaseInstructions(): string {
  return [
    "You are bilig's workbook assistant inside a spreadsheet product.",
    'Help with the active workbook only.',
    'Use workbook tools for workbook reads, edits, and verification.',
    'Use built-in search or network access when the workbook task needs external context.',
    'If the request needs a capability the available tools do not provide, say what is missing.',
  ].join(' ')
}

export function createWorkbookAgentDeveloperInstructions(): string {
  return [
    'Inspect before you edit unfamiliar cells or ranges.',
    'Prefer the smallest workbook tool that matches the request.',
    'When the request refers to the current cell, selection, or visible area, use the browser workbook context tools first.',
    'Prefer semantic selectors such as named ranges, tables, and table columns over hard-coded coordinates when the workbook exposes them.',
    'Use workbook or range reads for workbook-wide structure or unseen regions instead of guessing.',
    'Use list_tables and list_named_ranges when the workbook likely has meaningful named structures and the user refers to them conceptually.',
    'Use list_sheets, get_sheet_view, get_used_range, and get_current_region when the request is really about sheet layout or navigation state rather than raw cell contents.',
    'Use list_pivots when workbook summaries or edits might depend on pivot outputs.',
    'Use list_charts when workbook summaries or edits may depend on existing chart objects, and use the chart tools when the user asks to create, update, move, or delete charts.',
    'Use list_images and list_shapes when workbook summaries or edits may depend on anchored media objects, and use the media tools when the user asks to insert, move, resize, update, or delete images or shapes.',
    'Use the audit tools when the request is about workbook health, including broken references, hidden-row dependencies, inconsistent formulas, used-range bloat, or performance hotspots.',
    'Use list_data_validation_rules before editing controlled-input ranges, dropdowns, or checkbox cells, and use the data validation tools when the user asks for dropdown-style constraints.',
    'Use get_conditional_formats before reformatting regions that may already carry rule-based emphasis, and use the conditional format tools when the user asks for threshold, text-match, or formula-driven highlighting.',
    'Use get_comments before editing cells that may already carry discussion context or notes, and use the comment and note tools when the user asks to annotate workbook cells.',
    'Use get_protection_status before mutating sensitive sheets or ranges, and use the protection tools when the user asks to lock cells, protect ranges, protect sheets, or hide formulas.',
    'Use set_formula when the user explicitly wants formulas written, not generic values.',
    'Use verify_invariants when the user asks you to validate workbook integrity or when you need a structural confidence check after complex workbook changes.',
    'Range and cell reads expose formatting metadata, including fill/background, font, alignment, borders, number format, intersecting data validation rules, conditional formatting rules, protection metadata, cell comments or notes, and anchored images or shapes when present.',
    'Use the workflow tool only for built-in multi-step or durable tasks.',
    'Use direct structural sheet tools for one-step sheet edits that should happen immediately.',
    'Apply workbook changes directly when the session policy allows it.',
    'When the session policy routes a change set to owner review, summarize the prepared review item in workbook terms.',
    'External search or network context can support an answer, but workbook state must come from workbook tools.',
    'Do not invent unsupported workbook capabilities.',
  ].join(' ')
}

function stringifyJson(value: unknown): string {
  return JSON.stringify(value, null, 2)
}

function formatToolContentItems(
  contentItems:
    | Array<
        | {
            type: 'inputText'
            text: string
          }
        | {
            type: 'inputImage'
            imageUrl: string
          }
      >
    | null
    | undefined,
): string | null {
  if (!contentItems || contentItems.length === 0) {
    return null
  }
  return contentItems.map((item) => (item.type === 'inputText' ? item.text : `[image] ${item.imageUrl}`)).join('\n')
}

function textFromUserContent(
  content: readonly {
    type: 'text'
    text: string
  }[],
): string {
  return content.map((item) => item.text).join('\n')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isUserTextContentItem(value: unknown): value is { type: 'text'; text: string } {
  return isRecord(value) && value['type'] === 'text' && typeof value['text'] === 'string'
}

function isUserMessageItem(
  item: CodexThreadItem,
): item is Extract<CodexThreadItem, { type: 'userMessage' }> & { content: Array<{ type: 'text'; text: string }> } {
  return item.type === 'userMessage' && Array.isArray(item.content) && item.content.every((entry) => isUserTextContentItem(entry))
}

function isAgentMessageItem(item: CodexThreadItem): item is Extract<CodexThreadItem, { type: 'agentMessage' }> {
  return item.type === 'agentMessage' && typeof item.text === 'string' && (item.phase === null || typeof item.phase === 'string')
}

function isPlanItem(item: CodexThreadItem): item is Extract<CodexThreadItem, { type: 'plan' }> {
  return item.type === 'plan' && typeof item.text === 'string'
}

function extractReasoningFragments(value: unknown): string[] {
  if (typeof value === 'string' && value.trim().length > 0) {
    return [value]
  }
  if (Array.isArray(value)) {
    return value.flatMap((entry) => extractReasoningFragments(entry))
  }
  if (!isRecord(value)) {
    return []
  }
  return [
    ...extractReasoningFragments(value['text']),
    ...extractReasoningFragments(value['summary']),
    ...extractReasoningFragments(value['content']),
  ]
}

function extractReasoningText(item: CodexThreadItem): string {
  if (!isRecord(item) || item['type'] !== 'reasoning') {
    return ''
  }
  const text = extractReasoningFragments(item).join('\n').trim()
  return text
}

function isToolContentItem(item: unknown): item is
  | {
      type: 'inputText'
      text: string
    }
  | {
      type: 'inputImage'
      imageUrl: string
    } {
  return (
    typeof item === 'object' &&
    item !== null &&
    (('type' in item && item.type === 'inputText' && 'text' in item && typeof item.text === 'string') ||
      ('type' in item && item.type === 'inputImage' && 'imageUrl' in item && typeof item.imageUrl === 'string'))
  )
}

function isDynamicToolCallItem(item: CodexThreadItem): item is Extract<CodexThreadItem, { type: 'dynamicToolCall' }> {
  return (
    item.type === 'dynamicToolCall' &&
    typeof item.tool === 'string' &&
    (item.status === 'inProgress' || item.status === 'completed' || item.status === 'failed') &&
    (item.contentItems === null || (Array.isArray(item.contentItems) && item.contentItems.every((entry) => isToolContentItem(entry)))) &&
    (item.success === null || typeof item.success === 'boolean')
  )
}

function isCommandExecutionItem(item: CodexThreadItem): item is Extract<CodexThreadItem, { type: 'commandExecution' }> {
  return (
    item.type === 'commandExecution' &&
    typeof item.command === 'string' &&
    typeof item.cwd === 'string' &&
    (item.status === 'inProgress' || item.status === 'completed' || item.status === 'failed') &&
    (item.processId === null || typeof item.processId === 'string') &&
    (item.aggregatedOutput === null || typeof item.aggregatedOutput === 'string') &&
    (item.exitCode === null || typeof item.exitCode === 'number') &&
    (item.durationMs === null || typeof item.durationMs === 'number')
  )
}

function looksMostlyPrintable(value: string): boolean {
  if (value.length === 0 || value.includes('\uFFFD')) {
    return false
  }
  let printable = 0
  let nonWhitespace = 0
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index)
    if (code === 9 || code === 10 || code === 13) {
      printable += 1
      continue
    }
    if (code >= 32 && code !== 127) {
      printable += 1
      if (code !== 32) {
        nonWhitespace += 1
      }
    }
  }
  return nonWhitespace > 0 && printable / value.length >= 0.9
}

function stripAnsiControlSequences(input: string): string {
  const output: string[] = []
  let index = 0
  while (index < input.length) {
    const char = input.charAt(index)
    if (char !== '\u001b') {
      output.push(char)
      index += 1
      continue
    }
    const next = input[index + 1]
    if (next === '[') {
      index += 2
      while (index < input.length) {
        const code = input.charCodeAt(index)
        index += 1
        if (code >= 0x40 && code <= 0x7e) {
          break
        }
      }
      continue
    }
    if (next === ']') {
      index += 2
      while (index < input.length) {
        if (input[index] === '\u0007') {
          index += 1
          break
        }
        if (input[index] === '\u001b' && input[index + 1] === '\\') {
          index += 2
          break
        }
        index += 1
      }
      continue
    }
    if (next === '(' || next === ')') {
      index += 3
      continue
    }
    index += 1
  }
  return output.join('')
}

function sanitizeTerminalOutput(input: string): string {
  const stripped = stripAnsiControlSequences(input)
  const filtered: string[] = []
  let sawNonWhitespace = false
  let sawNonSpinner = false
  for (let index = 0; index < stripped.length; index += 1) {
    const code = stripped.charCodeAt(index)
    const char = stripped[index]!
    if (code === 9 || code === 10 || code === 13) {
      filtered.push(char)
      continue
    }
    if (code < 32 || code === 127) {
      continue
    }
    filtered.push(char)
    if (char.trim().length === 0) {
      continue
    }
    sawNonWhitespace = true
    if (code < 0x2800 || code > 0x28ff) {
      sawNonSpinner = true
    }
  }
  if (!sawNonWhitespace || !sawNonSpinner) {
    return ''
  }
  return filtered.join('')
}

function stripBase64Padding(value: string): string {
  return value.replace(/=+$/g, '')
}

function decodeBase64Token(token: string, options?: { allowShort?: boolean }): string | null {
  const allowShort = options?.allowShort ?? false
  if (!allowShort && token.length < 8) {
    return null
  }
  if (token.length < 2 || token.length % 4 === 1 || !/^[A-Za-z0-9+/]+={0,2}$/.test(token)) {
    return null
  }
  try {
    const padded = token.length % 4 === 0 ? token : `${token}${'='.repeat(4 - (token.length % 4))}`
    const buffer = Buffer.from(padded, 'base64')
    if (stripBase64Padding(buffer.toString('base64')) !== stripBase64Padding(token)) {
      return null
    }
    const decoded = buffer.toString('utf8')
    if (looksMostlyPrintable(decoded)) {
      return decoded
    }
    return sanitizeTerminalOutput(decoded)
  } catch {
    return null
  }
}

function decodeBase64TokenIfPrintable(token: string): string | null {
  const decoded = decodeBase64Token(token, { allowShort: true })
  return decoded !== null && looksMostlyPrintable(decoded) ? decoded : null
}

export function decodeCommandExecutionOutput(raw: string): string {
  if (raw.length < 4 || raw.length % 4 === 1) {
    return raw
  }
  const decodedSingle = decodeBase64Token(raw)
  if (decodedSingle !== null) {
    return decodedSingle
  }
  if (!raw.includes('=')) {
    return raw
  }

  const firstNonBase64Char = raw.search(/[^A-Za-z0-9+/=]/)
  const base64Prefix = firstNonBase64Char === -1 ? raw : raw.slice(0, firstNonBase64Char)
  const suffix = firstNonBase64Char === -1 ? '' : raw.slice(firstNonBase64Char)
  if (!base64Prefix.includes('=')) {
    return raw
  }

  const tokens: string[] = []
  let cursor = 0
  for (let index = 0; index < base64Prefix.length; index += 1) {
    if (base64Prefix[index] !== '=') {
      continue
    }
    let tokenEnd = index
    while (tokenEnd < base64Prefix.length && base64Prefix[tokenEnd] === '=') {
      tokenEnd += 1
    }
    const token = base64Prefix.slice(cursor, tokenEnd)
    if (token.length > 0) {
      tokens.push(token)
    }
    cursor = tokenEnd
    index = tokenEnd - 1
  }

  if (tokens.length === 0) {
    return raw
  }

  const decodedPieces: string[] = []
  for (const token of tokens) {
    const decoded = decodeBase64Token(token, { allowShort: true })
    if (decoded === null) {
      return raw
    }
    decodedPieces.push(decoded)
  }

  const remainder = cursor < base64Prefix.length ? base64Prefix.slice(cursor) : ''
  const decodedRemainder = remainder ? decodeBase64TokenIfPrintable(remainder) : null
  return `${decodedPieces.join('')}${decodedRemainder ?? remainder}${suffix}`
}

function commandExecutionArgumentsText(item: Extract<CodexThreadItem, { type: 'commandExecution' }>): string {
  return stringifyJson({
    command: item.command,
    cwd: item.cwd,
    status: item.status,
    processId: item.processId,
    exitCode: item.exitCode,
    durationMs: item.durationMs,
  })
}

function commandExecutionToolStatus(
  status: Extract<CodexThreadItem, { type: 'commandExecution' }>['status'],
): WorkbookAgentTimelineEntry['toolStatus'] {
  return status === 'completed' ? 'completed' : status === 'failed' ? 'failed' : 'inProgress'
}

export function createCommandExecutionOutputEntry(input: {
  id: string
  turnId: string | null
  outputText: string
}): WorkbookAgentTimelineEntry {
  return {
    id: input.id,
    kind: 'tool',
    turnId: input.turnId,
    text: null,
    phase: null,
    toolName: COMMAND_EXECUTION_TOOL_NAME,
    toolStatus: 'inProgress',
    argumentsText: null,
    outputText: input.outputText,
    success: null,
    citations: [],
  }
}

export function appendCommandExecutionOutput(entry: WorkbookAgentTimelineEntry, outputText: string): WorkbookAgentTimelineEntry {
  return {
    ...entry,
    kind: 'tool',
    text: null,
    toolName: entry.toolName ?? COMMAND_EXECUTION_TOOL_NAME,
    toolStatus: entry.toolStatus ?? 'inProgress',
    outputText: `${entry.outputText ?? ''}${outputText}`,
  }
}

export function createSystemEntry(
  id: string,
  turnId: string | null,
  text: string,
  citations: readonly WorkbookAgentTimelineCitation[] = [],
): WorkbookAgentTimelineEntry {
  return {
    id,
    kind: 'system',
    turnId,
    text,
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [...citations],
  }
}

export function createTextTimelineEntry(input: {
  id: string
  kind: WorkbookAgentTextEntryKind
  turnId: string | null
  text: string
  phase?: string | null
  citations?: readonly WorkbookAgentTimelineCitation[]
}): WorkbookAgentTimelineEntry {
  return {
    id: input.id,
    kind: input.kind,
    turnId: input.turnId,
    text: input.text,
    phase: input.phase ?? null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [...(input.citations ?? [])],
  }
}

export function mapThreadItemToEntry(item: CodexThreadItem, turnId: string | null): WorkbookAgentTimelineEntry {
  if (isUserMessageItem(item)) {
    return {
      id: item.id,
      kind: 'user',
      turnId,
      text: textFromUserContent(item.content),
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
      citations: [],
    }
  }

  if (isAgentMessageItem(item)) {
    return createTextTimelineEntry({
      id: item.id,
      kind: 'assistant',
      turnId,
      text: item.text,
      phase: item.phase ?? null,
    })
  }

  if (isPlanItem(item)) {
    return createTextTimelineEntry({
      id: item.id,
      kind: 'plan',
      turnId,
      text: item.text,
    })
  }

  if (item.type === 'reasoning') {
    return createTextTimelineEntry({
      id: item.id,
      kind: 'reasoning',
      turnId,
      text: extractReasoningText(item),
    })
  }

  if (isDynamicToolCallItem(item)) {
    return {
      id: item.id,
      kind: 'tool',
      turnId,
      text: null,
      phase: null,
      toolName: normalizeWorkbookAgentToolName(item.tool),
      toolStatus: item.status,
      argumentsText: stringifyJson(item.arguments),
      outputText: formatToolContentItems(item.contentItems),
      success: item.success ?? null,
      citations: [],
    }
  }

  if (isCommandExecutionItem(item)) {
    return {
      id: item.id,
      kind: 'tool',
      turnId,
      text: null,
      phase: null,
      toolName: COMMAND_EXECUTION_TOOL_NAME,
      toolStatus: commandExecutionToolStatus(item.status),
      argumentsText: commandExecutionArgumentsText(item),
      outputText: item.aggregatedOutput === null ? null : decodeCommandExecutionOutput(item.aggregatedOutput),
      success: item.exitCode === null ? null : item.exitCode === 0,
      citations: [],
    }
  }

  return createSystemEntry(item.id, turnId, `Codex emitted ${item.type}.`)
}

export function buildEntriesFromThread(thread: CodexThread): WorkbookAgentTimelineEntry[] {
  const entries: WorkbookAgentTimelineEntry[] = []
  for (const turn of thread.turns) {
    for (const item of turn.items) {
      entries.push(mapThreadItemToEntry(item, turn.id))
    }
  }
  return entries
}
