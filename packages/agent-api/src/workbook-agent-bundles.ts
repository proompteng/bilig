import type { SpreadsheetEngine } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import {
  applyWorkbookAgentAnnotationCommand,
  deriveWorkbookAgentAnnotationCommandPreviewRanges,
  describeWorkbookAgentAnnotationCommand,
  estimateWorkbookAgentAnnotationCommandAffectedCells,
  isHighRiskWorkbookAgentAnnotationCommand,
  isWorkbookAgentAnnotationCommand,
  isWorkbookScopeAnnotationCommand,
} from './workbook-agent-annotation-commands.js'
import {
  applyWorkbookAgentConditionalFormatCommand,
  deriveWorkbookAgentConditionalFormatCommandPreviewRanges,
  describeWorkbookAgentConditionalFormatCommand,
  estimateWorkbookAgentConditionalFormatCommandAffectedCells,
  isHighRiskWorkbookAgentConditionalFormatCommand,
  isWorkbookAgentConditionalFormatCommand,
  isWorkbookScopeConditionalFormatCommand,
} from './workbook-agent-conditional-format-commands.js'
import {
  applyWorkbookAgentMediaCommand,
  deriveWorkbookAgentMediaCommandPreviewRanges,
  describeWorkbookAgentMediaCommand,
  estimateWorkbookAgentMediaCommandAffectedCells,
  isHighRiskWorkbookAgentMediaCommand,
  isWorkbookAgentMediaCommand,
  isWorkbookScopeMediaCommand,
} from './workbook-agent-media-commands.js'
import {
  applyWorkbookAgentObjectCommand,
  deriveWorkbookAgentObjectCommandPreviewRanges,
  describeWorkbookAgentObjectCommand,
  estimateWorkbookAgentObjectCommandAffectedCells,
  isHighRiskWorkbookAgentObjectCommand,
  isWorkbookAgentObjectCommand,
  isWorkbookScopeObjectCommand,
} from './workbook-agent-object-commands.js'
import {
  applyWorkbookAgentProtectionCommand,
  deriveWorkbookAgentProtectionCommandPreviewRanges,
  describeWorkbookAgentProtectionCommand,
  estimateWorkbookAgentProtectionCommandAffectedCells,
  isHighRiskWorkbookAgentProtectionCommand,
  isWorkbookAgentProtectionCommand,
  isWorkbookScopeProtectionCommand,
} from './workbook-agent-protection-commands.js'
import {
  applyWorkbookAgentStructuralCommand,
  describeWorkbookAgentStructuralCommand,
  deriveWorkbookAgentStructuralCommandPreviewRanges,
  estimateWorkbookAgentStructuralCommandAffectedCells,
  isHighRiskWorkbookAgentStructuralCommand,
  isWorkbookAgentStructuralCommand,
  isWorkbookScopeStructuralCommand,
} from './workbook-agent-structural-commands.js'
import {
  applyWorkbookAgentValidationCommand,
  deriveWorkbookAgentValidationCommandPreviewRanges,
  describeWorkbookAgentValidationCommand,
  estimateWorkbookAgentValidationCommandAffectedCells,
  isHighRiskWorkbookAgentValidationCommand,
  isWorkbookAgentValidationCommand,
  isWorkbookScopeValidationCommand,
} from './workbook-agent-validation-commands.js'
import type {
  WorkbookAgentAcceptedScope,
  WorkbookAgentAppliedBy,
  WorkbookAgentBundleScope,
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
  WorkbookAgentContextRef,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewRange,
  WorkbookAgentRiskClass,
  WorkbookAgentSharedReviewState,
} from './workbook-agent-bundle-types.js'
import { sameWorkbookAgentPreviewRange } from './workbook-agent-preview-summary.js'

export type {
  WorkbookAgentAcceptedScope,
  WorkbookAgentAppliedBy,
  WorkbookAgentBundleScope,
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
  WorkbookAgentContextRef,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewCellDiff,
  WorkbookAgentPreviewChangeKind,
  WorkbookAgentPreviewEffectSummary,
  WorkbookAgentPreviewRange,
  WorkbookAgentPreviewRangeRole,
  WorkbookAgentPreviewSummary,
  WorkbookAgentRiskClass,
  WorkbookAgentSharedReviewRecommendation,
  WorkbookAgentSharedReviewState,
  WorkbookAgentSharedReviewStatus,
  WorkbookAgentUiSelectionRef,
  WorkbookAgentViewportRef,
  WorkbookAgentWriteCellInput,
} from './workbook-agent-bundle-types.js'
export {
  isWorkbookAgentCommand,
  isWorkbookAgentCommandBundle,
  isWorkbookAgentContextRef,
  isWorkbookAgentExecutionRecord,
} from './workbook-agent-bundle-guards.js'
export {
  areWorkbookAgentPreviewSummariesEqual,
  decodeWorkbookAgentPreviewSummary,
  isWorkbookAgentPreviewCellDiff,
  isWorkbookAgentPreviewEffectSummary,
  isWorkbookAgentPreviewRange,
  isWorkbookAgentPreviewSummary,
} from './workbook-agent-preview-summary.js'

function rangeLabel(range: WorkbookAgentPreviewRange): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`
}

export function describeWorkbookAgentCommand(command: WorkbookAgentCommand): string {
  if (isWorkbookAgentStructuralCommand(command)) {
    return describeWorkbookAgentStructuralCommand(command)
  }
  if (isWorkbookAgentObjectCommand(command)) {
    return describeWorkbookAgentObjectCommand(command)
  }
  if (isWorkbookAgentMediaCommand(command)) {
    return describeWorkbookAgentMediaCommand(command)
  }
  if (isWorkbookAgentProtectionCommand(command)) {
    return describeWorkbookAgentProtectionCommand(command)
  }
  if (isWorkbookAgentValidationCommand(command)) {
    return describeWorkbookAgentValidationCommand(command)
  }
  if (isWorkbookAgentConditionalFormatCommand(command)) {
    return describeWorkbookAgentConditionalFormatCommand(command)
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    return describeWorkbookAgentAnnotationCommand(command)
  }
  switch (command.kind) {
    case 'writeRange': {
      const ranges = deriveWorkbookAgentCommandPreviewRanges(command)
      return ranges[0] ? `Write cells in ${rangeLabel(ranges[0])}` : 'Write cells'
    }
    case 'setRangeFormulas':
      return `Set formulas in ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[0]!)}`
    case 'clearRange':
      return `Clear ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[0]!)}`
    case 'formatRange':
      return `Format ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[0]!)}`
    case 'fillRange':
      return `Fill ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[1]!)}`
    case 'copyRange':
      return `Copy into ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[1]!)}`
    case 'moveRange':
      return `Move cells to ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[1]!)}`
    default: {
      const exhaustive: never = command
      return String(exhaustive)
    }
  }
}

function dedupePreviewRanges(ranges: readonly WorkbookAgentPreviewRange[]): WorkbookAgentPreviewRange[] {
  const nextRanges: WorkbookAgentPreviewRange[] = []
  ranges.forEach((range) => {
    if (!nextRanges.some((existing) => sameWorkbookAgentPreviewRange(existing, range))) {
      nextRanges.push(range)
    }
  })
  return nextRanges
}

function summarizeCommands(commands: readonly WorkbookAgentCommand[]): string {
  if (commands.length === 0) {
    return 'No workbook changes staged'
  }
  if (commands.length === 1) {
    return describeWorkbookAgentCommand(commands[0]!)
  }
  const firstSummary = describeWorkbookAgentCommand(commands[0]!)
  return `${firstSummary} and ${String(commands.length - 1)} more change${commands.length === 2 ? '' : 's'}`
}

function isSelectionOnlyCommand(command: WorkbookAgentCommand, context: WorkbookAgentContextRef | null): boolean {
  if (!context) {
    return false
  }
  const selectionSheet = context.selection.sheetName
  const selectionRange = context.selection.range ?? {
    startAddress: context.selection.address,
    endAddress: context.selection.address,
  }
  const ranges = deriveWorkbookAgentCommandPreviewRanges(command)
  if (ranges.length !== 1) {
    return false
  }
  const range = ranges[0]
  if (!range) {
    return false
  }
  return (
    range.role === 'target' &&
    range.sheetName === selectionSheet &&
    range.startAddress === selectionRange.startAddress &&
    range.endAddress === selectionRange.endAddress
  )
}

function deriveWorkbookAgentRiskClass(
  commands: readonly WorkbookAgentCommand[],
  context: WorkbookAgentContextRef | null,
): WorkbookAgentRiskClass {
  if (
    commands.some(
      (command) =>
        (isWorkbookAgentStructuralCommand(command) && isHighRiskWorkbookAgentStructuralCommand(command)) ||
        (isWorkbookAgentObjectCommand(command) && isHighRiskWorkbookAgentObjectCommand(command)) ||
        (isWorkbookAgentMediaCommand(command) && isHighRiskWorkbookAgentMediaCommand(command)) ||
        (isWorkbookAgentProtectionCommand(command) && isHighRiskWorkbookAgentProtectionCommand(command)) ||
        (isWorkbookAgentValidationCommand(command) && isHighRiskWorkbookAgentValidationCommand(command)) ||
        (isWorkbookAgentConditionalFormatCommand(command) && isHighRiskWorkbookAgentConditionalFormatCommand(command)) ||
        (isWorkbookAgentAnnotationCommand(command) && isHighRiskWorkbookAgentAnnotationCommand(command)),
    )
  ) {
    return 'high'
  }
  if (commands.every((command) => command.kind === 'formatRange' && isSelectionOnlyCommand(command, context))) {
    return 'low'
  }
  return 'medium'
}

function deriveWorkbookAgentBundleScope(
  commands: readonly WorkbookAgentCommand[],
  context: WorkbookAgentContextRef | null,
): WorkbookAgentBundleScope {
  if (
    commands.some(
      (command) =>
        (isWorkbookAgentStructuralCommand(command) && isWorkbookScopeStructuralCommand(command)) ||
        (isWorkbookAgentObjectCommand(command) && isWorkbookScopeObjectCommand(command)) ||
        (isWorkbookAgentMediaCommand(command) && isWorkbookScopeMediaCommand(command)) ||
        (isWorkbookAgentProtectionCommand(command) && isWorkbookScopeProtectionCommand(command)) ||
        (isWorkbookAgentValidationCommand(command) && isWorkbookScopeValidationCommand(command)) ||
        (isWorkbookAgentConditionalFormatCommand(command) && isWorkbookScopeConditionalFormatCommand(command)) ||
        (isWorkbookAgentAnnotationCommand(command) && isWorkbookScopeAnnotationCommand(command)),
    )
  ) {
    return 'workbook'
  }

  const ranges = commands.flatMap((command) => deriveWorkbookAgentCommandPreviewRanges(command))
  if (ranges.length === 0) {
    return 'sheet'
  }

  const distinctSheets = new Set(ranges.map((range) => range.sheetName))
  if (distinctSheets.size > 1) {
    return 'workbook'
  }

  if (context && commands.every((command) => isSelectionOnlyCommand(command, context))) {
    return 'selection'
  }

  return 'sheet'
}

function estimateWorkbookAgentAffectedCells(commands: readonly WorkbookAgentCommand[]): number | null {
  let total = 0
  let sawCount = false
  commands.forEach((command) => {
    const next = estimateWorkbookAgentCommandAffectedCells(command)
    if (next !== null) {
      total += next
      sawCount = true
    }
  })
  return sawCount ? total : null
}

export function createWorkbookAgentCommandBundle(input: {
  bundleId?: string
  documentId: string
  threadId: string
  turnId: string
  goalText: string
  baseRevision: number
  context: WorkbookAgentContextRef | null
  commands: readonly WorkbookAgentCommand[]
  now: number
  sharedReview?: WorkbookAgentSharedReviewState | null
}): WorkbookAgentCommandBundle {
  const commands = [...input.commands]
  const scope = deriveWorkbookAgentBundleScope(commands, input.context)
  const riskClass = deriveWorkbookAgentRiskClass(commands, input.context)
  return {
    id: input.bundleId ?? crypto.randomUUID(),
    documentId: input.documentId,
    threadId: input.threadId,
    turnId: input.turnId,
    goalText: input.goalText,
    summary: summarizeCommands(commands),
    scope,
    riskClass,
    baseRevision: input.baseRevision,
    createdAtUnixMs: input.now,
    context: input.context ? structuredClone(input.context) : null,
    commands: commands.map((command) => structuredClone(command)),
    affectedRanges: dedupePreviewRanges(commands.flatMap((command) => deriveWorkbookAgentCommandPreviewRanges(command))),
    estimatedAffectedCells: estimateWorkbookAgentAffectedCells(commands),
    sharedReview: input.sharedReview ? structuredClone(input.sharedReview) : null,
  }
}

export function appendWorkbookAgentCommandToBundle(input: {
  previousBundle: WorkbookAgentCommandBundle | null
  documentId: string
  threadId: string
  turnId: string
  goalText: string
  baseRevision: number
  context: WorkbookAgentContextRef | null
  command: WorkbookAgentCommand
  now: number
}): WorkbookAgentCommandBundle {
  const previousBundle =
    input.previousBundle && input.previousBundle.threadId === input.threadId && input.previousBundle.turnId === input.turnId
      ? input.previousBundle
      : null
  return createWorkbookAgentCommandBundle({
    ...(previousBundle ? { bundleId: previousBundle.id } : {}),
    documentId: input.documentId,
    threadId: input.threadId,
    turnId: input.turnId,
    goalText: input.goalText,
    baseRevision: input.baseRevision,
    context: input.context,
    commands: [...(previousBundle?.commands ?? []), input.command],
    now: previousBundle ? previousBundle.createdAtUnixMs : input.now,
    sharedReview: previousBundle?.sharedReview ?? null,
  })
}

export function describeWorkbookAgentBundle(bundle: WorkbookAgentCommandBundle): string {
  const affectedCells =
    bundle.estimatedAffectedCells === null
      ? 'unknown affected cell count'
      : `${String(bundle.estimatedAffectedCells)} affected cell${bundle.estimatedAffectedCells === 1 ? '' : 's'}`
  return [
    `Prepared workbook review item: ${bundle.summary}.`,
    `Risk: ${bundle.riskClass}. Scope: ${bundle.scope}.`,
    `Change target: ${affectedCells}.`,
  ].join(' ')
}

export function normalizeWorkbookAgentCommandIndexes(
  bundle: Pick<WorkbookAgentCommandBundle, 'commands'>,
  commandIndexes: readonly number[] | null | undefined,
): number[] {
  if (commandIndexes === null || commandIndexes === undefined) {
    return bundle.commands.map((_command, index) => index)
  }
  const requested = new Set<number>()
  commandIndexes.forEach((index) => {
    if (Number.isInteger(index) && index >= 0 && index < bundle.commands.length) {
      requested.add(index)
    }
  })
  return bundle.commands.flatMap((_command, index) => (requested.has(index) ? [index] : []))
}

export function isFullWorkbookAgentCommandSelection(input: {
  bundle: Pick<WorkbookAgentCommandBundle, 'commands'>
  commandIndexes: readonly number[] | null | undefined
}): boolean {
  return normalizeWorkbookAgentCommandIndexes(input.bundle, input.commandIndexes).length === input.bundle.commands.length
}

export function projectWorkbookAgentBundle(input: {
  bundle: WorkbookAgentCommandBundle
  commandIndexes: readonly number[] | null | undefined
  bundleId?: string
  baseRevision?: number
  now?: number
}): WorkbookAgentCommandBundle | null {
  const selectedIndexes = normalizeWorkbookAgentCommandIndexes(input.bundle, input.commandIndexes)
  if (selectedIndexes.length === 0) {
    return null
  }
  if (selectedIndexes.length === input.bundle.commands.length) {
    return {
      ...structuredClone(input.bundle),
      ...(input.bundleId ? { id: input.bundleId } : {}),
      ...(input.baseRevision !== undefined ? { baseRevision: input.baseRevision } : {}),
      ...(input.now !== undefined ? { createdAtUnixMs: input.now } : {}),
    }
  }
  return createWorkbookAgentCommandBundle({
    ...(input.bundleId ? { bundleId: input.bundleId } : {}),
    documentId: input.bundle.documentId,
    threadId: input.bundle.threadId,
    turnId: input.bundle.turnId,
    goalText: input.bundle.goalText,
    baseRevision: input.baseRevision ?? input.bundle.baseRevision,
    context: input.bundle.context,
    commands: selectedIndexes.map((index) => input.bundle.commands[index]!),
    now: input.now ?? input.bundle.createdAtUnixMs,
    sharedReview: input.bundle.sharedReview ?? null,
  })
}

export function splitWorkbookAgentCommandBundle(input: {
  bundle: WorkbookAgentCommandBundle
  acceptedCommandIndexes: readonly number[] | null | undefined
  remainingBaseRevision?: number
  remainingBundleId?: string
  now?: number
}): {
  acceptedBundle: WorkbookAgentCommandBundle | null
  remainingBundle: WorkbookAgentCommandBundle | null
  acceptedScope: WorkbookAgentAcceptedScope | null
  acceptedCommandIndexes: number[]
} {
  const acceptedCommandIndexes = normalizeWorkbookAgentCommandIndexes(input.bundle, input.acceptedCommandIndexes)
  if (acceptedCommandIndexes.length === 0) {
    return {
      acceptedBundle: null,
      remainingBundle: structuredClone(input.bundle),
      acceptedScope: null,
      acceptedCommandIndexes,
    }
  }
  const acceptedSet = new Set(acceptedCommandIndexes)
  const remainingCommandIndexes = input.bundle.commands.flatMap((_command, index) => (acceptedSet.has(index) ? [] : [index]))
  return {
    acceptedBundle: projectWorkbookAgentBundle({
      bundle: input.bundle,
      commandIndexes: acceptedCommandIndexes,
      bundleId: input.bundle.id,
    }),
    remainingBundle: projectWorkbookAgentBundle({
      bundle: input.bundle,
      commandIndexes: remainingCommandIndexes,
      ...(input.remainingBundleId ? { bundleId: input.remainingBundleId } : {}),
      ...(input.remainingBaseRevision !== undefined ? { baseRevision: input.remainingBaseRevision } : {}),
      ...(input.now !== undefined ? { now: input.now } : {}),
    }),
    acceptedScope: acceptedCommandIndexes.length === input.bundle.commands.length ? 'full' : 'partial',
    acceptedCommandIndexes,
  }
}

export function buildWorkbookAgentExecutionRecord(input: {
  bundle: WorkbookAgentCommandBundle
  actorUserId: string
  planText: string | null
  preview: WorkbookAgentExecutionRecord['preview']
  appliedRevision: number
  appliedBy: WorkbookAgentAppliedBy
  acceptedScope: WorkbookAgentAcceptedScope
  now: number
}): WorkbookAgentExecutionRecord {
  return {
    id: crypto.randomUUID(),
    bundleId: input.bundle.id,
    documentId: input.bundle.documentId,
    threadId: input.bundle.threadId,
    turnId: input.bundle.turnId,
    actorUserId: input.actorUserId,
    goalText: input.bundle.goalText,
    planText: input.planText,
    summary: input.bundle.summary,
    scope: input.bundle.scope,
    riskClass: input.bundle.riskClass,
    acceptedScope: input.acceptedScope,
    appliedBy: input.appliedBy,
    baseRevision: input.bundle.baseRevision,
    appliedRevision: input.appliedRevision,
    createdAtUnixMs: input.bundle.createdAtUnixMs,
    appliedAtUnixMs: input.now,
    context: input.bundle.context ? structuredClone(input.bundle.context) : null,
    commands: input.bundle.commands.map((command) => structuredClone(command)),
    preview: input.preview ? structuredClone(input.preview) : null,
  }
}

function normalizeFormula(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula
}

function normalizeRangeBounds(range: CellRangeRef): {
  sheetName: string
  startAddress: string
  endAddress: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
} {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    startRow,
    endRow,
    startCol,
    endCol,
  }
}

function countRangeCells(range: CellRangeRef): number {
  const bounds = normalizeRangeBounds(range)
  return (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1)
}

function writeRangeToRange(command: Extract<WorkbookAgentCommand, { kind: 'writeRange' }>): CellRangeRef {
  const start = parseCellAddress(command.startAddress, command.sheetName)
  const width = command.values.reduce((maxWidth, row) => Math.max(maxWidth, row.length), 0)
  return {
    sheetName: command.sheetName,
    startAddress: command.startAddress,
    endAddress: formatAddress(start.row + command.values.length - 1, start.col + width - 1),
  }
}

export function estimateWorkbookAgentCommandAffectedCells(command: WorkbookAgentCommand): number | null {
  if (isWorkbookAgentStructuralCommand(command)) {
    return estimateWorkbookAgentStructuralCommandAffectedCells(command)
  }
  if (isWorkbookAgentObjectCommand(command)) {
    return estimateWorkbookAgentObjectCommandAffectedCells(command)
  }
  if (isWorkbookAgentMediaCommand(command)) {
    return estimateWorkbookAgentMediaCommandAffectedCells(command)
  }
  if (isWorkbookAgentProtectionCommand(command)) {
    return estimateWorkbookAgentProtectionCommandAffectedCells(command)
  }
  if (isWorkbookAgentValidationCommand(command)) {
    return estimateWorkbookAgentValidationCommandAffectedCells(command)
  }
  if (isWorkbookAgentConditionalFormatCommand(command)) {
    return estimateWorkbookAgentConditionalFormatCommandAffectedCells(command)
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    return estimateWorkbookAgentAnnotationCommandAffectedCells(command)
  }
  switch (command.kind) {
    case 'writeRange':
      return command.values.reduce((sum, row) => sum + row.length, 0)
    case 'setRangeFormulas':
      return command.formulas.reduce((sum, row) => sum + row.length, 0)
    case 'clearRange':
    case 'formatRange':
      return countRangeCells(command.range)
    case 'fillRange':
    case 'copyRange':
    case 'moveRange':
      return countRangeCells(command.target)
    default: {
      const exhaustive: never = command
      return exhaustive
    }
  }
}

export function deriveWorkbookAgentCommandPreviewRanges(command: WorkbookAgentCommand): WorkbookAgentPreviewRange[] {
  if (isWorkbookAgentStructuralCommand(command)) {
    return deriveWorkbookAgentStructuralCommandPreviewRanges(command)
  }
  if (isWorkbookAgentObjectCommand(command)) {
    return deriveWorkbookAgentObjectCommandPreviewRanges(command)
  }
  if (isWorkbookAgentMediaCommand(command)) {
    return deriveWorkbookAgentMediaCommandPreviewRanges(command)
  }
  if (isWorkbookAgentProtectionCommand(command)) {
    return deriveWorkbookAgentProtectionCommandPreviewRanges(command)
  }
  if (isWorkbookAgentValidationCommand(command)) {
    return deriveWorkbookAgentValidationCommandPreviewRanges(command)
  }
  if (isWorkbookAgentConditionalFormatCommand(command)) {
    return deriveWorkbookAgentConditionalFormatCommandPreviewRanges(command)
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    return deriveWorkbookAgentAnnotationCommandPreviewRanges(command)
  }
  switch (command.kind) {
    case 'writeRange':
      return [
        {
          ...writeRangeToRange(command),
          role: 'target',
        },
      ]
    case 'setRangeFormulas':
      return [
        {
          ...normalizeRangeBounds(command.range),
          role: 'target',
        },
      ]
    case 'clearRange':
    case 'formatRange':
      return [
        {
          ...normalizeRangeBounds(command.range),
          role: 'target',
        },
      ]
    case 'fillRange':
    case 'copyRange':
    case 'moveRange':
      return [
        {
          ...normalizeRangeBounds(command.source),
          role: 'source',
        },
        {
          ...normalizeRangeBounds(command.target),
          role: 'target',
        },
      ]
    default: {
      const exhaustive: never = command
      return exhaustive
    }
  }
}

export function applyWorkbookAgentCommand(engine: SpreadsheetEngine, command: WorkbookAgentCommand): void {
  if (isWorkbookAgentStructuralCommand(command)) {
    applyWorkbookAgentStructuralCommand(engine, command)
    return
  }
  if (isWorkbookAgentObjectCommand(command)) {
    applyWorkbookAgentObjectCommand(engine, command)
    return
  }
  if (isWorkbookAgentMediaCommand(command)) {
    applyWorkbookAgentMediaCommand(engine, command)
    return
  }
  if (isWorkbookAgentProtectionCommand(command)) {
    applyWorkbookAgentProtectionCommand(engine, command)
    return
  }
  if (isWorkbookAgentValidationCommand(command)) {
    applyWorkbookAgentValidationCommand(engine, command)
    return
  }
  if (isWorkbookAgentConditionalFormatCommand(command)) {
    applyWorkbookAgentConditionalFormatCommand(engine, command)
    return
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    applyWorkbookAgentAnnotationCommand(engine, command)
    return
  }
  switch (command.kind) {
    case 'writeRange': {
      const start = parseCellAddress(command.startAddress, command.sheetName)
      command.values.forEach((rowValues, rowOffset) => {
        rowValues.forEach((cellInput, colOffset) => {
          const address = formatAddress(start.row + rowOffset, start.col + colOffset)
          if (cellInput === null) {
            engine.clearCell(command.sheetName, address)
            return
          }
          if (typeof cellInput === 'string' || typeof cellInput === 'number' || typeof cellInput === 'boolean') {
            engine.setCellValue(command.sheetName, address, cellInput)
            return
          }
          if ('formula' in cellInput) {
            engine.setCellFormula(command.sheetName, address, normalizeFormula(cellInput.formula))
            return
          }
          engine.setCellValue(command.sheetName, address, cellInput.value)
        })
      })
      return
    }
    case 'setRangeFormulas':
      engine.setRangeFormulas(
        command.range,
        command.formulas.map((row) => row.map((formula) => normalizeFormula(formula))),
      )
      return
    case 'clearRange':
      engine.clearRange(command.range)
      return
    case 'formatRange':
      if (command.patch !== undefined) {
        engine.setRangeStyle(command.range, command.patch)
      }
      if (command.numberFormat !== undefined) {
        engine.setRangeNumberFormat(command.range, command.numberFormat)
      }
      return
    case 'fillRange':
      engine.fillRange(command.source, command.target)
      return
    case 'copyRange':
      engine.copyRange(command.source, command.target)
      return
    case 'moveRange':
      engine.moveRange(command.source, command.target)
      return
    default: {
      const exhaustive: never = command
      throw new Error(`Unhandled workbook agent command: ${JSON.stringify(exhaustive)}`)
    }
  }
}

export function applyWorkbookAgentCommandBundle(engine: SpreadsheetEngine, bundle: Pick<WorkbookAgentCommandBundle, 'commands'>): void {
  bundle.commands.forEach((command) => {
    applyWorkbookAgentCommand(engine, command)
  })
}
