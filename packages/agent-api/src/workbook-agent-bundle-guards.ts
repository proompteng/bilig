import type { CellRangeRef, LiteralInput } from '@bilig/protocol'
import { isWorkbookAgentAnnotationCommandValue } from './workbook-agent-annotation-commands.js'
import { isWorkbookAgentConditionalFormatCommandValue } from './workbook-agent-conditional-format-commands.js'
import { isWorkbookAgentMediaCommandValue } from './workbook-agent-media-commands.js'
import { isWorkbookAgentObjectCommandValue } from './workbook-agent-object-commands.js'
import { isWorkbookAgentProtectionCommandValue } from './workbook-agent-protection-commands.js'
import { isWorkbookAgentStructuralCommandValue } from './workbook-agent-structural-commands.js'
import { isWorkbookAgentValidationCommandValue } from './workbook-agent-validation-commands.js'
import type {
  WorkbookAgentAcceptedScope,
  WorkbookAgentAppliedBy,
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
  WorkbookAgentContextRef,
  WorkbookAgentExecutionRecord,
  WorkbookAgentSharedReviewRecommendation,
  WorkbookAgentSharedReviewState,
  WorkbookAgentSharedReviewStatus,
  WorkbookAgentWriteCellInput,
} from './workbook-agent-bundle-types.js'
import { isWorkbookAgentPreviewRange, isWorkbookAgentPreviewSummary } from './workbook-agent-preview-summary.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafeUnixMs(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isLiteralInputValue(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['startAddress'] === 'string' &&
    typeof value['endAddress'] === 'string'
  )
}

function isWriteCellInput(value: unknown): value is WorkbookAgentWriteCellInput {
  return (
    isLiteralInputValue(value) ||
    (isRecord(value) && isLiteralInputValue(value['value'])) ||
    (isRecord(value) && typeof value['formula'] === 'string' && value['formula'].length > 0)
  )
}

function isCommandArray(value: unknown): value is WorkbookAgentCommand[] {
  return Array.isArray(value) && value.every((entry) => isWorkbookAgentCommand(entry))
}

function isAppliedBy(value: unknown): value is WorkbookAgentAppliedBy {
  return value === 'user' || value === 'auto'
}

function isSharedReviewStatus(value: unknown): value is WorkbookAgentSharedReviewStatus {
  return value === 'pending' || value === 'approved' || value === 'rejected'
}

function isSharedReviewRecommendation(value: unknown): value is WorkbookAgentSharedReviewRecommendation {
  return (
    isRecord(value) &&
    typeof value['userId'] === 'string' &&
    (value['decision'] === 'approved' || value['decision'] === 'rejected') &&
    isSafeUnixMs(value['decidedAtUnixMs'])
  )
}

function isSharedReviewState(value: unknown): value is WorkbookAgentSharedReviewState {
  return (
    isRecord(value) &&
    typeof value['ownerUserId'] === 'string' &&
    isSharedReviewStatus(value['status']) &&
    (value['decidedByUserId'] === null || typeof value['decidedByUserId'] === 'string') &&
    (value['decidedAtUnixMs'] === null || isSafeUnixMs(value['decidedAtUnixMs'])) &&
    Array.isArray(value['recommendations']) &&
    value['recommendations'].every((entry) => isSharedReviewRecommendation(entry))
  )
}

function isAcceptedScope(value: unknown): value is WorkbookAgentAcceptedScope {
  return value === 'full' || value === 'partial'
}

export function isWorkbookAgentContextRef(value: unknown): value is WorkbookAgentContextRef {
  return (
    isRecord(value) &&
    isRecord(value['selection']) &&
    typeof value['selection']['sheetName'] === 'string' &&
    typeof value['selection']['address'] === 'string' &&
    (value['selection']['range'] === undefined ||
      (isRecord(value['selection']['range']) &&
        typeof value['selection']['range']['startAddress'] === 'string' &&
        typeof value['selection']['range']['endAddress'] === 'string')) &&
    isRecord(value['viewport']) &&
    typeof value['viewport']['rowStart'] === 'number' &&
    typeof value['viewport']['rowEnd'] === 'number' &&
    typeof value['viewport']['colStart'] === 'number' &&
    typeof value['viewport']['colEnd'] === 'number'
  )
}

export function isWorkbookAgentCommand(value: unknown): value is WorkbookAgentCommand {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  if (isWorkbookAgentStructuralCommandValue(value)) {
    return true
  }
  if (isWorkbookAgentObjectCommandValue(value)) {
    return true
  }
  if (isWorkbookAgentMediaCommandValue(value)) {
    return true
  }
  if (isWorkbookAgentProtectionCommandValue(value)) {
    return true
  }
  if (isWorkbookAgentValidationCommandValue(value)) {
    return true
  }
  if (isWorkbookAgentConditionalFormatCommandValue(value)) {
    return true
  }
  if (isWorkbookAgentAnnotationCommandValue(value)) {
    return true
  }
  switch (value['kind']) {
    case 'writeRange':
    case 'setRangeFormulas':
      return value['kind'] === 'writeRange'
        ? typeof value['sheetName'] === 'string' &&
            typeof value['startAddress'] === 'string' &&
            Array.isArray(value['values']) &&
            value['values'].every((row) => Array.isArray(row) && row.length > 0 && row.every((cellValue) => isWriteCellInput(cellValue)))
        : isCellRangeRef(value['range']) &&
            Array.isArray(value['formulas']) &&
            value['formulas'].every(
              (row) =>
                Array.isArray(row) && row.length > 0 && row.every((formula) => typeof formula === 'string' && formula.trim().length > 0),
            )
    case 'clearRange':
      return isCellRangeRef(value['range'])
    case 'formatRange':
      return (
        isCellRangeRef(value['range']) &&
        (value['patch'] === undefined || isRecord(value['patch'])) &&
        (value['numberFormat'] === undefined || typeof value['numberFormat'] === 'string' || isRecord(value['numberFormat']))
      )
    case 'fillRange':
    case 'copyRange':
    case 'moveRange':
      return isCellRangeRef(value['source']) && isCellRangeRef(value['target'])
    default:
      return false
  }
}

export function isWorkbookAgentCommandBundle(value: unknown): value is WorkbookAgentCommandBundle {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['documentId'] === 'string' &&
    typeof value['threadId'] === 'string' &&
    typeof value['turnId'] === 'string' &&
    typeof value['goalText'] === 'string' &&
    typeof value['summary'] === 'string' &&
    (value['scope'] === 'selection' || value['scope'] === 'sheet' || value['scope'] === 'workbook') &&
    (value['riskClass'] === 'low' || value['riskClass'] === 'medium' || value['riskClass'] === 'high') &&
    typeof value['baseRevision'] === 'number' &&
    typeof value['createdAtUnixMs'] === 'number' &&
    (value['context'] === null || isWorkbookAgentContextRef(value['context'])) &&
    isCommandArray(value['commands']) &&
    Array.isArray(value['affectedRanges']) &&
    value['affectedRanges'].every((entry) => isWorkbookAgentPreviewRange(entry)) &&
    (value['sharedReview'] === undefined || value['sharedReview'] === null || isSharedReviewState(value['sharedReview'])) &&
    (value['estimatedAffectedCells'] === null || typeof value['estimatedAffectedCells'] === 'number')
  )
}

export function isWorkbookAgentExecutionRecord(value: unknown): value is WorkbookAgentExecutionRecord {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['bundleId'] === 'string' &&
    typeof value['documentId'] === 'string' &&
    typeof value['threadId'] === 'string' &&
    typeof value['turnId'] === 'string' &&
    typeof value['actorUserId'] === 'string' &&
    typeof value['goalText'] === 'string' &&
    (value['planText'] === null || typeof value['planText'] === 'string') &&
    typeof value['summary'] === 'string' &&
    (value['scope'] === 'selection' || value['scope'] === 'sheet' || value['scope'] === 'workbook') &&
    (value['riskClass'] === 'low' || value['riskClass'] === 'medium' || value['riskClass'] === 'high') &&
    isAcceptedScope(value['acceptedScope']) &&
    isAppliedBy(value['appliedBy']) &&
    typeof value['baseRevision'] === 'number' &&
    typeof value['appliedRevision'] === 'number' &&
    typeof value['createdAtUnixMs'] === 'number' &&
    typeof value['appliedAtUnixMs'] === 'number' &&
    (value['context'] === null || isWorkbookAgentContextRef(value['context'])) &&
    isCommandArray(value['commands']) &&
    (value['preview'] === null || isWorkbookAgentPreviewSummary(value['preview']))
  )
}
