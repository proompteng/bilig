import type { SpreadsheetEngine } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { WorkbookImageSnapshot, WorkbookShapeSnapshot } from '@bilig/protocol'
import type { WorkbookAgentCommand, WorkbookAgentPreviewRange } from './workbook-agent-bundles.js'

export type WorkbookAgentMediaCommand = Extract<
  WorkbookAgentCommand,
  { kind: 'upsertImage' } | { kind: 'deleteImage' } | { kind: 'upsertShape' } | { kind: 'deleteShape' }
>

const SHAPE_TYPE_VALUES = [
  'rectangle',
  'roundedRectangle',
  'ellipse',
  'line',
  'arrow',
  'textBox',
] as const satisfies readonly WorkbookShapeSnapshot['shapeType'][]
const SHAPE_TYPES = new Set<string>(SHAPE_TYPE_VALUES)

function isWorkbookShapeType(value: unknown): value is WorkbookShapeSnapshot['shapeType'] {
  return typeof value === 'string' && SHAPE_TYPES.has(value)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isWorkbookImageSnapshot(value: unknown): value is WorkbookImageSnapshot {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['sheetName'] === 'string' &&
    typeof value['address'] === 'string' &&
    typeof value['sourceUrl'] === 'string' &&
    typeof value['rows'] === 'number' &&
    Number.isInteger(value['rows']) &&
    value['rows'] > 0 &&
    typeof value['cols'] === 'number' &&
    Number.isInteger(value['cols']) &&
    value['cols'] > 0 &&
    (value['altText'] === undefined || typeof value['altText'] === 'string')
  )
}

function isWorkbookShapeSnapshot(value: unknown): value is WorkbookShapeSnapshot {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['sheetName'] === 'string' &&
    typeof value['address'] === 'string' &&
    isWorkbookShapeType(value['shapeType']) &&
    typeof value['rows'] === 'number' &&
    Number.isInteger(value['rows']) &&
    value['rows'] > 0 &&
    typeof value['cols'] === 'number' &&
    Number.isInteger(value['cols']) &&
    value['cols'] > 0 &&
    (value['text'] === undefined || typeof value['text'] === 'string') &&
    (value['fillColor'] === undefined || typeof value['fillColor'] === 'string') &&
    (value['strokeColor'] === undefined || typeof value['strokeColor'] === 'string')
  )
}

function targetRange(sheetName: string, address: string, rows: number, cols: number) {
  const anchor = parseCellAddress(address, sheetName)
  return {
    sheetName,
    startAddress: address,
    endAddress: formatAddress(anchor.row + Math.max(rows, 1) - 1, anchor.col + Math.max(cols, 1) - 1),
  }
}

export function isWorkbookAgentMediaCommandKind(kind: string): kind is WorkbookAgentMediaCommand['kind'] {
  return kind === 'upsertImage' || kind === 'deleteImage' || kind === 'upsertShape' || kind === 'deleteShape'
}

export function isWorkbookAgentMediaCommand(command: WorkbookAgentCommand): command is WorkbookAgentMediaCommand {
  return isWorkbookAgentMediaCommandKind(command.kind)
}

export function isWorkbookAgentMediaCommandValue(value: unknown): value is WorkbookAgentMediaCommand {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'upsertImage':
      return isWorkbookImageSnapshot(value['image'])
    case 'deleteImage':
      return typeof value['id'] === 'string'
    case 'upsertShape':
      return isWorkbookShapeSnapshot(value['shape'])
    case 'deleteShape':
      return typeof value['id'] === 'string'
    default:
      return false
  }
}

export function isHighRiskWorkbookAgentMediaCommand(_command: WorkbookAgentMediaCommand): boolean {
  return false
}

export function isWorkbookScopeMediaCommand(_command: WorkbookAgentMediaCommand): boolean {
  return false
}

export function describeWorkbookAgentMediaCommand(command: WorkbookAgentMediaCommand): string {
  switch (command.kind) {
    case 'upsertImage':
      return `Set image ${command.image.id} at ${command.image.sheetName}!${command.image.address}`
    case 'deleteImage':
      return `Delete image ${command.id}`
    case 'upsertShape':
      return `Set shape ${command.shape.id} at ${command.shape.sheetName}!${command.shape.address}`
    case 'deleteShape':
      return `Delete shape ${command.id}`
  }
}

export function estimateWorkbookAgentMediaCommandAffectedCells(command: WorkbookAgentMediaCommand): number | null {
  switch (command.kind) {
    case 'upsertImage':
      return Math.max(command.image.rows, 1) * Math.max(command.image.cols, 1)
    case 'upsertShape':
      return Math.max(command.shape.rows, 1) * Math.max(command.shape.cols, 1)
    case 'deleteImage':
    case 'deleteShape':
      return null
  }
}

export function deriveWorkbookAgentMediaCommandPreviewRanges(command: WorkbookAgentMediaCommand): WorkbookAgentPreviewRange[] {
  switch (command.kind) {
    case 'upsertImage':
      return [
        {
          ...targetRange(command.image.sheetName, command.image.address, command.image.rows, command.image.cols),
          role: 'target',
        },
      ]
    case 'upsertShape':
      return [
        {
          ...targetRange(command.shape.sheetName, command.shape.address, command.shape.rows, command.shape.cols),
          role: 'target',
        },
      ]
    case 'deleteImage':
    case 'deleteShape':
      return []
  }
}

export function applyWorkbookAgentMediaCommand(engine: SpreadsheetEngine, command: WorkbookAgentMediaCommand): void {
  switch (command.kind) {
    case 'upsertImage':
      engine.setImage(command.image)
      return
    case 'deleteImage':
      engine.deleteImage(command.id)
      return
    case 'upsertShape':
      engine.setShape(command.shape)
      return
    case 'deleteShape':
      engine.deleteShape(command.id)
      return
  }
}
