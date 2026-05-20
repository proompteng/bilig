import { formatAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import {
  canonicalizeWorkbookChangeRange,
  workbookChangeRangeFromAddresses,
  type WorkbookChangeRange,
  type WorkbookEventPayload,
} from '@bilig/zero-sync'

export interface WorkbookChangeDescriptor {
  readonly eventKind: WorkbookEventPayload['kind']
  readonly summary: string
  readonly sheetName: string | null
  readonly anchorAddress: string | null
  readonly range: WorkbookChangeRange | null
}

interface CommitCellOpDescriptor {
  readonly sheetName: string
  readonly address: string
  readonly kind: 'upsertCell' | 'deleteCell'
  readonly formula?: string
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeRange(range: CellRangeRef): WorkbookChangeRange {
  const normalized = canonicalizeWorkbookChangeRange(range, isRecord(range) ? range['scope'] : undefined)
  if (!normalized) {
    throw new Error(`Invalid workbook change range ${range.sheetName}!${range.startAddress}:${range.endAddress}`)
  }
  return normalized
}

function rangeLabel(range: WorkbookChangeRange): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`
}

function columnLabel(columnIndex: number): string {
  return formatAddress(0, columnIndex).replace(/[0-9]+$/u, '')
}

function rowLabel(rowIndex: number): string {
  return String(rowIndex + 1)
}

function rowRangeLabel(startRow: number, count: number): string {
  return count <= 1 ? rowLabel(startRow) : `${rowLabel(startRow)}:${rowLabel(startRow + count - 1)}`
}

function columnRangeLabel(startCol: number, count: number): string {
  return count <= 1 ? columnLabel(startCol) : `${columnLabel(startCol)}:${columnLabel(startCol + count - 1)}`
}

function rangeFromAddresses(sheetName: string, addresses: readonly string[]): WorkbookChangeRange | null {
  return workbookChangeRangeFromAddresses(sheetName, addresses)
}

function summarizeCommitCellOps(ops: readonly CommitCellOpDescriptor[]): WorkbookChangeDescriptor {
  const sheetNames = new Set(ops.map((op) => op.sheetName))
  if (sheetNames.size > 1) {
    const allUpserts = ops.every((op) => op.kind === 'upsertCell')
    const allDeletes = ops.every((op) => op.kind === 'deleteCell')
    const allFormulas = ops.every((op) => op.kind === 'upsertCell' && typeof op.formula === 'string')
    return {
      eventKind: 'renderCommit',
      summary: allFormulas
        ? `Filled ${ops.length} formulas across ${sheetNames.size} sheets`
        : allUpserts
          ? `Updated ${ops.length} cells across ${sheetNames.size} sheets`
          : allDeletes
            ? `Cleared ${ops.length} cells across ${sheetNames.size} sheets`
            : `Changed ${ops.length} cells across ${sheetNames.size} sheets`,
      sheetName: null,
      anchorAddress: null,
      range: null,
    }
  }
  const range = rangeFromAddresses(
    ops[0]!.sheetName,
    ops.map((op) => op.address),
  )
  const allUpserts = ops.every((op) => op.kind === 'upsertCell')
  const allDeletes = ops.every((op) => op.kind === 'deleteCell')
  const allFormulas = ops.every((op) => op.kind === 'upsertCell' && typeof op.formula === 'string')

  if (ops.length === 1) {
    const op = ops[0]!
    if (op.kind === 'deleteCell') {
      return {
        eventKind: 'renderCommit',
        summary: `Cleared ${op.sheetName}!${op.address}`,
        sheetName: op.sheetName,
        anchorAddress: op.address,
        range,
      }
    }
    return {
      eventKind: 'renderCommit',
      summary: typeof op.formula === 'string' ? `Set formula in ${op.sheetName}!${op.address}` : `Updated ${op.sheetName}!${op.address}`,
      sheetName: op.sheetName,
      anchorAddress: op.address,
      range,
    }
  }

  return {
    eventKind: 'renderCommit',
    summary: allFormulas
      ? `Filled ${ops.length} formulas in ${rangeLabel(range!)}`
      : allUpserts
        ? `Updated ${ops.length} cells in ${rangeLabel(range!)}`
        : allDeletes
          ? `Cleared ${ops.length} cells in ${rangeLabel(range!)}`
          : `Changed ${ops.length} cells in ${rangeLabel(range!)}`,
    sheetName: range?.sheetName ?? ops[0]!.sheetName,
    anchorAddress: range?.startAddress ?? ops[0]!.address,
    range,
  }
}

function summarizeRenderCommit(payload: Extract<WorkbookEventPayload, { kind: 'renderCommit' }>): WorkbookChangeDescriptor {
  const cellOps = payload.ops.flatMap((op): CommitCellOpDescriptor[] => {
    if (!isRecord(op) || typeof op['kind'] !== 'string') {
      return []
    }
    if (
      (op['kind'] === 'upsertCell' || op['kind'] === 'deleteCell') &&
      typeof op['sheetName'] === 'string' &&
      typeof op['addr'] === 'string'
    ) {
      const base = {
        sheetName: op['sheetName'],
        address: op['addr'],
        kind: op['kind'],
      } satisfies Omit<CommitCellOpDescriptor, 'formula'>
      return [typeof op['formula'] === 'string' ? { ...base, formula: op['formula'] } : base]
    }
    return []
  })

  if (cellOps.length === payload.ops.length && cellOps.length > 0) {
    return summarizeCommitCellOps(cellOps)
  }

  if (payload.ops.length === 1) {
    const op = payload.ops[0]!
    if (op.kind === 'upsertSheet' && op.name) {
      return {
        eventKind: 'renderCommit',
        summary: `Created sheet ${op.name}`,
        sheetName: op.name,
        anchorAddress: 'A1',
        range: {
          sheetName: op.name,
          startAddress: 'A1',
          endAddress: 'A1',
          scope: 'sheet',
        },
      }
    }
    if (op.kind === 'renameSheet' && op.oldName && op.newName) {
      return {
        eventKind: 'renderCommit',
        summary: `Renamed sheet ${op.oldName} to ${op.newName}`,
        sheetName: op.newName,
        anchorAddress: 'A1',
        range: {
          sheetName: op.newName,
          startAddress: 'A1',
          endAddress: 'A1',
          scope: 'sheet',
        },
      }
    }
    if (op.kind === 'deleteSheet' && op.name) {
      return {
        eventKind: 'renderCommit',
        summary: `Deleted sheet ${op.name}`,
        sheetName: null,
        anchorAddress: null,
        range: null,
      }
    }
  }

  return {
    eventKind: 'renderCommit',
    summary: `Applied ${payload.ops.length} workbook changes`,
    sheetName: null,
    anchorAddress: null,
    range: null,
  }
}

export function buildWorkbookChangeDescriptor(payload: WorkbookEventPayload): WorkbookChangeDescriptor {
  switch (payload.kind) {
    case 'applyAgentCommandBundle': {
      const firstTargetRange = payload.bundle.affectedRanges.find((range) => range.role === 'target')
      return {
        eventKind: payload.kind,
        summary: payload.bundle.summary,
        sheetName: firstTargetRange?.sheetName ?? payload.bundle.context?.selection.sheetName ?? null,
        anchorAddress: firstTargetRange?.startAddress ?? payload.bundle.context?.selection.address ?? null,
        range:
          firstTargetRange === undefined
            ? null
            : {
                sheetName: firstTargetRange.sheetName,
                startAddress: firstTargetRange.startAddress,
                endAddress: firstTargetRange.endAddress,
              },
      }
    }
    case 'setCellValue':
      return {
        eventKind: payload.kind,
        summary: `Updated ${payload.sheetName}!${payload.address}`,
        sheetName: payload.sheetName,
        anchorAddress: payload.address,
        range: { sheetName: payload.sheetName, startAddress: payload.address, endAddress: payload.address },
      }
    case 'setCellFormula':
      return {
        eventKind: payload.kind,
        summary: `Set formula in ${payload.sheetName}!${payload.address}`,
        sheetName: payload.sheetName,
        anchorAddress: payload.address,
        range: { sheetName: payload.sheetName, startAddress: payload.address, endAddress: payload.address },
      }
    case 'clearCell':
      return {
        eventKind: payload.kind,
        summary: `Cleared ${payload.sheetName}!${payload.address}`,
        sheetName: payload.sheetName,
        anchorAddress: payload.address,
        range: { sheetName: payload.sheetName, startAddress: payload.address, endAddress: payload.address },
      }
    case 'clearRange': {
      const range = normalizeRange(payload.range)
      return {
        eventKind: payload.kind,
        summary: `Cleared ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'fillRange': {
      const range = normalizeRange(payload.target)
      return {
        eventKind: payload.kind,
        summary: `Filled ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'copyRange': {
      const range = normalizeRange(payload.target)
      return {
        eventKind: payload.kind,
        summary: `Copied into ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'moveRange': {
      const range = normalizeRange(payload.target)
      return {
        eventKind: payload.kind,
        summary: `Moved cells to ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'insertRows': {
      const anchorAddress = formatAddress(payload.start, 0)
      return {
        eventKind: payload.kind,
        summary: `Inserted rows ${rowRangeLabel(payload.start, payload.count)} on ${payload.sheetName}`,
        sheetName: payload.sheetName,
        anchorAddress,
        range: {
          sheetName: payload.sheetName,
          startAddress: anchorAddress,
          endAddress: formatAddress(payload.start + payload.count - 1, 0),
          scope: 'rows',
        },
      }
    }
    case 'deleteRows': {
      const anchorAddress = formatAddress(payload.start, 0)
      return {
        eventKind: payload.kind,
        summary: `Deleted rows ${rowRangeLabel(payload.start, payload.count)} on ${payload.sheetName}`,
        sheetName: payload.sheetName,
        anchorAddress,
        range: {
          sheetName: payload.sheetName,
          startAddress: anchorAddress,
          endAddress: formatAddress(payload.start + payload.count - 1, 0),
          scope: 'rows',
        },
      }
    }
    case 'insertColumns': {
      const anchorAddress = formatAddress(0, payload.start)
      return {
        eventKind: payload.kind,
        summary: `Inserted columns ${columnRangeLabel(payload.start, payload.count)} on ${payload.sheetName}`,
        sheetName: payload.sheetName,
        anchorAddress,
        range: {
          sheetName: payload.sheetName,
          startAddress: anchorAddress,
          endAddress: formatAddress(0, payload.start + payload.count - 1),
          scope: 'columns',
        },
      }
    }
    case 'deleteColumns': {
      const anchorAddress = formatAddress(0, payload.start)
      return {
        eventKind: payload.kind,
        summary: `Deleted columns ${columnRangeLabel(payload.start, payload.count)} on ${payload.sheetName}`,
        sheetName: payload.sheetName,
        anchorAddress,
        range: {
          sheetName: payload.sheetName,
          startAddress: anchorAddress,
          endAddress: formatAddress(0, payload.start + payload.count - 1),
          scope: 'columns',
        },
      }
    }
    case 'updateRowMetadata': {
      const anchorAddress = formatAddress(payload.startRow, 0)
      return {
        eventKind: payload.kind,
        summary: `Updated rows ${rowRangeLabel(payload.startRow, payload.count)} on ${payload.sheetName}`,
        sheetName: payload.sheetName,
        anchorAddress,
        range: {
          sheetName: payload.sheetName,
          startAddress: anchorAddress,
          endAddress: formatAddress(payload.startRow + payload.count - 1, 0),
          scope: 'rows',
        },
      }
    }
    case 'updateColumnMetadata': {
      const anchorAddress = formatAddress(0, payload.startCol)
      return {
        eventKind: payload.kind,
        summary: `Updated columns ${columnRangeLabel(payload.startCol, payload.count)} on ${payload.sheetName}`,
        sheetName: payload.sheetName,
        anchorAddress,
        range: {
          sheetName: payload.sheetName,
          startAddress: anchorAddress,
          endAddress: formatAddress(0, payload.startCol + payload.count - 1),
          scope: 'columns',
        },
      }
    }
    case 'updateColumnWidth': {
      const anchorAddress = formatAddress(0, payload.columnIndex)
      return {
        eventKind: payload.kind,
        summary: `Resized column ${columnLabel(payload.columnIndex)} on ${payload.sheetName}`,
        sheetName: payload.sheetName,
        anchorAddress,
        range: { sheetName: payload.sheetName, startAddress: anchorAddress, endAddress: anchorAddress, scope: 'columns' },
      }
    }
    case 'setFreezePane':
      return {
        eventKind: payload.kind,
        summary:
          payload.rows === 0 && payload.cols === 0
            ? `Cleared freeze panes on ${payload.sheetName}`
            : `Set freeze panes on ${payload.sheetName}`,
        sheetName: payload.sheetName,
        anchorAddress: 'A1',
        range: { sheetName: payload.sheetName, startAddress: 'A1', endAddress: 'A1', scope: 'sheet' },
      }
    case 'mergeCells': {
      const range = normalizeRange(payload.range)
      return {
        eventKind: payload.kind,
        summary: `Merged ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'unmergeCells': {
      const range = normalizeRange(payload.range)
      return {
        eventKind: payload.kind,
        summary: `Unmerged cells in ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'setRangeStyle': {
      const range = normalizeRange(payload.range)
      return {
        eventKind: payload.kind,
        summary: `Formatted ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'clearRangeStyle': {
      const range = normalizeRange(payload.range)
      return {
        eventKind: payload.kind,
        summary: `Cleared formatting in ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'setRangeNumberFormat': {
      const range = normalizeRange(payload.range)
      return {
        eventKind: payload.kind,
        summary: `Changed number format in ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'clearRangeNumberFormat': {
      const range = normalizeRange(payload.range)
      return {
        eventKind: payload.kind,
        summary: `Cleared number format in ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      }
    }
    case 'renderCommit':
      return summarizeRenderCommit(payload)
    case 'restoreVersion':
      return {
        eventKind: payload.kind,
        summary: `Restored version ${payload.versionName}`,
        sheetName: payload.sheetName ?? null,
        anchorAddress: payload.address ?? null,
        range:
          payload.sheetName && payload.address
            ? { sheetName: payload.sheetName, startAddress: payload.address, endAddress: payload.address }
            : payload.sheetName
              ? { sheetName: payload.sheetName, startAddress: 'A1', endAddress: 'A1', scope: 'sheet' }
              : null,
      }
    case 'revertChange':
      return {
        eventKind: payload.kind,
        summary: `Reverted r${payload.targetRevision}: ${payload.targetSummary}`,
        sheetName: payload.sheetName ?? payload.range?.sheetName ?? null,
        anchorAddress: payload.address ?? payload.range?.startAddress ?? null,
        range: payload.range ? normalizeRange(payload.range) : null,
      }
    case 'redoChange':
      return {
        eventKind: payload.kind,
        summary: `Redid r${payload.targetRevision}: ${payload.targetSummary}`,
        sheetName: payload.sheetName ?? payload.range?.sheetName ?? null,
        anchorAddress: payload.address ?? payload.range?.startAddress ?? null,
        range: payload.range ? normalizeRange(payload.range) : null,
      }
    case 'applyBatch':
      return {
        eventKind: payload.kind,
        summary: `Applied ${payload.batch.ops.length} synced operations`,
        sheetName: null,
        anchorAddress: null,
        range: null,
      }
    default: {
      const exhaustive: never = payload
      return exhaustive
    }
  }
}
