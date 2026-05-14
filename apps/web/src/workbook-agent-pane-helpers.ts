import type { WorkbookAgentUiContext } from '@bilig/contracts'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function readMessageEventData(event: Event): string | null {
  return event instanceof MessageEvent && typeof event.data === 'string' ? event.data : null
}

export function formatWorkbookAgentContextLabel(context: WorkbookAgentUiContext): string {
  const selectionRange = context.selection.range
  const address =
    selectionRange && selectionRange.startAddress !== selectionRange.endAddress
      ? `${selectionRange.startAddress}:${selectionRange.endAddress}`
      : context.selection.address
  return `${context.selection.sheetName}!${address}`
}

export function stringifyWorkbookAgentContextSyncKey(context: WorkbookAgentUiContext): string {
  const rendered = context.rendered
  return JSON.stringify({
    selection: context.selection,
    viewport: context.viewport,
    rendered:
      rendered === undefined
        ? null
        : {
            capturedRevision: rendered.capturedRevision ?? null,
            selection: rendered.selection,
            visibleRange: rendered.visibleRange,
          },
  })
}

export function readAppliedRevision(record: unknown): number | null {
  if (!isRecord(record)) {
    return null
  }
  const revision = record['appliedRevision']
  return typeof revision === 'number' && Number.isInteger(revision) && revision >= 0 ? revision : null
}

export function normalizeWorkbookAgentErrorMessage(error: string): string {
  if (error.includes('thread/start.dynamicTools requires experimentalApi capability')) {
    return 'Retry in a moment.'
  }
  if (error.includes('Invalid Codex initialize response')) {
    return 'Retry in a moment.'
  }
  return error
}
