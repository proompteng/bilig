import type { SpreadsheetEngine } from '@bilig/core'
import type { EngineEvent } from '@bilig/protocol'
import { applyWorkbookEvent, isAuthoritativeWorkbookEventRecord, type AuthoritativeWorkbookEventRecord } from '@bilig/zero-sync'
import type { WorkerEngine } from './worker-runtime-support.js'

export interface AppliedAuthoritativeWorkbookEvents {
  readonly authoritativeEngineEvents: readonly EngineEvent[]
  readonly absorbedMutationIds: ReadonlySet<string>
  readonly payloads: AuthoritativeWorkbookEventRecord['payload'][]
  readonly previousSheets: readonly {
    readonly sheetId: number
    readonly name: string
  }[]
}

export function applyAuthoritativeWorkbookEvents(
  authoritativeEngine: SpreadsheetEngine & WorkerEngine,
  events: readonly unknown[],
): AppliedAuthoritativeWorkbookEvents {
  requireAuthoritativeWorkbookEventRecords(events)

  const previousSheets = [...authoritativeEngine.workbook.sheetsByName.values()].map((sheet) => ({
    sheetId: sheet.id,
    name: sheet.name,
  }))
  const authoritativeEngineEvents: EngineEvent[] = []
  const unsubscribe = authoritativeEngine.subscribe((event) => {
    authoritativeEngineEvents.push(event)
  })
  const absorbedMutationIds = new Set(
    events.flatMap((event) => (typeof event.clientMutationId === 'string' ? [event.clientMutationId] : [])),
  )
  try {
    events.forEach((event) => {
      applyWorkbookEvent(authoritativeEngine, event.payload)
    })
  } finally {
    unsubscribe()
  }

  return {
    authoritativeEngineEvents,
    absorbedMutationIds,
    payloads: events.map((event) => event.payload),
    previousSheets,
  }
}

function requireAuthoritativeWorkbookEventRecords(
  events: readonly unknown[],
): asserts events is readonly AuthoritativeWorkbookEventRecord[] {
  if (!events.every((event) => isAuthoritativeWorkbookEventRecord(event))) {
    throw new Error('Invalid authoritative workbook event batch')
  }
}
