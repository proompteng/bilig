import type { WorkbookSnapshot } from '@bilig/protocol'

export function exportWorkerRuntimeSnapshot(args: {
  readonly exportProjectionSnapshot: (() => WorkbookSnapshot) | null
  readonly getReadyAuthoritativeSnapshot: () => WorkbookSnapshot | null
  readonly pendingMutationCount: number
}): WorkbookSnapshot {
  if (args.exportProjectionSnapshot) {
    return args.exportProjectionSnapshot()
  }

  const readyAuthoritativeSnapshot = args.pendingMutationCount === 0 ? args.getReadyAuthoritativeSnapshot() : null
  if (readyAuthoritativeSnapshot) {
    return readyAuthoritativeSnapshot
  }

  throw new Error('Workbook worker runtime projection snapshot is not ready')
}
