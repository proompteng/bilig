import type { WorkbookAgentCommandBundle, WorkbookAgentPreviewSummary } from '@bilig/agent-api'

const MAX_CACHED_PREVIEWS = 32

const settledPreviewCache = new Map<string, WorkbookAgentPreviewSummary>()
const inFlightPreviewCache = new Map<string, Promise<WorkbookAgentPreviewSummary>>()

function rememberWorkbookAgentPreview(requestKey: string, preview: WorkbookAgentPreviewSummary): WorkbookAgentPreviewSummary {
  settledPreviewCache.delete(requestKey)
  settledPreviewCache.set(requestKey, preview)
  while (settledPreviewCache.size > MAX_CACHED_PREVIEWS) {
    const oldestRequestKey = settledPreviewCache.keys().next().value
    if (!oldestRequestKey) {
      break
    }
    settledPreviewCache.delete(oldestRequestKey)
  }
  return preview
}

export function createWorkbookAgentPreviewRequestKey(input: {
  readonly bundle: Pick<WorkbookAgentCommandBundle, 'id' | 'baseRevision'>
  readonly commandIndexes: readonly number[]
}): string {
  return [input.bundle.id, String(input.bundle.baseRevision), input.commandIndexes.join(',')].join(':')
}

export function clearWorkbookAgentPreviewCache(): void {
  settledPreviewCache.clear()
  inFlightPreviewCache.clear()
}

export function readCachedWorkbookAgentPreview(requestKey: string): WorkbookAgentPreviewSummary | null {
  return settledPreviewCache.get(requestKey) ?? null
}

export function loadWorkbookAgentPreview(input: {
  readonly requestKey: string
  readonly load: () => Promise<WorkbookAgentPreviewSummary>
}): Promise<WorkbookAgentPreviewSummary> {
  const cached = settledPreviewCache.get(input.requestKey)
  if (cached) {
    return Promise.resolve(cached)
  }
  const inFlight = inFlightPreviewCache.get(input.requestKey)
  if (inFlight) {
    return inFlight
  }
  const nextPromise = (async () => {
    try {
      const preview = await input.load()
      return rememberWorkbookAgentPreview(input.requestKey, preview)
    } finally {
      inFlightPreviewCache.delete(input.requestKey)
    }
  })()
  inFlightPreviewCache.set(input.requestKey, nextPromise)
  return nextPromise
}
