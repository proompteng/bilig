export interface WorkbookPerfBootstrapResultLike {
  readonly restoredFromPersistence: boolean
  readonly requiresAuthoritativeHydrate: boolean
}

interface WorkbookPerformanceApi {
  mark(markName: string): void
  measure(measureName: string, startOrOptions?: string | { start?: string; end?: string }, endMark?: string): void
}

export interface WorkbookPerfSession {
  readonly scope: string
  markShellMounted(): void
  noteBootstrapResult(result: WorkbookPerfBootstrapResultLike): void
  markFirstAuthoritativePatchVisible(): void
  markFirstAssistantDeltaVisible?(): void
  markFirstAgentApplyVisible?(): void
  markFirstLocalEditApplied?(): void
  markFirstPasteApplied?(): void
  markFirstPreviewVisible?(): void
  markFirstReconcileStarted(): void
  markFirstReconcileSettled(): void
  markFirstSelectionVisible(): void
}

let nextWorkbookPerfSessionId = 1

function resolvePerformanceApi(performanceApi?: WorkbookPerformanceApi | null): WorkbookPerformanceApi | null {
  if (performanceApi) {
    return performanceApi
  }
  const globalPerformance = globalThis.performance
  if (!globalPerformance || typeof globalPerformance.mark !== 'function' || typeof globalPerformance.measure !== 'function') {
    return null
  }
  return globalPerformance
}

function recordMeasure(performanceApi: WorkbookPerformanceApi, measureName: string, startMark: string, endMark: string): void {
  try {
    performanceApi.measure(measureName, { start: startMark, end: endMark })
    return
  } catch {
    try {
      performanceApi.measure(measureName, startMark, endMark)
    } catch {
      // Ignore measurement failures so the product path stays non-blocking.
    }
  }
}

export function createWorkbookPerfSession(input: {
  readonly documentId: string
  readonly performance?: WorkbookPerformanceApi | null
  readonly scope?: string
}): WorkbookPerfSession {
  const performanceApi = resolvePerformanceApi(input.performance)
  const scope = input.scope ?? `bilig:${input.documentId}:perf-session:${nextWorkbookPerfSessionId++}`
  const startMark = `${scope}:start`
  const markedEvents = new Set<string>()

  if (performanceApi) {
    try {
      performanceApi.mark(startMark)
    } catch {
      // Ignore performance API failures so the shell keeps booting.
    }
  }

  const markEvent = (eventName: string): void => {
    if (!performanceApi || markedEvents.has(eventName)) {
      return
    }
    const eventMark = `${scope}:${eventName}`
    try {
      performanceApi.mark(eventMark)
      recordMeasure(performanceApi, `${scope}:time-to-${eventName}`, startMark, eventMark)
      markedEvents.add(eventName)
    } catch {
      // Ignore performance API failures so the shell keeps booting.
    }
  }

  return {
    scope,
    markShellMounted() {
      markEvent('shell-mounted')
    },
    noteBootstrapResult(result) {
      markEvent(
        result.restoredFromPersistence && !result.requiresAuthoritativeHydrate ? 'local-restore-ready' : 'authoritative-hydrate-required',
      )
    },
    markFirstAuthoritativePatchVisible() {
      markEvent('first-authoritative-patch-visible')
    },
    markFirstAssistantDeltaVisible() {
      markEvent('first-assistant-delta-visible')
    },
    markFirstAgentApplyVisible() {
      markEvent('first-agent-apply-visible')
    },
    markFirstLocalEditApplied() {
      markEvent('first-local-edit-applied')
    },
    markFirstPasteApplied() {
      markEvent('first-paste-applied')
    },
    markFirstPreviewVisible() {
      markEvent('first-preview-visible')
    },
    markFirstReconcileStarted() {
      markEvent('first-reconcile-started')
    },
    markFirstReconcileSettled() {
      markEvent('first-reconcile-settled')
    },
    markFirstSelectionVisible() {
      markEvent('first-selection-visible')
    },
  }
}
