import { WORKBOOK_AGENT_TOOL_NAMES, normalizeWorkbookAgentToolName } from '@bilig/agent-api'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

const RENDERED_CONTEXT_WAIT_TIMEOUT_MS = 20_000
const RENDERED_CONTEXT_POLL_INTERVAL_MS = 50

type WorkbookAgentUiContext = WorkbookAgentThreadState['durable']['context']

export function hasRenderedContext(context: WorkbookAgentUiContext): boolean {
  return context?.rendered !== undefined
}

function renderedRevision(context: WorkbookAgentUiContext): number | null {
  const capturedRevision = context?.rendered?.capturedRevision
  if (typeof capturedRevision === 'number' && Number.isSafeInteger(capturedRevision) && capturedRevision >= 0) {
    return capturedRevision
  }
  return null
}

export function hasRenderedContextAtRevision(context: WorkbookAgentUiContext, minRevision: number): boolean {
  const revision = renderedRevision(context)
  return revision !== null && revision >= minRevision
}

export function shouldWaitForRenderedTool(toolName: string): boolean {
  const normalizedTool = normalizeWorkbookAgentToolName(toolName)
  return (
    normalizedTool === WORKBOOK_AGENT_TOOL_NAMES.readRenderedSelection ||
    normalizedTool === WORKBOOK_AGENT_TOOL_NAMES.readRenderedRange ||
    normalizedTool === WORKBOOK_AGENT_TOOL_NAMES.applyAndVerify
  )
}

export async function waitForWorkbookAgentRenderedContext(input: {
  readonly minRevision: number
  readonly refreshContext: () => Promise<WorkbookAgentUiContext>
  readonly isReady?: (context: WorkbookAgentUiContext) => Promise<boolean>
  readonly delay?: (ms: number) => Promise<void>
  readonly now?: () => number
  readonly timeoutMs?: number
  readonly pollIntervalMs?: number
}): Promise<WorkbookAgentUiContext> {
  const delay =
    input.delay ??
    ((ms: number) =>
      new Promise<void>((resolve) => {
        setTimeout(resolve, ms)
      }))
  const now = input.now ?? Date.now
  const timeoutMs = input.timeoutMs ?? RENDERED_CONTEXT_WAIT_TIMEOUT_MS
  const pollIntervalMs = input.pollIntervalMs ?? RENDERED_CONTEXT_POLL_INTERVAL_MS
  const deadline = now() + timeoutMs

  const pollRenderedContext = async (latestContext: WorkbookAgentUiContext): Promise<WorkbookAgentUiContext> => {
    if (hasRenderedContextAtRevision(latestContext, input.minRevision) && (!input.isReady || (await input.isReady(latestContext)))) {
      return latestContext
    }
    if (!hasRenderedContext(latestContext) || now() >= deadline) {
      return latestContext
    }
    await delay(pollIntervalMs)
    return await pollRenderedContext(await input.refreshContext())
  }

  return await pollRenderedContext(await input.refreshContext())
}
