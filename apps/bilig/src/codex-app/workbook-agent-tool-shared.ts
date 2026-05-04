import type { CodexDynamicToolCallResult, WorkbookAgentCommandBundle, WorkbookAgentExecutionRecord } from '@bilig/agent-api'

export interface WorkbookAgentStageCommandResult {
  readonly bundle: WorkbookAgentCommandBundle
  readonly executionRecord: WorkbookAgentExecutionRecord | null
  readonly disposition?: 'queuedForTurnApply' | 'reviewQueued'
}

export function textToolResult(text: string, success = true): CodexDynamicToolCallResult {
  return {
    success,
    contentItems: [{ type: 'inputText', text }],
  }
}

export function stringifyJson(value: unknown): string {
  return JSON.stringify(value, null, 2)
}
