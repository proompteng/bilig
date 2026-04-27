export type JsonPrimitive = boolean | number | string | null
export type JsonValue = JsonPrimitive | JsonValue[] | { [key: string]: JsonValue }

export type CodexRequestId = string | number

export type CodexAskForApproval =
  | 'untrusted'
  | 'on-failure'
  | 'on-request'
  | 'never'
  | {
      granular: {
        mcp_elicitations: boolean
        request_permissions?: boolean
        rules: boolean
        sandbox_approval: boolean
        skill_approval?: boolean
      }
    }

export type CodexApprovalsReviewer = 'user' | 'auto_review' | 'guardian_subagent'
export type CodexSandboxMode = 'read-only' | 'workspace-write' | 'danger-full-access'
export type CodexReasoningEffort = 'none' | 'minimal' | 'low' | 'medium' | 'high' | 'xhigh'
export type CodexMessagePhase = 'commentary' | 'final_answer'
export type CodexThreadActiveFlag = 'waitingOnApproval' | 'waitingOnUserInput'
export type CodexThreadStatus =
  | { type: 'notLoaded' }
  | { type: 'idle' }
  | { type: 'systemError' }
  | { type: 'active'; activeFlags: CodexThreadActiveFlag[] }
export type CodexSessionSource = 'cli' | 'vscode' | 'exec' | 'appServer' | 'unknown' | { custom: string } | { subAgent: JsonValue }
export type CodexDynamicToolCallStatus = 'inProgress' | 'completed' | 'failed'

export interface CodexClientInfo {
  name: string
  title?: string | null
  version: string
}

export interface CodexInitializeCapabilities {
  experimentalApi?: boolean
  optOutNotificationMethods?: string[] | null
}

export interface CodexInitializeResponse {
  userAgent: string
  codexHome: string
  platformFamily: string
  platformOs: string
}

export interface CodexThread {
  id: string
  preview: string
  turns: CodexTurn[]
  agentNickname?: string | null
  agentRole?: string | null
  cliVersion?: string
  createdAt?: number
  cwd?: string
  ephemeral?: boolean
  forkedFromId?: string | null
  gitInfo?: JsonValue | null
  modelProvider?: string
  name?: string | null
  path?: string | null
  source?: CodexSessionSource
  status?: CodexThreadStatus
  updatedAt?: number
}

export interface CodexTurn {
  id: string
  status: 'completed' | 'interrupted' | 'failed' | 'inProgress'
  items: CodexThreadItem[]
  completedAt?: number | null
  durationMs?: number | null
  error?: CodexTurnError | null
  startedAt?: number | null
}

export interface CodexTurnError {
  message: string
  additionalDetails?: string | null
  codexErrorInfo?: JsonValue | null
  [key: string]: JsonValue | undefined
}

export interface CodexThreadStartResponse {
  thread: CodexThread
  approvalPolicy?: CodexAskForApproval
  approvalsReviewer?: CodexApprovalsReviewer
  cwd?: string
  instructionSources?: string[]
  model?: string
  modelProvider?: string
  permissionProfile?: JsonValue | null
  reasoningEffort?: CodexReasoningEffort | null
  sandbox?: JsonValue
  serviceTier?: 'fast' | 'flex' | null
}

export interface CodexTurnStartResponse {
  turn: CodexTurn
}

export type CodexUserInput =
  | {
      type: 'text'
      text: string
      text_elements?: JsonValue[]
    }
  | {
      type: 'image'
      url: string
    }
  | {
      type: 'localImage'
      path: string
    }
  | {
      type: 'skill'
      name: string
      path: string
    }
  | {
      type: 'mention'
      name: string
      path: string
    }

export interface CodexDynamicToolSpec {
  name: string
  description: string
  inputSchema: JsonValue
  deferLoading?: boolean
  namespace?: string | null
}

export type CodexDynamicToolCallResult = {
  contentItems: Array<
    | {
        type: 'inputText'
        text: string
      }
    | {
        type: 'inputImage'
        imageUrl: string
      }
  >
  success: boolean
}

export interface CodexDynamicToolCallRequest {
  threadId: string
  turnId: string
  callId: string
  tool: string
  arguments: JsonValue
  namespace?: string | null
}

export type CodexThreadItem =
  | {
      type: 'userMessage'
      id: string
      content: CodexUserInput[]
    }
  | {
      type: 'agentMessage'
      id: string
      text: string
      phase?: string | null
      memoryCitation?: JsonValue | null
    }
  | {
      type: 'plan'
      id: string
      text: string
    }
  | {
      type: 'reasoning'
      id: string
      content?: JsonValue[]
      summary?: JsonValue[]
    }
  | {
      type: 'dynamicToolCall'
      id: string
      tool: string
      arguments: JsonValue
      status: CodexDynamicToolCallStatus
      namespace?: string | null
      contentItems?: Array<
        | {
            type: 'inputText'
            text: string
          }
        | {
            type: 'inputImage'
            imageUrl: string
          }
      > | null
      success?: boolean | null
      durationMs?: number | null
    }
  | {
      type: 'webSearch'
      id: string
      query: string
      action?: JsonValue | null
    }
  | {
      type: 'imageView'
      id: string
      path: string
    }
  | {
      type: 'imageGeneration'
      id: string
      result: string
      status: string
      revisedPrompt?: string | null
      savedPath?: string | null
    }
  | {
      type: 'enteredReviewMode' | 'exitedReviewMode'
      id: string
      review: string
    }
  | {
      type: 'contextCompaction'
      id: string
    }
  | {
      type: string
      id: string
      [key: string]: JsonValue | undefined
    }

export type CodexServerNotification =
  | {
      method: 'thread/started'
      params: {
        thread: CodexThread
      }
    }
  | {
      method: 'turn/started'
      params: {
        threadId: string
        turn: CodexTurn
      }
    }
  | {
      method: 'turn/completed'
      params: {
        threadId: string
        turn: CodexTurn
      }
    }
  | {
      method: 'item/started'
      params: {
        threadId: string
        turnId: string
        item: CodexThreadItem
      }
    }
  | {
      method: 'item/completed'
      params: {
        threadId: string
        turnId: string
        item: CodexThreadItem
      }
    }
  | {
      method: 'item/agentMessage/delta'
      params: {
        threadId: string
        turnId: string
        itemId: string
        delta: string
      }
    }
  | {
      method: 'item/plan/delta'
      params: {
        threadId: string
        turnId: string
        itemId: string
        delta: string
      }
    }
  | {
      method: 'item/reasoning/delta'
      params: {
        threadId: string
        turnId: string
        itemId: string
        delta: string
      }
    }
  | {
      method: 'item/reasoning/textDelta' | 'item/reasoning/summaryTextDelta'
      params: {
        threadId: string
        turnId: string
        itemId: string
        delta: string
      }
    }
  | {
      method: 'item/reasoning/summaryPartAdded'
      params: {
        threadId: string
        turnId: string
        itemId: string
        summaryIndex?: number
      }
    }
  | {
      method: 'error'
      params:
        | {
            error: CodexTurnError
            threadId: string
            turnId: string
            willRetry: boolean
          }
        | {
            message?: string
            [key: string]: JsonValue | undefined
          }
    }

export type CodexServerRequest =
  | {
      method: 'item/tool/call'
      id: CodexRequestId
      params: CodexDynamicToolCallRequest
    }
  | {
      method: string
      id: CodexRequestId
      params?: JsonValue
    }

export interface CodexJsonRpcError {
  code: number
  message: string
  data?: JsonValue
}

export interface CodexJsonRpcResponse<TResult = JsonValue> {
  id: CodexRequestId
  result?: TResult
  error?: CodexJsonRpcError
}
