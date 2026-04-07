export type JsonPrimitive = boolean | number | string | null;
export type JsonValue = JsonPrimitive | JsonValue[] | { [key: string]: JsonValue };

export type CodexRequestId = number;

export interface CodexClientInfo {
  name: string;
  title: string;
  version: string;
}

export interface CodexInitializeCapabilities {
  experimentalApi: boolean;
  optOutNotificationMethods?: string[] | null;
}

export interface CodexInitializeResponse {
  userAgent: string;
  codexHome?: string;
  platformFamily?: string;
  platformOs?: string;
}

export interface CodexThread {
  id: string;
  preview: string;
  turns: CodexTurn[];
}

export interface CodexTurn {
  id: string;
  status: "completed" | "interrupted" | "failed" | "inProgress";
  items: CodexThreadItem[];
  error: {
    message?: string;
  } | null;
}

export interface CodexThreadStartResponse {
  thread: CodexThread;
}

export interface CodexTurnStartResponse {
  turn: CodexTurn;
}

export type CodexUserInput = {
  type: "text";
  text: string;
};

export interface CodexDynamicToolSpec {
  name: string;
  description: string;
  inputSchema: JsonValue;
  deferLoading?: boolean;
}

export type CodexDynamicToolCallResult = {
  contentItems: Array<
    | {
        type: "inputText";
        text: string;
      }
    | {
        type: "inputImage";
        imageUrl: string;
      }
  >;
  success: boolean;
};

export interface CodexDynamicToolCallRequest {
  threadId: string;
  turnId: string;
  callId: string;
  tool: string;
  arguments: JsonValue;
}

export type CodexThreadItem =
  | {
      type: "userMessage";
      id: string;
      content: CodexUserInput[];
    }
  | {
      type: "agentMessage";
      id: string;
      text: string;
      phase: string | null;
      memoryCitation: unknown;
    }
  | {
      type: "plan";
      id: string;
      text: string;
    }
  | {
      type: "dynamicToolCall";
      id: string;
      tool: string;
      arguments: JsonValue;
      status: "inProgress" | "completed" | "failed";
      contentItems: Array<
        | {
            type: "inputText";
            text: string;
          }
        | {
            type: "inputImage";
            imageUrl: string;
          }
      > | null;
      success: boolean | null;
      durationMs: number | null;
    }
  | {
      type: string;
      id: string;
      [key: string]: JsonValue | undefined;
    };

export type CodexServerNotification =
  | {
      method: "thread/started";
      params: {
        thread: CodexThread;
      };
    }
  | {
      method: "turn/started";
      params: {
        threadId: string;
        turn: CodexTurn;
      };
    }
  | {
      method: "turn/completed";
      params: {
        threadId: string;
        turn: CodexTurn;
      };
    }
  | {
      method: "item/started";
      params: {
        threadId: string;
        turnId: string;
        item: CodexThreadItem;
      };
    }
  | {
      method: "item/completed";
      params: {
        threadId: string;
        turnId: string;
        item: CodexThreadItem;
      };
    }
  | {
      method: "item/agentMessage/delta";
      params: {
        threadId: string;
        turnId: string;
        itemId: string;
        delta: string;
      };
    }
  | {
      method: "item/plan/delta";
      params: {
        threadId: string;
        turnId: string;
        itemId: string;
        delta: string;
      };
    }
  | {
      method: "error";
      params: {
        message?: string;
        [key: string]: JsonValue | undefined;
      };
    };

export type CodexServerRequest =
  | {
      method: "item/tool/call";
      id: CodexRequestId;
      params: CodexDynamicToolCallRequest;
    }
  | {
      method: string;
      id: CodexRequestId;
      params?: JsonValue;
    };

export interface CodexJsonRpcError {
  code: number;
  message: string;
  data?: JsonValue;
}

export interface CodexJsonRpcResponse<TResult = JsonValue> {
  id: CodexRequestId;
  result?: TResult;
  error?: CodexJsonRpcError;
}
