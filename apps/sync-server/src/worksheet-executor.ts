import { decodeAgentFrame, encodeAgentFrame, type AgentFrame } from "@bilig/agent-api";

export interface WorksheetExecutor {
  execute(frame: AgentFrame): Promise<AgentFrame>;
}

export interface HttpWorksheetExecutorOptions {
  baseUrl: string;
  fetchImpl?: typeof fetch;
}

function normalizeBaseUrl(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

export function createHttpWorksheetExecutor(
  options: HttpWorksheetExecutorOptions,
): WorksheetExecutor {
  const baseUrl = normalizeBaseUrl(options.baseUrl);
  const fetchImpl = options.fetchImpl ?? fetch;

  return {
    async execute(frame) {
      const response = await fetchImpl(`${baseUrl}/v1/agent/frames`, {
        method: "POST",
        headers: {
          "content-type": "application/octet-stream",
        },
        body: Buffer.from(encodeAgentFrame(frame)),
      });
      if (!response.ok) {
        throw new Error(`Worksheet executor request failed with status ${response.status}`);
      }
      return decodeAgentFrame(new Uint8Array(await response.arrayBuffer()));
    },
  };
}
