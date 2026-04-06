import { describe, expect, it } from "vitest";
import type { ZeroSyncService } from "../zero/service.js";
import type {
  CodexAppServerClientOptions,
  CodexAppServerTransport,
} from "./codex-app-server-client.js";
import { createWorkbookAgentService } from "./workbook-agent-service.js";
import type { CodexServerNotification, CodexTurn } from "./codex-app-server-types.js";

class FakeCodexTransport implements CodexAppServerTransport {
  private readonly listeners = new Set<(notification: CodexServerNotification) => void>();
  private turnCounter = 0;

  async ensureReady() {
    return {
      userAgent: "fake",
      codexHome: "/tmp/fake-codex",
      platformFamily: "unix",
      platformOs: "macos",
    };
  }

  subscribe(listener: (notification: CodexServerNotification) => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  async threadStart() {
    return {
      id: "thr-test",
      preview: "",
      turns: [],
    };
  }

  async threadResume(input: { threadId: string }) {
    return {
      id: input.threadId,
      preview: "",
      turns: [],
    };
  }

  async turnStart(): Promise<CodexTurn> {
    this.turnCounter += 1;
    return {
      id: `turn-${String(this.turnCounter)}`,
      status: "inProgress",
      items: [],
      error: null,
    };
  }

  async turnInterrupt() {}

  async close() {}

  emit(notification: CodexServerNotification): void {
    this.listeners.forEach((listener) => listener(notification));
  }
}

function createZeroSyncStub(): ZeroSyncService {
  return {
    enabled: true,
    async initialize() {},
    async close() {},
    async handleQuery() {
      throw new Error("not used");
    },
    async handleMutate() {
      throw new Error("not used");
    },
    async inspectWorkbook() {
      throw new Error("not used");
    },
    async applyServerMutator() {},
    async loadAuthoritativeEvents() {
      throw new Error("not used");
    },
  };
}

describe("workbook agent service", () => {
  it("streams assistant updates into the session timeline", async () => {
    const fakeCodex = new FakeCodexTransport();
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
        fakeCodex,
    });

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      expect(snapshot.threadId).toBe("thr-test");

      const events: unknown[] = [];
      const unsubscribe = service.subscribe("agent-session-1", (event) => {
        events.push(event);
      });

      const inProgress = await service.startTurn({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          prompt: "Summarize Sheet1",
        },
      });

      expect(inProgress.status).toBe("inProgress");
      expect(inProgress.entries.some((entry) => entry.kind === "user")).toBe(true);

      fakeCodex.emit({
        method: "item/agentMessage/delta",
        params: {
          threadId: "thr-test",
          turnId: "turn-1",
          itemId: "msg-1",
          delta: "Checking Sheet1",
        },
      });
      fakeCodex.emit({
        method: "item/completed",
        params: {
          threadId: "thr-test",
          turnId: "turn-1",
          item: {
            type: "agentMessage",
            id: "msg-1",
            text: "Checking Sheet1",
            phase: null,
            memoryCitation: null,
          },
        },
      });
      fakeCodex.emit({
        method: "turn/completed",
        params: {
          threadId: "thr-test",
          turn: {
            id: "turn-1",
            status: "completed",
            items: [],
            error: null,
          },
        },
      });

      const finalSnapshot = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      });

      expect(finalSnapshot.status).toBe("idle");
      expect(finalSnapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            id: "msg-1",
            kind: "assistant",
            text: "Checking Sheet1",
          }),
        ]),
      );
      expect(events).toContainEqual({
        type: "assistantDelta",
        itemId: "msg-1",
        delta: "Checking Sheet1",
      });

      unsubscribe();
    } finally {
      await service.close();
    }
  });
});
