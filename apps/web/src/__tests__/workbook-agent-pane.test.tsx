// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { toast } from "sonner";
import { WorkbookToastRegion } from "../WorkbookToastRegion.js";
import { clearWorkbookAgentPreviewCache } from "../workbook-agent-preview-cache.js";
import { useWorkbookAgentPane } from "../use-workbook-agent-pane.js";

async function flushToasts(): Promise<void> {
  await act(async () => {
    await Promise.resolve();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

class MockEventSource {
  static latest: MockEventSource | null = null;
  readonly url: string;
  private readonly listeners = new Map<string, Set<(event: Event) => void>>();

  constructor(url: string) {
    this.url = url;
    MockEventSource.latest = this;
  }

  close() {}

  addEventListener(type: string, listener: (event: Event) => void): void {
    const entries = this.listeners.get(type) ?? new Set();
    entries.add(listener);
    this.listeners.set(type, entries);
  }

  removeEventListener(type: string, listener: (event: Event) => void): void {
    const entries = this.listeners.get(type);
    if (!entries) {
      return;
    }
    entries.delete(listener);
    if (entries.size === 0) {
      this.listeners.delete(type);
    }
  }

  emit(data: unknown): void {
    this.listeners.get("message")?.forEach((listener) => {
      listener(
        new MessageEvent("message", {
          data: JSON.stringify(data),
        }),
      );
    });
  }

  emitError(): void {
    this.listeners.get("error")?.forEach((listener) => {
      listener(new Event("error"));
    });
  }
}

function requestUrl(input: RequestInfo | URL): string {
  if (typeof input === "string") {
    return input;
  }
  if (input instanceof URL) {
    return input.href;
  }
  return input.url;
}

function requestBody(init: RequestInit | undefined): unknown {
  if (!init || typeof init.body !== "string") {
    return null;
  }
  return JSON.parse(init.body) as unknown;
}

function requestMethod(init: RequestInit | undefined): string {
  return init?.method ?? "GET";
}

function createDefaultWorkflowContext() {
  return {
    selection: {
      sheetName: "Sheet1",
      address: "A1",
    },
    viewport: {
      rowStart: 0,
      rowEnd: 10,
      colStart: 0,
      colEnd: 5,
    },
  };
}

function createSnapshot(overrides: Record<string, unknown> = {}) {
  const overrideEntries = Array.isArray(overrides["entries"])
    ? overrides["entries"].map((entry) =>
        typeof entry === "object" && entry !== null && !("citations" in entry)
          ? {
              ...entry,
              citations: [],
            }
          : entry,
      )
    : undefined;
  return {
    sessionId: "agent-session-1",
    documentId: "doc-1",
    threadId: "thr-1",
    executionPolicy: "autoApplyAll",
    scope: "private",
    executionPolicy: "autoApplyAll",
    status: "idle",
    activeTurnId: null,
    lastError: null,
    context: createDefaultWorkflowContext(),
    entries: [
      {
        id: "assistant-1",
        kind: "assistant",
        turnId: "turn-1",
        text: "",
        phase: null,
        toolName: null,
        toolStatus: null,
        argumentsText: null,
        outputText: null,
        success: null,
        citations: [],
      },
    ],
    pendingBundle: null,
    executionRecords: [],
    workflowRuns: [],
    ...overrides,
    ...(overrideEntries ? { entries: overrideEntries } : {}),
  };
}

function createPreviewSummary(overrides: Record<string, unknown> = {}) {
  return {
    ranges: [],
    structuralChanges: [],
    cellDiffs: [],
    effectSummary: {
      displayedCellDiffCount: 0,
      truncatedCellDiffs: false,
      inputChangeCount: 0,
      formulaChangeCount: 0,
      styleChangeCount: 0,
      numberFormatChangeCount: 0,
      structuralChangeCount: 0,
    },
    ...overrides,
  };
}

function createThreadSummary(overrides: Record<string, unknown> = {}) {
  return {
    threadId: "thr-1",
    scope: "private",
    ownerUserId: "alex@example.com",
    updatedAtUnixMs: 100,
    entryCount: 1,
    hasPendingBundle: false,
    latestEntryText: null,
    ...overrides,
  };
}

interface MockZeroAgentHarness {
  readonly zero: {
    materialize(query: unknown): {
      readonly data: unknown;
      addListener(listener: (value: unknown) => void): () => void;
      destroy(): void;
    };
  };
}

function createMockZeroAgentHarness(input: {
  readonly initialThreadSummaries: unknown;
  readonly initialWorkflowRuns: unknown;
}): MockZeroAgentHarness {
  let threadSummaryValue = input.initialThreadSummaries;
  let workflowRunValue = input.initialWorkflowRuns;
  const threadSummaryListeners = new Set<(value: unknown) => void>();
  const workflowRunListeners = new Set<(value: unknown) => void>();
  let materializeCallCount = 0;

  return {
    zero: {
      materialize(_query: unknown) {
        const isThreadSummaryQuery = materializeCallCount === 0;
        materializeCallCount += 1;
        return {
          get data() {
            return isThreadSummaryQuery ? threadSummaryValue : workflowRunValue;
          },
          addListener(listener: (value: unknown) => void) {
            const listeners = isThreadSummaryQuery ? threadSummaryListeners : workflowRunListeners;
            listeners.add(listener);
            return () => {
              listeners.delete(listener);
            };
          },
          destroy() {},
        };
      },
    },
  };
}

function AgentHarness(props: {
  readonly currentUserId?: string;
  readonly previewBundle?: Parameters<typeof useWorkbookAgentPane>[0]["previewBundle"];
  readonly zero?: Parameters<typeof useWorkbookAgentPane>[0]["zero"];
  readonly zeroEnabled?: boolean;
}) {
  const { agentError, agentPanel, clearAgentError } = useWorkbookAgentPane({
    currentUserId: props.currentUserId ?? "alex@example.com",
    documentId: "doc-1",
    enabled: true,
    getContext: () => ({
      selection: {
        sheetName: "Sheet1",
        address: "A1",
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    }),
    previewBundle: props.previewBundle ?? vi.fn(async () => createPreviewSummary()),
    ...(props.zero ? { zero: props.zero } : {}),
    ...(props.zeroEnabled !== undefined ? { zeroEnabled: props.zeroEnabled } : {}),
  });

  return (
    <div>
      <WorkbookToastRegion
        toasts={
          agentError
            ? [
                {
                  id: "agent-error",
                  tone: "error",
                  message: agentError,
                  onDismiss: clearAgentError,
                },
              ]
            : []
        }
      />
      {agentPanel}
    </div>
  );
}

beforeEach(() => {
  vi.stubGlobal("EventSource", MockEventSource);
  window.sessionStorage.clear();
  clearWorkbookAgentPreviewCache();
});

afterEach(() => {
  toast.dismiss();
  vi.restoreAllMocks();
  window.sessionStorage.clear();
  clearWorkbookAgentPreviewCache();
  document.body.innerHTML = "";
});

describe("workbook agent pane", () => {
  it("renders the assistant rail without the skill-card strip", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response(JSON.stringify(createSnapshot()), {
            status: 200,
            headers: { "content-type": "application/json" },
          }),
      ),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const input = host.querySelector("[data-testid='workbook-agent-input']");
    expect(input instanceof HTMLTextAreaElement).toBe(true);
    expect(input instanceof HTMLTextAreaElement ? input.value : "").toBe("");
    expect(host.textContent).not.toContain("Local Skills");
    expect(host.textContent).not.toContain("Inspect Selection");
    expect(host.textContent).not.toContain(
      "Ask the assistant to inspect, edit, or restructure this workbook.",
    );
    expect(host.textContent).not.toContain("Sheet1!A1");
    expect(input instanceof HTMLTextAreaElement ? input.getAttribute("placeholder") : null).toBe(
      "Ask the workbook assistant",
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("renders durable thread summaries and workflow runs from Zero projections", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const zero = createMockZeroAgentHarness({
      initialThreadSummaries: [
        createThreadSummary({
          threadId: "thr-1",
          scope: "shared",
          ownerUserId: "casey@example.com",
          latestEntryText: "Completed workflow: Summarize Workbook",
        }),
      ],
      initialWorkflowRuns: [
        {
          runId: "wf-zero-1",
          threadId: "thr-1",
          startedByUserId: "casey@example.com",
          workflowTemplate: "summarizeWorkbook",
          title: "Summarize Workbook",
          summary: "Summarized workbook structure across 2 sheets.",
          status: "completed",
          createdAtUnixMs: 1,
          updatedAtUnixMs: 2,
          completedAtUnixMs: 2,
          errorMessage: null,
          steps: [
            {
              stepId: "inspect-workbook",
              label: "Inspect workbook structure",
              status: "completed",
              summary: "Read durable workbook structure across 2 sheets.",
              updatedAtUnixMs: 1,
            },
          ],
          artifact: {
            kind: "markdown",
            title: "Workbook Summary",
            text: "## Workbook Summary",
          },
        },
      ],
    });
    sessionStorage.setItem("bilig:workbook-agent:doc-1", JSON.stringify({ threadId: "thr-1" }));
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads/thr-1")) {
        return new Response(
          JSON.stringify(createSnapshot({ threadId: "thr-1", workflowRuns: [] })),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness zero={zero.zero} zeroEnabled />);
    });

    expect(host.querySelector("[data-testid='workbook-agent-scope-private']")).toBeNull();
    expect(host.querySelector("[data-testid='workbook-agent-scope-shared']")).toBeNull();
    expect(host.textContent).toContain("Workflows");
    expect(host.textContent).toContain("Summarize Workbook");
    expect(host.textContent).toContain("Workbook Summary");
    expect(
      fetchSpy.mock.calls.filter(
        ([input, init]) =>
          requestUrl(input).endsWith("/chat/threads") && requestMethod(init) === "GET",
      ),
    ).toHaveLength(0);

    await act(async () => {
      root.unmount();
    });
  });

  it("hides applied preview system timeline entries from the assistant rail", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({ threadId: "thr-1" }),
    );
    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: RequestInfo | URL) => {
        const url = requestUrl(input);
        if (url.endsWith("/chat/threads")) {
          return new Response(JSON.stringify([]), {
            status: 200,
            headers: { "content-type": "application/json" },
          });
        }
        if (url.endsWith("/chat/threads/thr-1")) {
          return new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: "system-apply:run-1",
                    kind: "system",
                    turnId: "turn-1",
                    text: "Applied preview bundle at revision r7: Write cells in Sheet1!B2",
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                    citations: [
                      {
                        kind: "range",
                        sheetName: "Sheet1",
                        startAddress: "B2",
                        endAddress: "B2",
                        role: "target",
                      },
                      {
                        kind: "revision",
                        revision: 7,
                      },
                    ],
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { "content-type": "application/json" },
            },
          );
        }
        throw new Error(`Unexpected fetch to ${url}`);
      }),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.textContent).not.toContain("Applied preview bundle at revision r7");
    expect(host.textContent).not.toContain("Sheet1!B2");
    expect(host.textContent).not.toContain("r7");

    await act(async () => {
      root.unmount();
    });
  });

  it("renders durable workflow runs in the assistant rail", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({ threadId: "thr-1" }),
    );
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                workflowRuns: [
                  {
                    runId: "wf-1",
                    threadId: "thr-1",
                    startedByUserId: "alex@example.com",
                    workflowTemplate: "summarizeWorkbook",
                    title: "Summarize Workbook",
                    summary: "Summarized workbook structure across 2 sheets.",
                    status: "completed",
                    createdAtUnixMs: 1,
                    updatedAtUnixMs: 2,
                    completedAtUnixMs: 2,
                    errorMessage: null,
                    steps: [
                      {
                        stepId: "inspect-workbook",
                        label: "Inspect workbook structure",
                        status: "completed",
                        summary: "Read durable workbook structure across 2 sheets.",
                        updatedAtUnixMs: 1,
                      },
                      {
                        stepId: "draft-summary",
                        label: "Draft summary artifact",
                        status: "completed",
                        summary: "Prepared the durable workbook summary artifact for the thread.",
                        updatedAtUnixMs: 2,
                      },
                    ],
                    artifact: {
                      kind: "markdown",
                      title: "Workbook Summary",
                      text: "## Workbook Summary\n\nSheets: 2\n### Sheets\n- Sheet1",
                    },
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { "content-type": "application/json" },
            },
          ),
      ),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.textContent).toContain("Workflows");
    expect(host.textContent).toContain("Summarize Workbook");
    expect(host.textContent).toContain("Inspect workbook structure");
    expect(host.textContent).toContain("Workbook Summary");
    expect(host.textContent).toContain("Sheets: 2");
    expect(host.textContent).toContain("Done");

    await act(async () => {
      root.unmount();
    });
  });

  it("loads durable thread summaries into the assistant rail", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads")) {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: "thr-shared",
              scope: "shared",
              entryCount: 4,
              hasPendingBundle: true,
              latestEntryText: "Applied preview bundle at revision r7",
            }),
            createThreadSummary({
              threadId: "thr-private",
              scope: "private",
              entryCount: 2,
              latestEntryText: "Preview bundle staged",
            }),
          ]),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-shared']")).not.toBeNull();
    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-private']")).not.toBeNull();
    expect(host.textContent).toContain("Shared");
    expect(host.textContent).toContain("Pending");
    expect(host.textContent).toContain("4 items");
    expect(host.textContent).toContain("Applied preview bundle at revision r7");

    await act(async () => {
      root.unmount();
    });
  });

  it("switches to a durable thread from the summary strip", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: "thr-2",
              scope: "shared",
              entryCount: 3,
            }),
          ]),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads/thr-2")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              scope: "shared",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const threadButton = host.querySelector("[data-testid='workbook-agent-thread-thr-2']");
    expect(threadButton instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(threadButton instanceof HTMLButtonElement)) {
        throw new Error("Thread button not found");
      }
      threadButton.click();
    });

    expect(MockEventSource.latest?.url).toBe("/v2/documents/doc-1/chat/threads/thr-2/events");
    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-2']")).toBeNull();
    expect(host.querySelector("[data-testid='workbook-agent-scope-private']")).toBeNull();
    expect(host.querySelector("[data-testid='workbook-agent-scope-shared']")).toBeNull();
    expect(fetchSpy).toHaveBeenCalledWith("/v2/documents/doc-1/chat/threads/thr-2");

    await act(async () => {
      root.unmount();
    });
  });

  it("hides the summary strip when it would only repeat the active thread", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: "thr-1",
              scope: "private",
              entryCount: 64,
              latestEntryText: "Done — prepaid expenses now exists as a sheet.",
            }),
          ]),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads/thr-1")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-1",
              scope: "private",
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-1']")).toBeNull();
    expect(host.textContent).not.toContain("64 items");

    await act(async () => {
      root.unmount();
    });
  });

  it("does not render thread scope controls", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input);
        if (url.endsWith("/chat/threads") && requestMethod(init) === "GET") {
          return new Response(JSON.stringify([]), {
            status: 200,
            headers: { "content-type": "application/json" },
          });
        }
        throw new Error(`Unexpected fetch to ${url}`);
      }),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.querySelector("[data-testid='workbook-agent-scope-private']")).toBeNull();
    expect(host.querySelector("[data-testid='workbook-agent-scope-shared']")).toBeNull();

    await act(async () => {
      root.unmount();
    });
  });

  it("restores a new-thread draft after remount", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: RequestInfo | URL) => {
        const url = requestUrl(input);
        if (url.endsWith("/chat/threads")) {
          return new Response(JSON.stringify([]), {
            status: 200,
            headers: { "content-type": "application/json" },
          });
        }
        throw new Error(`Unexpected fetch to ${url}`);
      }),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const input = host.querySelector("[data-testid='workbook-agent-input']");
    expect(input instanceof HTMLTextAreaElement).toBe(true);

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error("Agent input not found");
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      );
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, "set") : null;
      if (typeof valueSetter !== "function") {
        throw new Error("Textarea value setter not found");
      }
      Reflect.apply(valueSetter, input, ["Persisted draft"]);
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });

    await act(async () => {
      root.unmount();
    });

    const remountRoot = createRoot(host);
    await act(async () => {
      remountRoot.render(<AgentHarness />);
    });

    const restoredInput = host.querySelector("[data-testid='workbook-agent-input']");
    expect(restoredInput instanceof HTMLTextAreaElement ? restoredInput.value : null).toBe(
      "Persisted draft",
    );

    await act(async () => {
      remountRoot.unmount();
    });
  });

  it("submits the draft on Enter from the chat composer", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(JSON.stringify(createSnapshot({ entries: [] })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/turns")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: "inProgress",
              activeTurnId: "turn-1",
              entries: [
                {
                  id: "optimistic-user:turn-1",
                  kind: "user",
                  turnId: "turn-1",
                  text: "Summarize this sheet",
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const input = host.querySelector("[data-testid='workbook-agent-input']");
    expect(input instanceof HTMLTextAreaElement).toBe(true);

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error("Agent input not found");
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      );
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, "set") : null;
      if (typeof valueSetter !== "function") {
        throw new Error("Textarea value setter not found");
      }
      Reflect.apply(valueSetter, input, ["Summarize this sheet"]);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(
        new KeyboardEvent("keydown", {
          bubbles: true,
          key: "Enter",
        }),
      );
    });

    const turnCall = fetchSpy.mock.calls.find(([requestInput]) =>
      requestUrl(requestInput).endsWith("/chat/threads/thr-1/turns"),
    );
    expect(turnCall?.[0]).toBe("/v2/documents/doc-1/chat/threads/thr-1/turns");
    expect(host.textContent).not.toContain("Reviewing workbook context and drafting a response.");
    const nextInput = host.querySelector("[data-testid='workbook-agent-input']");
    expect(nextInput instanceof HTMLTextAreaElement ? nextInput.value : null).toBe("");

    await act(async () => {
      root.unmount();
    });
  });

  it("submits follow-up prompts through the durable thread route when a thread is already active", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(JSON.stringify(createSnapshot({ entries: [] })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-1/turns")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: "inProgress",
              activeTurnId: "turn-2",
              entries: [
                {
                  id: "optimistic-user:turn-2",
                  kind: "user",
                  turnId: "turn-2",
                  text: "Continue working",
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const input = host.querySelector("[data-testid='workbook-agent-input']");
    expect(input instanceof HTMLTextAreaElement).toBe(true);

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error("Agent input not found");
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      );
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, "set") : null;
      if (typeof valueSetter !== "function") {
        throw new Error("Textarea value setter not found");
      }
      Reflect.apply(valueSetter, input, ["Continue working"]);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(
        new KeyboardEvent("keydown", {
          bubbles: true,
          key: "Enter",
        }),
      );
    });

    const turnCall = fetchSpy.mock.calls.find(([requestInput]) =>
      requestUrl(requestInput).endsWith("/chat/threads/thr-1/turns"),
    );
    expect(turnCall?.[0]).toBe("/v2/documents/doc-1/chat/threads/thr-1/turns");
    expect(host.textContent).not.toContain("Reviewing workbook context and drafting a response.");

    await act(async () => {
      root.unmount();
    });
  });

  it("does not inject a synthetic progress row before the turn request resolves", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    let resolveTurnResponse: ((response: Response) => void) | null = null;
    const turnResponse = new Promise<Response>((resolve) => {
      resolveTurnResponse = resolve;
    });
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
        return new Response(JSON.stringify(createSnapshot()), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-1/turns")) {
        return await turnResponse;
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const input = host.querySelector("[data-testid='workbook-agent-input']");
    expect(input instanceof HTMLTextAreaElement).toBe(true);

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error("Agent input not found");
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      );
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, "set") : null;
      if (typeof valueSetter !== "function") {
        throw new Error("Textarea value setter not found");
      }
      Reflect.apply(valueSetter, input, ["yo"]);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(
        new KeyboardEvent("keydown", {
          bubbles: true,
          key: "Enter",
        }),
      );
      await Promise.resolve();
    });

    expect(host.textContent).toContain("yo");
    expect(host.textContent).not.toContain("Reviewing workbook context and drafting a response.");
    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).toBeNull();

    await act(async () => {
      resolveTurnResponse?.(
        new Response(
          JSON.stringify(
            createSnapshot({
              status: "inProgress",
              activeTurnId: "turn-3",
              entries: [
                {
                  id: "optimistic-user:turn-3",
                  kind: "user",
                  turnId: "turn-3",
                  text: "yo",
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        ),
      );
      await Promise.resolve();
    });

    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).not.toBeNull();
    expect(host.textContent).toContain("Thinking");

    await act(async () => {
      root.unmount();
    });
  });

  it("uses the composer button to interrupt an active turn", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: "inProgress",
              activeTurnId: "turn-1",
              entries: [
                {
                  id: "assistant-1",
                  kind: "assistant",
                  turnId: "turn-1",
                  text: "Working",
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/interrupt")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: "idle",
              activeTurnId: null,
              entries: [
                {
                  id: "assistant-1",
                  kind: "assistant",
                  turnId: "turn-1",
                  text: "Working",
                  phase: null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const button = host.querySelector("[data-testid='workbook-agent-send']");
    expect(button instanceof HTMLButtonElement).toBe(true);
    expect(button instanceof HTMLButtonElement ? button.getAttribute("aria-label") : null).toBe(
      "Stop",
    );

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Agent button not found");
      }
      button.click();
    });

    const interruptCall = fetchSpy.mock.calls.find(([input]) =>
      requestUrl(input).endsWith("/chat/threads/thr-1/interrupt"),
    );
    expect(interruptCall?.[0]).toBe("/v2/documents/doc-1/chat/threads/thr-1/interrupt");

    await act(async () => {
      root.unmount();
    });
  });

  it("renders structured workbook comprehension tool results in the rail", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: "tool-search",
                    kind: "tool",
                    turnId: "turn-1",
                    text: null,
                    phase: null,
                    toolName: "search_workbook",
                    toolStatus: "completed",
                    argumentsText: '{"query":"gross margin"}',
                    outputText: JSON.stringify({
                      query: "gross margin",
                      summary: { matchCount: 1, truncated: false },
                      matches: [
                        {
                          kind: "cell",
                          sheetName: "Sheet1",
                          address: "A2",
                          snippet: "Gross Margin",
                          reasons: ["value"],
                          score: 65,
                        },
                      ],
                    }),
                    success: true,
                  },
                  {
                    id: "tool-issues",
                    kind: "tool",
                    turnId: "turn-1",
                    text: null,
                    phase: null,
                    toolName: "find_formula_issues",
                    toolStatus: "completed",
                    argumentsText: "{}",
                    outputText: JSON.stringify({
                      summary: {
                        issueCount: 1,
                        scannedFormulaCells: 3,
                        errorCount: 1,
                        cycleCount: 0,
                        unsupportedCount: 0,
                      },
                      issues: [
                        {
                          sheetName: "Sheet1",
                          address: "C1",
                          formula: "=1/0",
                          valueText: "#DIV/0!",
                          issueKinds: ["error"],
                        },
                      ],
                    }),
                    success: true,
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { "content-type": "application/json" },
            },
          ),
      ),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(
      host.querySelector("[data-testid='workbook-agent-panel-scroll-viewport']"),
    ).not.toBeNull();
    expect(host.textContent).toContain("Search Workbook");
    expect(host.textContent).toContain("Find Formula Issues");
    expect(host.textContent).not.toContain("Gross Margin");
    expect(host.textContent).not.toContain("gross margin");
    expect(host.textContent).not.toContain("C1");

    const searchToggle = host.querySelector(
      "[data-testid='workbook-agent-tool-toggle-tool-search']",
    );
    const issuesToggle = host.querySelector(
      "[data-testid='workbook-agent-tool-toggle-tool-issues']",
    );
    expect(searchToggle instanceof HTMLButtonElement).toBe(true);
    expect(issuesToggle instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (
        !(searchToggle instanceof HTMLButtonElement) ||
        !(issuesToggle instanceof HTMLButtonElement)
      ) {
        throw new Error("Tool toggles not found");
      }
      searchToggle.click();
      issuesToggle.click();
    });

    expect(host.textContent).toContain("Gross Margin");
    expect(host.textContent).toContain("gross margin");
    expect(host.textContent).toContain("C1");

    await act(async () => {
      root.unmount();
    });
  });

  it("renders raw workbook tool payloads behind a collapsed human-readable tool row", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: "tool-read",
                    kind: "tool",
                    turnId: "turn-1",
                    text: null,
                    phase: null,
                    toolName: "read_workbook",
                    toolStatus: "completed",
                    argumentsText: JSON.stringify({
                      documentId: "bilig-demo",
                    }),
                    outputText: JSON.stringify({
                      summary: {
                        sheetCount: 2,
                        totalCellCount: 0,
                      },
                    }),
                    success: true,
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { "content-type": "application/json" },
            },
          ),
      ),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.textContent).toContain("Read Workbook");
    expect(host.textContent).not.toContain('"documentId":"bilig-demo"');
    expect(host.textContent).not.toContain('"sheetCount":2');

    const readToggle = host.querySelector("[data-testid='workbook-agent-tool-toggle-tool-read']");
    expect(readToggle instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(readToggle instanceof HTMLButtonElement)) {
        throw new Error("Read tool toggle not found");
      }
      readToggle.click();
    });

    const readPanelViewport = host.querySelector(
      "[data-testid='workbook-agent-tool-panel-tool-read-viewport']",
    );
    expect(readPanelViewport instanceof HTMLDivElement).toBe(true);
    expect(readPanelViewport?.className).toContain("h-44");
    expect(host.textContent).toContain('"documentId":"bilig-demo"');
    expect(host.textContent).toContain('"sheetCount":2');

    await act(async () => {
      root.unmount();
    });
  });

  it("summarizes attached selection ranges in tool rows", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: "tool-context",
                    kind: "tool",
                    turnId: "turn-1",
                    text: null,
                    phase: null,
                    toolName: "get_context",
                    toolStatus: "completed",
                    argumentsText: "{}",
                    outputText: JSON.stringify({
                      selection: {
                        sheetName: "Sheet1",
                        address: "E20",
                        range: {
                          startAddress: "C11",
                          endAddress: "F20",
                        },
                      },
                      visibleRange: {
                        sheetName: "Sheet1",
                        startAddress: "A1",
                        endAddress: "J20",
                      },
                    }),
                    success: true,
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { "content-type": "application/json" },
            },
          ),
      ),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.textContent).toContain("Get Context");
    expect(host.textContent).toContain("Sheet1!C11:F20");
    expect(host.textContent).not.toContain("Sheet1!E20");

    await act(async () => {
      root.unmount();
    });
  });

  it("hides raw app-server protocol errors behind user-facing copy", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response(
            JSON.stringify({
              error: "WORKBOOK_AGENT_RUNTIME_UNAVAILABLE",
              message: "thread/start.dynamicTools requires experimentalApi capability",
              retryable: true,
            }),
            {
              status: 503,
              headers: { "content-type": "application/json" },
            },
          ),
      ),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const input = host.querySelector("[data-testid='workbook-agent-input']");
    expect(input instanceof HTMLTextAreaElement).toBe(true);

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error("Agent input not found");
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      );
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, "set") : null;
      if (typeof valueSetter !== "function") {
        throw new Error("Textarea value setter not found");
      }
      Reflect.apply(valueSetter, input, ["Summarize this sheet"]);
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });

    const submit = host.querySelector("[data-testid='workbook-agent-send']");
    await act(async () => {
      if (!(submit instanceof HTMLButtonElement)) {
        throw new Error("Send button not found");
      }
      submit.click();
    });
    await flushToasts();

    expect(host.textContent).toContain("Retry in a moment.");
    expect(host.textContent).not.toContain(
      "thread/start.dynamicTools requires experimentalApi capability",
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("bootstraps the assistant session and streams assistant deltas into the rail", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input);
        if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
          return new Response(JSON.stringify(createSnapshot()), {
            status: 200,
            headers: { "content-type": "application/json" },
          });
        }
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain(
      "Thinking",
    );
    expect(MockEventSource.latest?.url).toContain("/v2/documents/doc-1/chat/threads/thr-1/events");

    await act(async () => {
      MockEventSource.latest?.emit({
        type: "entryTextDelta",
        itemId: "assistant-1",
        turnId: "turn-1",
        entryKind: "assistant",
        delta: "Updated Sheet1",
      });
    });

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain(
      "Updated Sheet1",
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("renders reasoning text immediately from streamed deltas without waiting for a snapshot refresh", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input);
        if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
          return new Response(
            JSON.stringify(
              createSnapshot({
                status: "inProgress",
                activeTurnId: "turn-1",
                entries: [
                  {
                    id: "optimistic-user:turn-1",
                    kind: "user",
                    turnId: "turn-1",
                    text: "Check version issues",
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                    citations: [],
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { "content-type": "application/json" },
            },
          );
        }
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain(
      "Thought",
    );

    await act(async () => {
      MockEventSource.latest?.emit({
        type: "entryTextDelta",
        itemId: "reasoning-1",
        turnId: "turn-1",
        entryKind: "reasoning",
        delta: "Examining version issues",
      });
    });

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain(
      "Thought",
    );
    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain(
      "Examining version issues",
    );

    await act(async () => {
      MockEventSource.latest?.emit({
        type: "entryTextDelta",
        itemId: "reasoning-1",
        turnId: "turn-1",
        entryKind: "reasoning",
        delta: " before deciding whether staged changes must be cleared.",
      });
    });

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain(
      "Examining version issues before deciding whether staged changes must be cleared.",
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("keeps the thinking row visible while tool activity is still streaming", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input);
        if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
          return new Response(
            JSON.stringify(
              createSnapshot({
                status: "inProgress",
                activeTurnId: "turn-1",
                entries: [
                  {
                    id: "optimistic-user:turn-1",
                    kind: "user",
                    turnId: "turn-1",
                    text: "Build the prepaid template",
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                    citations: [],
                  },
                  {
                    id: "tool-1",
                    kind: "tool",
                    turnId: "turn-1",
                    text: "",
                    phase: null,
                    toolName: "bilig_read_workbook",
                    toolStatus: "completed",
                    argumentsText: null,
                    outputText: null,
                    success: true,
                    citations: [],
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { "content-type": "application/json" },
            },
          );
        }
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(host.textContent).toContain("Read Workbook");
    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).not.toBeNull();
    expect(host.textContent).toContain("Thinking");

    await act(async () => {
      root.unmount();
    });
  });

  it("does not refetch thread summaries when stream snapshots arrive", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        threadId: "thr-1",
      }),
    );
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "GET") {
        return new Response(JSON.stringify([createThreadSummary({ threadId: "thr-1" })]), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
        return new Response(JSON.stringify(createSnapshot({ threadId: "thr-1" })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(
      fetchSpy.mock.calls.filter(
        ([input, init]) =>
          requestUrl(input).endsWith("/chat/threads") && requestMethod(init) === "GET",
      ),
    ).toHaveLength(1);

    await act(async () => {
      MockEventSource.latest?.emit({
        type: "snapshot",
        snapshot: createSnapshot({
          threadId: "thr-1",
          status: "inProgress",
          activeTurnId: "turn-2",
        }),
      });
    });

    expect(
      fetchSpy.mock.calls.filter(
        ([input, init]) =>
          requestUrl(input).endsWith("/chat/threads") && requestMethod(init) === "GET",
      ),
    ).toHaveLength(1);

    await act(async () => {
      root.unmount();
    });
  });

  it("recreates the assistant session and reconnects the stream after a stale session error", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );

    let resumeCount = 0;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
        resumeCount += 1;
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: resumeCount === 1 ? "agent-session-1" : "agent-session-2",
              threadId: "thr-1",
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    expect(MockEventSource.latest?.url).toContain("/v2/documents/doc-1/chat/threads/thr-1/events");

    await act(async () => {
      MockEventSource.latest?.emitError();
      await Promise.resolve();
      await Promise.resolve();
    });

    const sessionCalls = fetchSpy.mock.calls.filter(
      ([input, init]) =>
        requestUrl(input).endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET",
    );
    expect(sessionCalls).toHaveLength(2);
    expect(MockEventSource.latest?.url).toContain("/v2/documents/doc-1/chat/threads/thr-1/events");
    expect(window.sessionStorage.getItem("bilig:workbook-agent:doc-1")).toBe(
      JSON.stringify({
        threadId: "thr-1",
      }),
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("bootstraps from a stored durable thread id without requiring a stored session id", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        threadId: "thr-1",
      }),
    );
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-1",
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const bootstrapSessionCall = fetchSpy.mock.calls.find(([input, init]) => {
      return requestUrl(input).endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET";
    });
    expect(bootstrapSessionCall).toBeDefined();
    expect(MockEventSource.latest?.url).toContain("/v2/documents/doc-1/chat/threads/thr-1/events");
    expect(window.sessionStorage.getItem("bilig:workbook-agent:doc-1")).toContain(
      '"threadId":"thr-1"',
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("keeps restored private review items stable on load", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    const preview = createPreviewSummary({
      ranges: [
        {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A1",
          role: "target" as const,
        },
      ],
      cellDiffs: [
        {
          sheetName: "Sheet1",
          address: "A1",
          beforeInput: 1,
          beforeFormula: null,
          afterInput: 1,
          afterFormula: null,
          changeKinds: ["style"],
        },
      ],
      effectSummary: {
        displayedCellDiffCount: 1,
        truncatedCellDiffs: false,
        inputChangeCount: 0,
        formulaChangeCount: 0,
        styleChangeCount: 1,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    });
    const previewBundle = vi.fn(async () => preview);
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              pendingBundle: {
                id: "bundle-1",
                documentId: "doc-1",
                threadId: "thr-1",
                turnId: "turn-1",
                goalText: "Bold the selected cell",
                summary: "Format Sheet1!A1",
                scope: "selection",
                riskClass: "low",
                approvalMode: "auto",
                baseRevision: 3,
                createdAtUnixMs: 10,
                context: {
                  selection: {
                    sheetName: "Sheet1",
                    address: "A1",
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: "formatRange",
                    range: {
                      sheetName: "Sheet1",
                      startAddress: "A1",
                      endAddress: "A1",
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: "Sheet1",
                    startAddress: "A1",
                    endAddress: "A1",
                    role: "target",
                  },
                ],
                estimatedAffectedCells: 1,
              },
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      if (url.endsWith("/bundles/bundle-1/apply")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              pendingBundle: null,
              executionRecords: [
                {
                  id: "run-1",
                  bundleId: "bundle-1",
                  documentId: "doc-1",
                  threadId: "thr-1",
                  turnId: "turn-1",
                  actorUserId: "user@example.com",
                  goalText: "Bold the selected cell",
                  planText: "Apply bold formatting",
                  summary: "Format Sheet1!A1",
                  scope: "selection",
                  riskClass: "low",
                  approvalMode: "auto",
                  acceptedScope: "full",
                  appliedBy: "auto",
                  baseRevision: 3,
                  appliedRevision: 4,
                  createdAtUnixMs: 10,
                  appliedAtUnixMs: 20,
                  context: {
                    selection: {
                      sheetName: "Sheet1",
                      address: "A1",
                    },
                    viewport: {
                      rowStart: 0,
                      rowEnd: 10,
                      colStart: 0,
                      colEnd: 5,
                    },
                  },
                  commands: [
                    {
                      kind: "formatRange",
                      range: {
                        sheetName: "Sheet1",
                        startAddress: "A1",
                        endAddress: "A1",
                      },
                      patch: {
                        font: {
                          bold: true,
                        },
                      },
                    },
                  ],
                  preview,
                },
              ],
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      return new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness previewBundle={previewBundle} />);
    });

    await act(async () => {
      await Promise.resolve();
    });

    expect(previewBundle).toHaveBeenCalledTimes(1);
    const applyCall = fetchSpy.mock.calls.find(([input]) =>
      requestUrl(input).endsWith("/bundles/bundle-1/apply"),
    );
    expect(applyCall).toBeUndefined();
    expect(host.textContent).toContain("Apply");
    expect(host.textContent).not.toContain("Executions");
    expect(host.textContent).not.toContain("Replay");

    await act(async () => {
      root.unmount();
    });
  });

  it("does not auto-apply low-risk preview bundles on shared threads", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        threadId: "thr-shared",
      }),
    );
    const preview = createPreviewSummary({
      ranges: [
        {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A1",
          role: "target" as const,
        },
      ],
    });
    const previewBundle = vi.fn(async () => preview);
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads/thr-shared") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-shared",
              threadId: "thr-shared",
              scope: "shared",
              pendingBundle: {
                id: "bundle-shared-1",
                documentId: "doc-1",
                threadId: "thr-shared",
                turnId: "turn-1",
                goalText: "Bold the selected cell",
                summary: "Format Sheet1!A1",
                scope: "selection",
                riskClass: "low",
                approvalMode: "auto",
                baseRevision: 3,
                createdAtUnixMs: 10,
                context: {
                  selection: {
                    sheetName: "Sheet1",
                    address: "A1",
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: "formatRange",
                    range: {
                      sheetName: "Sheet1",
                      startAddress: "A1",
                      endAddress: "A1",
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: "Sheet1",
                    startAddress: "A1",
                    endAddress: "A1",
                    role: "target",
                  },
                ],
                estimatedAffectedCells: 1,
              },
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness previewBundle={previewBundle} />);
    });

    await act(async () => {
      await Promise.resolve();
    });

    expect(previewBundle).toHaveBeenCalled();
    expect(previewBundle.mock.calls[0]?.[0]).toMatchObject({
      id: "bundle-shared-1",
    });
    const applyCall = fetchSpy.mock.calls.find(([input]) =>
      requestUrl(input).endsWith("/bundles/bundle-shared-1/apply"),
    );
    expect(applyCall).toBeUndefined();

    await act(async () => {
      root.unmount();
    });
  });

  it("blocks collaborator approval of shared medium-risk bundles in the panel", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        threadId: "thr-shared",
      }),
    );
    const preview = createPreviewSummary({
      structuralChanges: ["Format selected range"],
    });
    const previewBundle = vi.fn(async () => preview);
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && (init?.method ?? "GET") === "GET") {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: "thr-shared",
              scope: "shared",
              ownerUserId: "alex@example.com",
              entryCount: 3,
              hasPendingBundle: true,
              latestEntryText: "Preview bundle staged",
            }),
          ]),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      if (url.endsWith("/chat/threads/thr-shared") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-shared",
              scope: "shared",
              pendingBundle: {
                id: "bundle-shared-2",
                documentId: "doc-1",
                threadId: "thr-shared",
                turnId: "turn-2",
                goalText: "Normalize the imported sheet",
                summary: "Normalize Sheet1!A1:A20",
                scope: "sheet",
                riskClass: "medium",
                approvalMode: "explicit",
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: {
                  selection: {
                    sheetName: "Sheet1",
                    address: "A1",
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: "formatRange",
                    range: {
                      sheetName: "Sheet1",
                      startAddress: "A1",
                      endAddress: "A20",
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: "Sheet1",
                    startAddress: "A1",
                    endAddress: "A20",
                    role: "target",
                  },
                ],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: "alex@example.com",
                  status: "pending",
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness currentUserId="casey@example.com" previewBundle={previewBundle} />);
    });

    await act(async () => {
      await Promise.resolve();
    });

    const applyButton = host.querySelector("[data-testid='workbook-agent-apply-pending']");
    if (!(applyButton instanceof HTMLButtonElement)) {
      throw new Error("Expected apply button to render");
    }
    expect(applyButton.disabled).toBe(true);
    expect(host.textContent).toContain(
      "Owner review routes medium/high-risk changes to Alex on this shared thread.",
    );
    expect(host.textContent).toContain("Owner review is in progress with Alex.");

    await act(async () => {
      root.unmount();
    });
  });

  it("lets the shared thread owner approve a medium-risk bundle before apply", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        threadId: "thr-shared",
      }),
    );
    const preview = createPreviewSummary({
      structuralChanges: ["Normalize selected range"],
    });
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && (init?.method ?? "GET") === "GET") {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: "thr-shared",
              scope: "shared",
              ownerUserId: "alex@example.com",
              entryCount: 3,
              hasPendingBundle: true,
              latestEntryText: "Preview bundle staged",
            }),
          ]),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      if (url.endsWith("/chat/threads/thr-shared") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-shared",
              scope: "shared",
              pendingBundle: {
                id: "bundle-shared-owner",
                documentId: "doc-1",
                threadId: "thr-shared",
                turnId: "turn-2",
                goalText: "Normalize the imported sheet",
                summary: "Normalize Sheet1!A1:A20",
                scope: "sheet",
                riskClass: "medium",
                approvalMode: "explicit",
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: {
                  selection: {
                    sheetName: "Sheet1",
                    address: "A1",
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: "formatRange",
                    range: {
                      sheetName: "Sheet1",
                      startAddress: "A1",
                      endAddress: "A20",
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: "Sheet1",
                    startAddress: "A1",
                    endAddress: "A20",
                    role: "target",
                  },
                ],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: "alex@example.com",
                  status: "pending",
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      if (url.endsWith("/bundles/bundle-shared-owner/review")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-shared",
              scope: "shared",
              pendingBundle: {
                id: "bundle-shared-owner",
                documentId: "doc-1",
                threadId: "thr-shared",
                turnId: "turn-2",
                goalText: "Normalize the imported sheet",
                summary: "Normalize Sheet1!A1:A20",
                scope: "sheet",
                riskClass: "medium",
                approvalMode: "explicit",
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: null,
                commands: [
                  {
                    kind: "formatRange",
                    range: {
                      sheetName: "Sheet1",
                      startAddress: "A1",
                      endAddress: "A20",
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: "alex@example.com",
                  status: "approved",
                  decidedByUserId: "alex@example.com",
                  decidedAtUnixMs: 25,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <AgentHarness
          currentUserId="alex@example.com"
          previewBundle={vi.fn(async () => preview)}
        />,
      );
    });

    await act(async () => {
      await Promise.resolve();
    });

    const applyButton = host.querySelector("[data-testid='workbook-agent-apply-pending']");
    const approveButton = host.querySelector("[data-testid='workbook-agent-review-approve']");
    if (!(applyButton instanceof HTMLButtonElement)) {
      throw new Error("Expected apply button");
    }
    if (!(approveButton instanceof HTMLButtonElement)) {
      throw new Error("Expected approve button");
    }
    expect(applyButton.disabled).toBe(true);

    await act(async () => {
      approveButton.click();
    });

    const reviewCall = fetchSpy.mock.calls.find(([input]) =>
      requestUrl(input).endsWith("/bundles/bundle-shared-owner/review"),
    );
    expect(requestBody(reviewCall?.[1])).toEqual({
      decision: "approved",
    });
    expect(host.textContent).toContain("Approved by Alex.");
    const refreshedApplyButton = host.querySelector("[data-testid='workbook-agent-apply-pending']");
    if (!(refreshedApplyButton instanceof HTMLButtonElement)) {
      throw new Error("Expected refreshed apply button");
    }
    expect(refreshedApplyButton.disabled).toBe(false);

    await act(async () => {
      root.unmount();
    });
  });

  it("lets collaborators recommend approval on shared medium-risk bundles", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        threadId: "thr-shared",
      }),
    );
    const preview = createPreviewSummary({
      structuralChanges: ["Normalize selected range"],
    });
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && (init?.method ?? "GET") === "GET") {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: "thr-shared",
              scope: "shared",
              ownerUserId: "alex@example.com",
              entryCount: 3,
              hasPendingBundle: true,
              latestEntryText: "Preview bundle staged",
            }),
          ]),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      if (url.endsWith("/chat/threads/thr-shared") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-shared",
              scope: "shared",
              pendingBundle: {
                id: "bundle-shared-collab",
                documentId: "doc-1",
                threadId: "thr-shared",
                turnId: "turn-2",
                goalText: "Normalize the imported sheet",
                summary: "Normalize Sheet1!A1:A20",
                scope: "sheet",
                riskClass: "medium",
                approvalMode: "explicit",
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: null,
                commands: [
                  {
                    kind: "formatRange",
                    range: {
                      sheetName: "Sheet1",
                      startAddress: "A1",
                      endAddress: "A20",
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: "alex@example.com",
                  status: "pending",
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      if (url.endsWith("/bundles/bundle-shared-collab/review")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-shared",
              scope: "shared",
              pendingBundle: {
                id: "bundle-shared-collab",
                documentId: "doc-1",
                threadId: "thr-shared",
                turnId: "turn-2",
                goalText: "Normalize the imported sheet",
                summary: "Normalize Sheet1!A1:A20",
                scope: "sheet",
                riskClass: "medium",
                approvalMode: "explicit",
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: null,
                commands: [
                  {
                    kind: "formatRange",
                    range: {
                      sheetName: "Sheet1",
                      startAddress: "A1",
                      endAddress: "A20",
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: "alex@example.com",
                  status: "pending",
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [
                    {
                      userId: "casey@example.com",
                      decision: "approved",
                      decidedAtUnixMs: 30,
                    },
                  ],
                },
              },
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      throw new Error(`Unexpected fetch to ${url}`);
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <AgentHarness currentUserId="casey@example.com" previewBundle={async () => preview} />,
      );
    });

    const approveButton = host.querySelector("[data-testid='workbook-agent-review-approve']");
    expect(approveButton instanceof HTMLButtonElement).toBe(true);
    expect(host.textContent).toContain("Owner review is in progress with Alex.");

    await act(async () => {
      if (!(approveButton instanceof HTMLButtonElement)) {
        throw new Error("Expected recommend approve button");
      }
      approveButton.click();
    });

    const reviewCall = fetchSpy.mock.calls.find(([input]) =>
      requestUrl(input).endsWith("/bundles/bundle-shared-collab/review"),
    );
    expect(requestBody(reviewCall?.[1])).toEqual({
      decision: "approved",
    });
    expect(host.textContent).toContain("1 approval recommendation");
    expect(host.textContent).toContain("You recommended approval.");

    await act(async () => {
      root.unmount();
    });
  });

  it("re-previews and applies only the selected command subset", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    window.sessionStorage.setItem(
      "bilig:workbook-agent:doc-1",
      JSON.stringify({
        sessionId: "agent-session-1",
        threadId: "thr-1",
      }),
    );
    const fullPreview = createPreviewSummary({
      ranges: [
        {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "C3",
          role: "target" as const,
        },
      ],
      effectSummary: {
        displayedCellDiffCount: 2,
        truncatedCellDiffs: false,
        inputChangeCount: 2,
        formulaChangeCount: 0,
        styleChangeCount: 0,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    });
    const subsetPreview = createPreviewSummary({
      ranges: [
        {
          sheetName: "Sheet1",
          startAddress: "C3",
          endAddress: "C3",
          role: "target" as const,
        },
      ],
      cellDiffs: [
        {
          sheetName: "Sheet1",
          address: "C3",
          beforeInput: null,
          beforeFormula: null,
          afterInput: 2,
          afterFormula: null,
          changeKinds: ["input"],
        },
      ],
      effectSummary: {
        displayedCellDiffCount: 1,
        truncatedCellDiffs: false,
        inputChangeCount: 1,
        formulaChangeCount: 0,
        styleChangeCount: 0,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    });
    const previewBundle = vi.fn(async (bundle) =>
      bundle.commands.length === 1 ? subsetPreview : fullPreview,
    );
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads/thr-1") && requestMethod(init) === "GET") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              pendingBundle: {
                id: "bundle-1",
                documentId: "doc-1",
                threadId: "thr-1",
                turnId: "turn-1",
                goalText: "Update two cells",
                summary: "Write cells in Sheet1!B2 and 1 more change",
                scope: "sheet",
                riskClass: "medium",
                approvalMode: "preview",
                baseRevision: 3,
                createdAtUnixMs: 10,
                context: {
                  selection: {
                    sheetName: "Sheet1",
                    address: "A1",
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: "writeRange",
                    sheetName: "Sheet1",
                    startAddress: "B2",
                    values: [[1]],
                  },
                  {
                    kind: "writeRange",
                    sheetName: "Sheet1",
                    startAddress: "C3",
                    values: [[2]],
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: "Sheet1",
                    startAddress: "B2",
                    endAddress: "B2",
                    role: "target",
                  },
                  {
                    sheetName: "Sheet1",
                    startAddress: "C3",
                    endAddress: "C3",
                    role: "target",
                  },
                ],
                estimatedAffectedCells: 2,
              },
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      if (url.endsWith("/bundles/bundle-1/apply")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              pendingBundle: {
                id: "bundle-2",
                documentId: "doc-1",
                threadId: "thr-1",
                turnId: "turn-1",
                goalText: "Update two cells",
                summary: "Write cells in Sheet1!B2",
                scope: "sheet",
                riskClass: "medium",
                approvalMode: "preview",
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: {
                  selection: {
                    sheetName: "Sheet1",
                    address: "A1",
                  },
                  viewport: {
                    rowStart: 0,
                    rowEnd: 10,
                    colStart: 0,
                    colEnd: 5,
                  },
                },
                commands: [
                  {
                    kind: "writeRange",
                    sheetName: "Sheet1",
                    startAddress: "B2",
                    values: [[1]],
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: "Sheet1",
                    startAddress: "B2",
                    endAddress: "B2",
                    role: "target",
                  },
                ],
                estimatedAffectedCells: 1,
              },
              executionRecords: [
                {
                  id: "run-1",
                  bundleId: "bundle-1",
                  documentId: "doc-1",
                  threadId: "thr-1",
                  turnId: "turn-1",
                  actorUserId: "user@example.com",
                  goalText: "Update two cells",
                  planText: "Apply only the second cell",
                  summary: "Write cells in Sheet1!C3",
                  scope: "sheet",
                  riskClass: "medium",
                  approvalMode: "preview",
                  acceptedScope: "partial",
                  appliedBy: "user",
                  baseRevision: 3,
                  appliedRevision: 4,
                  createdAtUnixMs: 10,
                  appliedAtUnixMs: 20,
                  context: {
                    selection: {
                      sheetName: "Sheet1",
                      address: "A1",
                    },
                    viewport: {
                      rowStart: 0,
                      rowEnd: 10,
                      colStart: 0,
                      colEnd: 5,
                    },
                  },
                  commands: [
                    {
                      kind: "writeRange",
                      sheetName: "Sheet1",
                      startAddress: "C3",
                      values: [[2]],
                    },
                  ],
                  preview: subsetPreview,
                },
              ],
            }),
          ),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      return new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness previewBundle={previewBundle} />);
    });

    expect(previewBundle).toHaveBeenCalledTimes(1);
    expect(host.textContent).toContain("2/2");

    const firstToggle = host.querySelector("[data-testid='workbook-agent-command-toggle-0']");
    expect(firstToggle instanceof HTMLInputElement).toBe(true);

    await act(async () => {
      firstToggle?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      await Promise.resolve();
    });

    expect(previewBundle).toHaveBeenCalledTimes(2);
    expect(previewBundle.mock.calls[1]?.[0]).toEqual(
      expect.objectContaining({
        commands: [
          {
            kind: "writeRange",
            sheetName: "Sheet1",
            startAddress: "C3",
            values: [[2]],
          },
        ],
      }),
    );
    expect(host.textContent).toContain("1/2");

    const applyButton = host.querySelector("[data-testid='workbook-agent-apply-pending']");
    expect(applyButton).toBeTruthy();

    await act(async () => {
      applyButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const applyCall = fetchSpy.mock.calls.find(([input]) =>
      requestUrl(input).endsWith("/bundles/bundle-1/apply"),
    );
    expect(applyCall?.[1]?.body).toBe(
      JSON.stringify({
        appliedBy: "user",
        commandIndexes: [1],
        preview: subsetPreview,
      }),
    );
    expect(host.textContent).toContain("Write cells in Sheet1!B2");
    expect(host.textContent).not.toContain("Write cells in Sheet1!C3");
    expect(host.textContent).not.toContain("Executions");
    expect(host.textContent).not.toContain("Replay");

    await act(async () => {
      root.unmount();
    });
  });
});
