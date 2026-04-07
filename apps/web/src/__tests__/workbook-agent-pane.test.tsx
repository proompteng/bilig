// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WorkbookToastRegion } from "../WorkbookToastRegion.js";
import { useWorkbookAgentPane } from "../use-workbook-agent-pane.js";

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

function createSnapshot(overrides: Record<string, unknown> = {}) {
  return {
    sessionId: "agent-session-1",
    documentId: "doc-1",
    threadId: "thr-1",
    status: "idle",
    activeTurnId: null,
    lastError: null,
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
      },
    ],
    pendingBundle: null,
    executionRecords: [],
    ...overrides,
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

function AgentHarness(props: {
  readonly previewBundle?: Parameters<typeof useWorkbookAgentPane>[0]["previewBundle"];
}) {
  const { agentError, agentPanel, clearAgentError } = useWorkbookAgentPane({
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
});

afterEach(() => {
  vi.restoreAllMocks();
  window.sessionStorage.clear();
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
    expect(host.textContent).toContain("Sheet1!A1");
    expect(input instanceof HTMLTextAreaElement ? input.getAttribute("placeholder") : null).toBe(
      "",
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("submits the draft on Enter from the chat composer", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input);
      if (url.endsWith("/agent/sessions")) {
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

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(fetchSpy.mock.calls[1]?.[0]).toBe(
      "/v2/documents/doc-1/agent/sessions/agent-session-1/turns",
    );
    const nextInput = host.querySelector("[data-testid='workbook-agent-input']");
    expect(nextInput instanceof HTMLTextAreaElement ? nextInput.value : null).toBe("");

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
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input);
      if (url.endsWith("/agent/sessions")) {
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

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(fetchSpy.mock.calls[1]?.[0]).toBe(
      "/v2/documents/doc-1/agent/sessions/agent-session-1/interrupt",
    );

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
                    toolName: "bilig_search_workbook",
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
                    toolName: "bilig_find_formula_issues",
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

    expect(host.textContent).toContain("Search Matches");
    expect(host.textContent).toContain("Gross Margin");
    expect(host.textContent).toContain("Formula Issues");
    expect(host.textContent).toContain("C1");

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
      vi.fn(async (input: RequestInfo | URL) => {
        const url = requestUrl(input);
        if (url.endsWith("/agent/sessions")) {
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
    expect(MockEventSource.latest?.url).toContain(
      "/v2/documents/doc-1/agent/sessions/agent-session-1/events",
    );

    await act(async () => {
      MockEventSource.latest?.emit({
        type: "assistantDelta",
        itemId: "assistant-1",
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
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input);
      if (url.endsWith("/agent/sessions")) {
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

    expect(MockEventSource.latest?.url).toContain(
      "/v2/documents/doc-1/agent/sessions/agent-session-1/events",
    );

    await act(async () => {
      MockEventSource.latest?.emitError();
      await Promise.resolve();
      await Promise.resolve();
    });

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(MockEventSource.latest?.url).toContain(
      "/v2/documents/doc-1/agent/sessions/agent-session-2/events",
    );
    expect(window.sessionStorage.getItem("bilig:workbook-agent:doc-1")).toContain(
      '"sessionId":"agent-session-2"',
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("auto-applies low-risk preview bundles after the local preview resolves", async () => {
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
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input);
      if (url.endsWith("/agent/sessions")) {
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
    expect(applyCall).toBeTruthy();
    expect(applyCall?.[1]?.body).toBe(
      JSON.stringify({
        appliedBy: "auto",
        commandIndexes: [0],
        preview,
      }),
    );
    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain("r4");

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
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input);
      if (url.endsWith("/agent/sessions")) {
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
    expect(host.textContent).toContain("Write cells in Sheet1!C3");

    await act(async () => {
      root.unmount();
    });
  });
});
