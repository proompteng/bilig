// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { toast } from "sonner";
import { WorkbookToastRegion } from "../WorkbookToastRegion.js";
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

function expectWorkflowStartBody(
  init: RequestInit | undefined,
  expected: Record<string, unknown>,
): void {
  expect(requestBody(init)).toEqual({
    ...expected,
    context: createDefaultWorkflowContext(),
  });
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
    scope: "private",
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

async function openWorkflowMenu(host: HTMLElement): Promise<void> {
  const toggle = host.querySelector("[data-testid='workbook-agent-workflow-toggle']");
  expect(toggle instanceof HTMLButtonElement).toBe(true);

  await act(async () => {
    if (!(toggle instanceof HTMLButtonElement)) {
      throw new Error("Workflow menu toggle not found");
    }
    toggle.click();
  });
}

beforeEach(() => {
  vi.stubGlobal("EventSource", MockEventSource);
  window.sessionStorage.clear();
});

afterEach(() => {
  toast.dismiss();
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

    expect(host.textContent).toContain("Shared");
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

  it("renders cited system timeline entries in the assistant rail", async () => {
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

    expect(host.textContent).toContain("Applied preview bundle at revision r7");
    expect(host.textContent).toContain("Sheet1!B2");
    expect(host.textContent).toContain("r7");

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

  it("starts built-in workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads/thr-2/workflows")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "findFormulaIssues",
                  title: "Find Formula Issues",
                  summary: "Found 2 formula issues across 3 scanned formula cells.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "scan-formula-cells",
                      label: "Scan formula cells",
                      status: "completed",
                      summary: "Scanned 3 formula cells and found 2 issues.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "draft-issue-report",
                      label: "Draft issue report",
                      status: "completed",
                      summary: "Prepared the durable formula issue report for the thread.",
                      updatedAtUnixMs: 2,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Formula Issues",
                    text: "## Formula Issues",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-findFormulaIssues']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Workflow button not found");
      }
      button.click();
    });

    const createThreadCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads") && requestMethod(requestInit) === "POST",
    );
    expect(requestBody(createThreadCall?.[1])).toEqual({
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
      scope: "private",
    });
    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expect(workflowCall?.[0]).toBe("/v2/documents/doc-1/chat/threads/thr-2/workflows");
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "findFormulaIssues",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Find Formula Issues");
    expect(host.textContent).toContain("Formula Issues");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts highlight-formula workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads/thr-2/workflows")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-highlight-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "highlightFormulaIssues",
                  title: "Highlight Formula Issues",
                  summary: "Staged highlight formatting for 2 formula issues on Sheet1.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "scan-formula-cells",
                      label: "Scan formula cells",
                      status: "completed",
                      summary: "Scanned 3 formula cells on Sheet1 and found 2 issues.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "stage-issue-highlights",
                      label: "Stage issue highlights",
                      status: "completed",
                      summary:
                        "Prepared 2 semantic formatting commands to highlight the detected formula issues.",
                      updatedAtUnixMs: 2,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Formula Issue Highlights",
                    text: "## Highlighted Formula Issues",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-highlightFormulaIssues']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Highlight workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "highlightFormulaIssues",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Highlight Formula Issues");
    expect(host.textContent).toContain("Formula Issue Highlights");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts repair-formula workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads/thr-2/workflows")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-repair-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "repairFormulaIssues",
                  title: "Repair Formula Issues",
                  summary: "Staged 1 formula repair on Sheet1 from nearby healthy formulas.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "scan-formula-cells",
                      label: "Scan formula cells",
                      status: "completed",
                      summary: "Scanned 2 formula cells on Sheet1 and found 1 issue.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "stage-formula-repairs",
                      label: "Stage formula repairs",
                      status: "completed",
                      summary: "Prepared 1 semantic write command for the repair preview bundle.",
                      updatedAtUnixMs: 2,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Formula Repair Preview",
                    text: "## Formula Repair Preview",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-repairFormulaIssues']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Repair workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "repairFormulaIssues",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Repair Formula Issues");
    expect(host.textContent).toContain("Formula Repair Preview");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts header-normalization workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads/thr-2/workflows")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-header-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "normalizeCurrentSheetHeaders",
                  title: "Normalize Current Sheet Headers",
                  summary: "Staged normalized headers for 2 cells on Sheet1.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-header-row",
                      label: "Inspect header row",
                      status: "completed",
                      summary: "Loaded the used range and current header row from Sheet1.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "stage-header-normalization",
                      label: "Stage header normalization",
                      status: "completed",
                      summary:
                        "Prepared the semantic write preview that normalizes 2 header cells.",
                      updatedAtUnixMs: 2,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Header Normalization Preview",
                    text: "## Header Normalization Preview",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-normalizeCurrentSheetHeaders']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Header normalization workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "normalizeCurrentSheetHeaders",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Normalize Current Sheet Headers");
    expect(host.textContent).toContain("Header Normalization Preview");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts number-format-normalization workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads")) {
        return new Response(JSON.stringify([]), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2")) {
        return new Response(JSON.stringify(createSnapshot({ threadId: "thr-2" })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2/workflows") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-number-format-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "normalizeCurrentSheetNumberFormats",
                  title: "Normalize Current Sheet Number Formats",
                  summary: "Staged normalized number formats for 3 columns on Sheet1.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-number-columns",
                      label: "Inspect numeric columns",
                      status: "completed",
                      summary: "Loaded numeric cells and header labels from Sheet1.",
                      updatedAtUnixMs: 1,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Number Format Normalization Preview",
                    text: "## Number Format Normalization Preview",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-normalizeCurrentSheetNumberFormats']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Number format workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "normalizeCurrentSheetNumberFormats",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Normalize Current Sheet Number Formats");
    expect(host.textContent).toContain("Number Format Normalization Preview");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts whitespace-normalization workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads")) {
        return new Response(JSON.stringify([]), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2")) {
        return new Response(JSON.stringify(createSnapshot({ threadId: "thr-2" })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2/workflows") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-whitespace-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "normalizeCurrentSheetWhitespace",
                  title: "Normalize Current Sheet Whitespace",
                  summary: "Staged normalized whitespace for 3 text cells on Sheet1.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-text-cells",
                      label: "Inspect text cells",
                      status: "completed",
                      summary: "Loaded the used range and string cells from Sheet1.",
                      updatedAtUnixMs: 1,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Whitespace Normalization Preview",
                    text: "## Whitespace Normalization Preview",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-normalizeCurrentSheetWhitespace']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Whitespace workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "normalizeCurrentSheetWhitespace",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Normalize Current Sheet Whitespace");
    expect(host.textContent).toContain("Whitespace Normalization Preview");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts formula fill-down workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads")) {
        return new Response(JSON.stringify([]), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2")) {
        return new Response(JSON.stringify(createSnapshot({ threadId: "thr-2" })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2/workflows") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-fill-formulas-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "fillCurrentSheetFormulasDown",
                  title: "Fill Current Sheet Formulas Down",
                  summary: "Staged formula fill-down for 1 column on Sheet1.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-formula-columns",
                      label: "Inspect formula columns",
                      status: "completed",
                      summary: "Loaded formula cells and blank fill gaps from Sheet1.",
                      updatedAtUnixMs: 1,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Formula Fill-Down Preview",
                    text: "## Formula Fill-Down Preview",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-fillCurrentSheetFormulasDown']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Formula fill-down workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "fillCurrentSheetFormulasDown",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Fill Current Sheet Formulas Down");
    expect(host.textContent).toContain("Formula Fill-Down Preview");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts header-style workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads")) {
        return new Response(JSON.stringify([]), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2")) {
        return new Response(JSON.stringify(createSnapshot({ threadId: "thr-2" })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2/workflows") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-style-headers-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "styleCurrentSheetHeaders",
                  title: "Style Current Sheet Headers",
                  summary: "Staged a consistent header style preview for Sheet1.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-header-row",
                      label: "Inspect header row",
                      status: "completed",
                      summary: "Loaded the used range and header row from Sheet1.",
                      updatedAtUnixMs: 1,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Header Style Preview",
                    text: "## Header Style Preview",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-styleCurrentSheetHeaders']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Header style workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "styleCurrentSheetHeaders",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Style Current Sheet Headers");
    expect(host.textContent).toContain("Header Style Preview");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts current-sheet review-tab workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads")) {
        return new Response(JSON.stringify([]), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2")) {
        return new Response(JSON.stringify(createSnapshot({ threadId: "thr-2" })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2/workflows") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-review-tab-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "createCurrentSheetReviewTab",
                  title: "Create Current Sheet Review Tab",
                  summary: "Staged a review-tab preview for Sheet1 into Sheet1 Review.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-source-sheet",
                      label: "Inspect source sheet",
                      status: "completed",
                      summary: "Loaded the used range from Sheet1.",
                      updatedAtUnixMs: 1,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Current Sheet Review Tab Preview",
                    text: "## Current Sheet Review Tab Preview",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-createCurrentSheetReviewTab']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Review-tab workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "createCurrentSheetReviewTab",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Create Current Sheet Review Tab");
    expect(host.textContent).toContain("Current Sheet Review Tab Preview");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts current-sheet rollup workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads")) {
        return new Response(JSON.stringify([]), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2")) {
        return new Response(JSON.stringify(createSnapshot({ threadId: "thr-2" })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2/workflows") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-rollup-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "createCurrentSheetRollup",
                  title: "Create Current Sheet Rollup",
                  summary: "Staged a rollup preview for Sheet1 into Sheet1 Rollup.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-source-sheet",
                      label: "Inspect source sheet",
                      status: "completed",
                      summary: "Loaded the used range and numeric columns from Sheet1.",
                      updatedAtUnixMs: 1,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Current Sheet Rollup Preview",
                    text: "## Current Sheet Rollup Preview",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-createCurrentSheetRollup']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Rollup workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(([requestInput]) =>
      requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows"),
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "createCurrentSheetRollup",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Create Current Sheet Rollup");
    expect(host.textContent).toContain("Current Sheet Rollup Preview");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts outlier-highlight workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads")) {
        return new Response(JSON.stringify([]), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2")) {
        return new Response(JSON.stringify(createSnapshot({ threadId: "thr-2" })), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads/thr-2/workflows") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-outlier-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "highlightCurrentSheetOutliers",
                  title: "Highlight Current Sheet Outliers",
                  summary:
                    "Staged outlier highlights for 1 cell across 1 numeric column on Sheet1.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-numeric-columns",
                      label: "Inspect numeric columns",
                      status: "completed",
                      summary: "Loaded numeric cells and header labels from Sheet1.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "stage-outlier-highlights",
                      label: "Stage outlier highlights",
                      status: "completed",
                      summary:
                        "Prepared 1 semantic formatting command to highlight numeric outliers on Sheet1.",
                      updatedAtUnixMs: 2,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Current Sheet Outlier Highlights",
                    text: "## Highlighted Numeric Outliers",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-highlightCurrentSheetOutliers']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Outlier workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows") &&
        requestMethod(requestInit) === "POST",
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "highlightCurrentSheetOutliers",
      sheetName: "Sheet1",
    });
    expect(host.textContent).toContain("Highlight Current Sheet Outliers");
    expect(host.textContent).toContain("Current Sheet Outlier Highlights");

    await act(async () => {
      root.unmount();
    });
  });

  it("cancels running workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    sessionStorage.setItem("bilig:workbook-agent:doc-1", JSON.stringify({ threadId: "thr-1" }));
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, _init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads/thr-1/workflows/wf-running-1/cancel")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: "thr-1",
              workflowRuns: [
                {
                  runId: "wf-running-1",
                  threadId: "thr-1",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "summarizeWorkbook",
                  title: "Summarize Workbook",
                  summary: "Cancelled workflow: Summarize Workbook",
                  status: "cancelled",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 4,
                  completedAtUnixMs: 4,
                  errorMessage: "Cancelled by alex@example.com.",
                  steps: [
                    {
                      stepId: "inspect-workbook",
                      label: "Inspect workbook structure",
                      status: "cancelled",
                      summary: "Workflow cancelled before this step completed.",
                      updatedAtUnixMs: 4,
                    },
                  ],
                  artifact: null,
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
        JSON.stringify(
          createSnapshot({
            threadId: "thr-1",
            workflowRuns: [
              {
                runId: "wf-running-1",
                threadId: "thr-1",
                startedByUserId: "alex@example.com",
                workflowTemplate: "summarizeWorkbook",
                title: "Summarize Workbook",
                summary: "Running workbook summary workflow.",
                status: "running",
                createdAtUnixMs: 1,
                updatedAtUnixMs: 2,
                completedAtUnixMs: null,
                errorMessage: null,
                steps: [
                  {
                    stepId: "inspect-workbook",
                    label: "Inspect workbook structure",
                    status: "running",
                    summary: "Reading durable workbook structure and layout metadata.",
                    updatedAtUnixMs: 2,
                  },
                  {
                    stepId: "draft-summary",
                    label: "Draft summary artifact",
                    status: "pending",
                    summary: "Waiting to assemble the durable workbook summary artifact.",
                    updatedAtUnixMs: 2,
                  },
                ],
                artifact: null,
              },
            ],
          }),
        ),
        {
          status: 200,
          headers: { "content-type": "application/json" },
        },
      );
    });
    vi.stubGlobal("fetch", fetchSpy);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    const button = document.querySelector(
      "[data-testid='workbook-agent-cancel-workflow-wf-running-1']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);
    expect(host.textContent).toContain("Running");

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Cancel workflow button not found");
      }
      button.click();
    });

    const cancelCall = fetchSpy.mock.calls.find(([requestInput]) =>
      requestUrl(requestInput).endsWith("/chat/threads/thr-1/workflows/wf-running-1/cancel"),
    );
    expect(cancelCall?.[0]).toBe(
      "/v2/documents/doc-1/chat/threads/thr-1/workflows/wf-running-1/cancel",
    );
    expect(requestMethod(cancelCall?.[1])).toBe("POST");
    expect(host.textContent).toContain("Cancelled");
    expect(host.textContent).toContain("Cancelled by alex@example.com.");
    expect(
      host.querySelector("[data-testid='workbook-agent-cancel-workflow-wf-running-1']"),
    ).toBeNull();

    await act(async () => {
      root.unmount();
    });
  });

  it("starts dependency trace workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads/thr-2/workflows")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-2",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "traceSelectionDependencies",
                  title: "Trace Selection Dependencies",
                  summary: "Traced 1 precedent and 1 dependent from Sheet1!B1.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-selection",
                      label: "Inspect current selection",
                      status: "completed",
                      summary: "Loaded workbook context for Sheet1!B1.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "trace-links",
                      label: "Trace workbook links",
                      status: "completed",
                      summary: "Traced 1 precedent and 1 dependent.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "draft-trace-report",
                      label: "Draft trace report",
                      status: "completed",
                      summary: "Prepared the durable dependency trace report for the thread.",
                      updatedAtUnixMs: 2,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Dependency Trace",
                    text: "## Dependency Trace\n\nRoot: Sheet1!B1",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-traceSelectionDependencies']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(([requestInput]) =>
      requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows"),
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "traceSelectionDependencies",
    });
    expect(host.textContent).toContain("Trace Selection Dependencies");
    expect(host.textContent).toContain("Dependency Trace");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts current-sheet summary workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads/thr-2/workflows")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-sheet-1",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "summarizeCurrentSheet",
                  title: "Summarize Current Sheet",
                  summary: "Summarized Sheet1 with 12 populated cells and 1 table.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-current-sheet",
                      label: "Inspect current sheet",
                      status: "completed",
                      summary:
                        "Read durable metadata for Sheet1, including used range, tables, pivots, spills, and axis metadata.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "draft-sheet-summary",
                      label: "Draft current sheet summary",
                      status: "completed",
                      summary:
                        "Prepared the durable current-sheet summary artifact for the thread.",
                      updatedAtUnixMs: 2,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Current Sheet Summary",
                    text: "## Current Sheet Summary\n\nSheet: Sheet1",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-summarizeCurrentSheet']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(([requestInput]) =>
      requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows"),
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "summarizeCurrentSheet",
    });
    expect(host.textContent).toContain("Summarize Current Sheet");
    expect(host.textContent).toContain("Current Sheet Summary");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts current-cell explanation workflows through the durable thread workflow route", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/chat/threads/thr-2/workflows")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-2",
              threadId: "thr-2",
              workflowRuns: [
                {
                  runId: "wf-3",
                  threadId: "thr-2",
                  startedByUserId: "alex@example.com",
                  workflowTemplate: "explainSelectionCell",
                  title: "Explain Current Cell",
                  summary: "Explained Sheet1!A1, including direct precedents and dependents.",
                  status: "completed",
                  createdAtUnixMs: 1,
                  updatedAtUnixMs: 2,
                  completedAtUnixMs: 2,
                  errorMessage: null,
                  steps: [
                    {
                      stepId: "inspect-selection",
                      label: "Inspect current selection",
                      status: "completed",
                      summary: "Loaded workbook context for Sheet1!A1.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "explain-cell",
                      label: "Explain current cell",
                      status: "completed",
                      summary: "Read the current value and workbook links for Sheet1!A1.",
                      updatedAtUnixMs: 1,
                    },
                    {
                      stepId: "draft-explanation",
                      label: "Draft explanation artifact",
                      status: "completed",
                      summary:
                        "Prepared the durable current-cell explanation artifact for the thread.",
                      updatedAtUnixMs: 2,
                    },
                  ],
                  artifact: {
                    kind: "markdown",
                    title: "Current Cell",
                    text: "## Current Cell\n\nCell: Sheet1!A1",
                  },
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

    await openWorkflowMenu(host);

    const button = document.querySelector(
      "[data-testid='workbook-agent-workflow-start-explainSelectionCell']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Workflow button not found");
      }
      button.click();
    });

    const workflowCall = fetchSpy.mock.calls.find(([requestInput]) =>
      requestUrl(requestInput).endsWith("/chat/threads/thr-2/workflows"),
    );
    expectWorkflowStartBody(workflowCall?.[1], {
      workflowTemplate: "explainSelectionCell",
    });
    expect(host.textContent).toContain("Explain Current Cell");
    expect(host.textContent).toContain("Current Cell");

    await act(async () => {
      root.unmount();
    });
  });

  it("does not expose workbook search controls in the tools dropdown", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AgentHarness />);
    });

    await openWorkflowMenu(host);

    expect(
      document.querySelector("[data-testid='workbook-agent-workflow-search-input']"),
    ).toBeNull();
    expect(
      document.querySelector("[data-testid='workbook-agent-workflow-start-searchWorkbookQuery']"),
    ).toBeNull();
    expect(host.textContent).not.toContain("Search workbook");

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
    expect(
      host
        .querySelector("[data-testid='workbook-agent-scope-shared']")
        ?.getAttribute("aria-pressed"),
    ).toBe("true");
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

  it("starts a shared thread when the shared scope is selected", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "GET") {
        return new Response(JSON.stringify([]), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-shared",
              threadId: "thr-shared",
              entries: [],
            }),
          ),
          {
            status: 200,
            headers: { "content-type": "application/json" },
          },
        );
      }
      if (url.endsWith("/turns")) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              sessionId: "agent-session-shared",
              threadId: "thr-shared",
              status: "inProgress",
              activeTurnId: "turn-1",
              entries: [
                {
                  id: "optimistic-user:turn-1",
                  kind: "user",
                  turnId: "turn-1",
                  text: "Share this thread",
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

    const sharedScopeButton = host.querySelector("[data-testid='workbook-agent-scope-shared']");
    expect(sharedScopeButton instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(sharedScopeButton instanceof HTMLButtonElement)) {
        throw new Error("Shared scope button not found");
      }
      sharedScopeButton.click();
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
      Reflect.apply(valueSetter, input, ["Share this thread"]);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(
        new KeyboardEvent("keydown", {
          bubbles: true,
          key: "Enter",
        }),
      );
    });

    expect(MockEventSource.latest?.url).toBe("/v2/documents/doc-1/chat/threads/thr-shared/events");
    const sessionCall = fetchSpy.mock.calls.find(
      ([requestInput, requestInit]) =>
        requestUrl(requestInput).endsWith("/chat/threads") &&
        typeof requestBody(requestInit) === "object" &&
        requestBody(requestInit) !== null &&
        "scope" in requestBody(requestInit) &&
        requestBody(requestInit)["scope"] === "shared",
    );
    expect(requestBody(sessionCall?.[1])).toEqual({
      scope: "shared",
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
    });
    expect(
      fetchSpy.mock.calls.some(([requestInput]) =>
        requestUrl(requestInput).endsWith("/chat/threads/thr-shared/turns"),
      ),
    ).toBe(true);

    await act(async () => {
      root.unmount();
    });
  });

  it("keeps separate drafts for new private and shared threads", async () => {
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
    const privateScopeButton = host.querySelector("[data-testid='workbook-agent-scope-private']");
    const sharedScopeButton = host.querySelector("[data-testid='workbook-agent-scope-shared']");
    expect(input instanceof HTMLTextAreaElement).toBe(true);
    expect(privateScopeButton instanceof HTMLButtonElement).toBe(true);
    expect(sharedScopeButton instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (
        !(input instanceof HTMLTextAreaElement) ||
        !(sharedScopeButton instanceof HTMLButtonElement) ||
        !(privateScopeButton instanceof HTMLButtonElement)
      ) {
        throw new Error("Agent controls not found");
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      );
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, "set") : null;
      if (typeof valueSetter !== "function") {
        throw new Error("Textarea value setter not found");
      }
      Reflect.apply(valueSetter, input, ["Private draft"]);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      sharedScopeButton.click();
    });

    const sharedInput = host.querySelector("[data-testid='workbook-agent-input']");
    expect(sharedInput instanceof HTMLTextAreaElement ? sharedInput.value : null).toBe("");

    await act(async () => {
      if (
        !(sharedInput instanceof HTMLTextAreaElement) ||
        !(sharedScopeButton instanceof HTMLButtonElement) ||
        !(privateScopeButton instanceof HTMLButtonElement)
      ) {
        throw new Error("Agent controls not found");
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      );
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, "set") : null;
      if (typeof valueSetter !== "function") {
        throw new Error("Textarea value setter not found");
      }
      Reflect.apply(valueSetter, sharedInput, ["Shared draft"]);
      sharedInput.dispatchEvent(new Event("input", { bubbles: true }));
      privateScopeButton.click();
    });

    const restoredPrivateInput = host.querySelector("[data-testid='workbook-agent-input']");
    expect(
      restoredPrivateInput instanceof HTMLTextAreaElement ? restoredPrivateInput.value : null,
    ).toBe("Private draft");

    await act(async () => {
      if (!(sharedScopeButton instanceof HTMLButtonElement)) {
        throw new Error("Shared scope button not found");
      }
      sharedScopeButton.click();
    });

    const restoredSharedInput = host.querySelector("[data-testid='workbook-agent-input']");
    expect(
      restoredSharedInput instanceof HTMLTextAreaElement ? restoredSharedInput.value : null,
    ).toBe("Shared draft");

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
    const sharedScopeButton = host.querySelector("[data-testid='workbook-agent-scope-shared']");
    expect(input instanceof HTMLTextAreaElement).toBe(true);
    expect(sharedScopeButton instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(sharedScopeButton instanceof HTMLButtonElement)) {
        throw new Error("Agent controls not found");
      }
      sharedScopeButton.click();
    });

    const sharedInput = host.querySelector("[data-testid='workbook-agent-input']");
    expect(sharedInput instanceof HTMLTextAreaElement).toBe(true);

    await act(async () => {
      if (!(sharedInput instanceof HTMLTextAreaElement)) {
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
      Reflect.apply(valueSetter, sharedInput, ["Persisted shared draft"]);
      sharedInput.dispatchEvent(new Event("input", { bubbles: true }));
    });

    await act(async () => {
      root.unmount();
    });

    const remountRoot = createRoot(host);
    await act(async () => {
      remountRoot.render(<AgentHarness />);
    });

    const remountedSharedScopeButton = host.querySelector(
      "[data-testid='workbook-agent-scope-shared']",
    );
    expect(remountedSharedScopeButton instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(remountedSharedScopeButton instanceof HTMLButtonElement)) {
        throw new Error("Shared scope button not found");
      }
      remountedSharedScopeButton.click();
    });

    const restoredInput = host.querySelector("[data-testid='workbook-agent-input']");
    expect(restoredInput instanceof HTMLTextAreaElement ? restoredInput.value : null).toBe(
      "Persisted shared draft",
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
    const nextInput = host.querySelector("[data-testid='workbook-agent-input']");
    expect(nextInput instanceof HTMLTextAreaElement ? nextInput.value : null).toBe("");
    expect(host.textContent).toContain("Reviewing workbook context and drafting a response.");

    await act(async () => {
      root.unmount();
    });
  });

  it("shows immediate non-spinner progress before the turn request resolves", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    let resolveTurnResponse: ((response: Response) => void) | null = null;
    const turnResponse = new Promise<Response>((resolve) => {
      resolveTurnResponse = resolve;
    });
    const fetchSpy = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input);
      if (url.endsWith("/chat/threads") && requestMethod(init) === "POST") {
        return Promise.resolve(
          new Response(JSON.stringify(createSnapshot({ entries: [] })), {
            status: 200,
            headers: { "content-type": "application/json" },
          }),
        );
      }
      if (url.endsWith("/chat/threads/thr-1/turns")) {
        return turnResponse;
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
    expect(host.textContent).toContain("Reviewing workbook context and drafting a response.");
    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).not.toBeNull();

    await act(async () => {
      resolveTurnResponse?.(
        new Response(
          JSON.stringify(
            createSnapshot({
              status: "inProgress",
              activeTurnId: "turn-1",
              entries: [
                {
                  id: "optimistic-user:turn-1",
                  kind: "user",
                  turnId: "turn-1",
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
      await turnResponse;
    });

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
    expect(host.textContent).toContain("Reviewing workbook context and drafting a response.");

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

    expect(host.textContent).toContain('"documentId":"bilig-demo"');
    expect(host.textContent).toContain('"sheetCount":2');

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
    expect(applyCall).toBeTruthy();
    expect(applyCall?.[1]?.body).toBe(
      JSON.stringify({
        appliedBy: "auto",
        commandIndexes: [0],
        preview,
      }),
    );
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
      "Only Alex can approve medium/high-risk changes on this shared thread.",
    );
    expect(host.textContent).toContain("Awaiting Alex's approval before this shared bundle");

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
    expect(host.textContent).toContain("Awaiting Alex's approval before this shared bundle");

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
