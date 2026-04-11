import { describe, expect, it } from "vitest";
import {
  listWorkbookAgentThreadSummaries,
  loadWorkbookAgentThreadState,
  saveWorkbookAgentThreadState,
} from "../workbook-chat-thread-store.js";
import type { QueryResultRow, Queryable } from "../store.js";

interface RecordedQuery {
  readonly text: string;
  readonly values: readonly unknown[] | undefined;
}

class FakeQueryable implements Queryable {
  readonly calls: RecordedQuery[] = [];

  constructor(
    private readonly responders: readonly ((
      text: string,
      values: readonly unknown[] | undefined,
    ) => QueryResultRow[] | null)[] = [],
  ) {}

  async query<T extends QueryResultRow = QueryResultRow>(
    text: string,
    values?: unknown[],
  ): Promise<{ rows: T[] }> {
    this.calls.push({ text, values });
    for (const responder of this.responders) {
      const rows = responder(text, values);
      if (rows) {
        return {
          rows: rows.filter((row): row is T => row !== null),
        };
      }
    }
    return { rows: [] };
  }
}

function createThreadState() {
  return {
    documentId: "doc-1",
    threadId: "thr-1",
    actorUserId: "alex@example.com",
    scope: "private" as const,
    context: {
      selection: {
        sheetName: "Sheet1",
        address: "A1",
      },
      viewport: {
        rowStart: 0,
        rowEnd: 20,
        colStart: 0,
        colEnd: 8,
      },
    },
    entries: [
      {
        id: "entry-user-1",
        kind: "user" as const,
        turnId: "turn-1",
        text: "Summarize Sheet1",
        phase: null,
        toolName: null,
        toolStatus: null,
        argumentsText: null,
        outputText: null,
        success: null,
        citations: [],
      },
      {
        id: "tool-call-1",
        kind: "tool" as const,
        turnId: "turn-1",
        text: null,
        phase: null,
        toolName: "bilig_read_workbook",
        toolStatus: "completed" as const,
        argumentsText: '{"sheetName":"Sheet1"}',
        outputText: '{"summary":"Loaded workbook"}',
        success: true,
        citations: [],
      },
      {
        id: "system-preview:bundle-1",
        kind: "system" as const,
        turnId: "turn-1",
        text: "Preview bundle staged",
        phase: null,
        toolName: null,
        toolStatus: null,
        argumentsText: null,
        outputText: null,
        success: null,
        citations: [
          {
            kind: "range" as const,
            sheetName: "Sheet1",
            startAddress: "B2",
            endAddress: "B2",
            role: "target" as const,
          },
        ],
      },
    ],
    pendingBundle: {
      id: "bundle-1",
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Normalize selection",
      summary: "Write cells in Sheet1!B2",
      scope: "selection" as const,
      riskClass: "low" as const,
      approvalMode: "preview" as const,
      baseRevision: 12,
      createdAtUnixMs: 100,
      context: {
        selection: {
          sheetName: "Sheet1",
          address: "A1",
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 8,
        },
      },
      commands: [
        {
          kind: "writeRange" as const,
          sheetName: "Sheet1",
          startAddress: "B2",
          values: [[42]],
        },
      ],
      affectedRanges: [
        {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B2",
          role: "target" as const,
        },
      ],
      estimatedAffectedCells: 1,
      sharedReview: null,
    },
    updatedAtUnixMs: 1234,
  };
}

describe("workbook-chat-thread-store", () => {
  it("persists thread metadata, timeline items, and pending bundle rows", async () => {
    const queryable = new FakeQueryable();

    await saveWorkbookAgentThreadState(queryable, createThreadState());

    expect(
      queryable.calls.some((call) => call.text.includes("INSERT INTO workbook_chat_thread")),
    ).toBe(true);
    const threadInsert = queryable.calls.find((call) =>
      call.text.includes("INSERT INTO workbook_chat_thread"),
    );
    expect(threadInsert?.values?.[5]).toBe(3);
    expect(threadInsert?.values?.[6]).toBe(true);
    expect(threadInsert?.values?.[7]).toBe("Preview bundle staged");
    expect(
      queryable.calls.some((call) => call.text.includes("DELETE FROM workbook_chat_item")),
    ).toBe(true);
    expect(
      queryable.calls.filter((call) => call.text.includes("INSERT INTO workbook_chat_item")).length,
    ).toBe(3);
    expect(
      queryable.calls.filter((call) => call.text.includes("INSERT INTO workbook_chat_tool_call"))
        .length,
    ).toBe(1);
    const bundleInsert = queryable.calls.find((call) =>
      call.text.includes("INSERT INTO workbook_pending_bundle"),
    );
    expect(bundleInsert?.values?.[3]).toBe("bundle-1");
    expect(bundleInsert?.values?.[7]).toBe("selection");
    expect(bundleInsert?.values?.[16]).toBe(JSON.stringify(null));
    const itemInsert = queryable.calls.find(
      (call) =>
        call.text.includes("INSERT INTO workbook_chat_item") &&
        call.values?.[3] === "system-preview:bundle-1",
    );
    expect(itemInsert?.values?.[14]).toBe(
      JSON.stringify([
        {
          kind: "range",
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B2",
          role: "target",
        },
      ]),
    );
    const toolInsert = queryable.calls.find(
      (call) =>
        call.text.includes("INSERT INTO workbook_chat_tool_call") &&
        call.values?.[3] === "tool-call-1",
    );
    expect(toolInsert?.values?.[6]).toBe("bilig_read_workbook");
    expect(toolInsert?.values?.[8]).toBe('{"sheetName":"Sheet1"}');
    expect(toolInsert?.values?.[9]).toBe('{"summary":"Loaded workbook"}');
  });

  it("loads a durable thread snapshot with entries and a pending bundle", async () => {
    const state = createThreadState();
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM workbook_chat_thread")
          ? [
              {
                workbookId: state.documentId,
                threadId: state.threadId,
                actorUserId: state.actorUserId,
                scope: state.scope,
                contextJson: state.context,
                updatedAtUnixMs: state.updatedAtUnixMs,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) =>
        text.includes("FROM workbook_chat_item")
          ? state.entries.map((entry, index) => ({
              entryId: entry.id,
              turnId: entry.turnId,
              kind: entry.kind,
              text: entry.text,
              phase: entry.phase,
              toolName: entry.toolName,
              toolStatus: entry.toolStatus,
              argumentsText: entry.argumentsText,
              outputText: entry.outputText,
              success: entry.success,
              citationsJson: entry.citations,
              sortOrder: index,
            }))
          : null,
      (text) =>
        text.includes("FROM workbook_chat_tool_call")
          ? state.entries
              .filter((entry) => entry.kind === "tool")
              .map((entry, index) => ({
                entryId: entry.id,
                turnId: entry.turnId,
                toolName: entry.toolName,
                toolStatus: entry.toolStatus,
                argumentsText: entry.argumentsText,
                outputText: entry.outputText,
                success: entry.success,
                sortOrder: index,
              }))
          : null,
      (text) =>
        text.includes("FROM workbook_pending_bundle")
          ? [
              {
                bundleId: state.pendingBundle?.id,
                workbookId: state.pendingBundle?.documentId,
                threadId: state.pendingBundle?.threadId,
                actorUserId: state.actorUserId,
                turnId: state.pendingBundle?.turnId,
                goalText: state.pendingBundle?.goalText,
                summary: state.pendingBundle?.summary,
                scope: state.pendingBundle?.scope,
                riskClass: state.pendingBundle?.riskClass,
                approvalMode: state.pendingBundle?.approvalMode,
                baseRevision: state.pendingBundle?.baseRevision,
                createdAtUnixMs: state.pendingBundle?.createdAtUnixMs,
                contextJson: state.pendingBundle?.context,
                commandsJson: state.pendingBundle?.commands,
                affectedRangesJson: state.pendingBundle?.affectedRanges,
                estimatedAffectedCells: state.pendingBundle?.estimatedAffectedCells,
                sharedReviewJson: state.pendingBundle?.sharedReview,
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    const loaded = await loadWorkbookAgentThreadState(queryable, {
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
    });

    expect(loaded).toEqual(state);
  });

  it("falls back to a collaborator-owned shared thread when the current user has no local row", async () => {
    const state = {
      ...createThreadState(),
      threadId: "thr-shared",
      actorUserId: "alex@example.com",
      scope: "shared" as const,
    };
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes("FROM workbook_chat_thread") && values?.[2] === "casey@example.com"
          ? [
              {
                workbookId: state.documentId,
                threadId: state.threadId,
                actorUserId: state.actorUserId,
                scope: state.scope,
                contextJson: state.context,
                updatedAtUnixMs: state.updatedAtUnixMs,
              } satisfies QueryResultRow,
            ]
          : null,
      (text, values) =>
        text.includes("FROM workbook_chat_item") && values?.[2] === "alex@example.com"
          ? state.entries.map((entry, index) => ({
              entryId: entry.id,
              turnId: entry.turnId,
              kind: entry.kind,
              text: entry.text,
              phase: entry.phase,
              toolName: entry.toolName,
              toolStatus: entry.toolStatus,
              argumentsText: entry.argumentsText,
              outputText: entry.outputText,
              success: entry.success,
              citationsJson: entry.citations,
              sortOrder: index,
            }))
          : null,
      (text, values) =>
        text.includes("FROM workbook_chat_tool_call") && values?.[2] === "alex@example.com"
          ? state.entries
              .filter((entry) => entry.kind === "tool")
              .map((entry, index) => ({
                entryId: entry.id,
                turnId: entry.turnId,
                toolName: entry.toolName,
                toolStatus: entry.toolStatus,
                argumentsText: entry.argumentsText,
                outputText: entry.outputText,
                success: entry.success,
                sortOrder: index,
              }))
          : null,
      (text, values) =>
        text.includes("FROM workbook_pending_bundle") && values?.[2] === "alex@example.com"
          ? [
              {
                bundleId: state.pendingBundle?.id,
                workbookId: state.pendingBundle?.documentId,
                threadId: state.pendingBundle?.threadId,
                actorUserId: state.actorUserId,
                turnId: state.pendingBundle?.turnId,
                goalText: state.pendingBundle?.goalText,
                summary: state.pendingBundle?.summary,
                scope: state.pendingBundle?.scope,
                riskClass: state.pendingBundle?.riskClass,
                approvalMode: state.pendingBundle?.approvalMode,
                baseRevision: state.pendingBundle?.baseRevision,
                createdAtUnixMs: state.pendingBundle?.createdAtUnixMs,
                contextJson: state.pendingBundle?.context,
                commandsJson: state.pendingBundle?.commands,
                affectedRangesJson: state.pendingBundle?.affectedRanges,
                estimatedAffectedCells: state.pendingBundle?.estimatedAffectedCells,
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    const loaded = await loadWorkbookAgentThreadState(queryable, {
      documentId: "doc-1",
      threadId: "thr-shared",
      actorUserId: "casey@example.com",
    });

    expect(loaded).toEqual(state);
  });

  it("hydrates tool call state from dedicated durable tool call rows", async () => {
    const state = createThreadState();
    const toolEntry = state.entries.find((entry) => entry.id === "tool-call-1");
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM workbook_chat_thread")
          ? [
              {
                workbookId: state.documentId,
                threadId: state.threadId,
                actorUserId: state.actorUserId,
                scope: state.scope,
                contextJson: state.context,
                updatedAtUnixMs: state.updatedAtUnixMs,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) =>
        text.includes("FROM workbook_chat_item")
          ? state.entries.map((entry, index) => ({
              entryId: entry.id,
              turnId: entry.turnId,
              kind: entry.kind,
              text: entry.text,
              phase: entry.phase,
              toolName: entry.id === "tool-call-1" ? null : entry.toolName,
              toolStatus: entry.id === "tool-call-1" ? null : entry.toolStatus,
              argumentsText: entry.id === "tool-call-1" ? null : entry.argumentsText,
              outputText: entry.id === "tool-call-1" ? null : entry.outputText,
              success: entry.id === "tool-call-1" ? null : entry.success,
              citationsJson: entry.citations,
              sortOrder: index,
            }))
          : null,
      (text) =>
        text.includes("FROM workbook_chat_tool_call") && toolEntry
          ? [
              {
                entryId: toolEntry.id,
                turnId: toolEntry.turnId,
                toolName: toolEntry.toolName,
                toolStatus: toolEntry.toolStatus,
                argumentsText: toolEntry.argumentsText,
                outputText: toolEntry.outputText,
                success: toolEntry.success,
                sortOrder: 1,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) =>
        text.includes("FROM workbook_pending_bundle")
          ? [
              {
                bundleId: state.pendingBundle?.id,
                workbookId: state.pendingBundle?.documentId,
                threadId: state.pendingBundle?.threadId,
                actorUserId: state.actorUserId,
                turnId: state.pendingBundle?.turnId,
                goalText: state.pendingBundle?.goalText,
                summary: state.pendingBundle?.summary,
                scope: state.pendingBundle?.scope,
                riskClass: state.pendingBundle?.riskClass,
                approvalMode: state.pendingBundle?.approvalMode,
                baseRevision: state.pendingBundle?.baseRevision,
                createdAtUnixMs: state.pendingBundle?.createdAtUnixMs,
                contextJson: state.pendingBundle?.context,
                commandsJson: state.pendingBundle?.commands,
                affectedRangesJson: state.pendingBundle?.affectedRanges,
                estimatedAffectedCells: state.pendingBundle?.estimatedAffectedCells,
                sharedReviewJson: state.pendingBundle?.sharedReview,
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    const loaded = await loadWorkbookAgentThreadState(queryable, {
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
    });

    expect(loaded?.entries.find((entry) => entry.id === "tool-call-1")).toEqual(toolEntry);
  });

  it("lists durable thread summaries ordered by most recent activity", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("ROW_NUMBER() OVER")
          ? [
              {
                threadId: "thr-2",
                scope: "shared",
                ownerUserId: "alex@example.com",
                updatedAtUnixMs: 200,
                entryCount: 3,
                hasPendingBundle: false,
                latestEntryText: "Applied shared cleanup at revision r7",
              } satisfies QueryResultRow,
              {
                threadId: "thr-1",
                scope: "private",
                ownerUserId: "alex@example.com",
                updatedAtUnixMs: 100,
                entryCount: 1,
                hasPendingBundle: true,
                latestEntryText: "Preview bundle staged",
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    const summaries = await listWorkbookAgentThreadSummaries(queryable, {
      documentId: "doc-1",
      actorUserId: "alex@example.com",
    });

    expect(summaries).toEqual([
      {
        threadId: "thr-2",
        scope: "shared",
        ownerUserId: "alex@example.com",
        updatedAtUnixMs: 200,
        entryCount: 3,
        hasPendingBundle: false,
        latestEntryText: "Applied shared cleanup at revision r7",
      },
      {
        threadId: "thr-1",
        scope: "private",
        ownerUserId: "alex@example.com",
        updatedAtUnixMs: 100,
        entryCount: 1,
        hasPendingBundle: true,
        latestEntryText: "Preview bundle staged",
      },
    ]);
  });

  it("includes collaborator-owned shared threads in the summary query", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("thread.scope = 'shared'")
          ? [
              {
                threadId: "thr-shared",
                scope: "shared",
                ownerUserId: "alex@example.com",
                updatedAtUnixMs: 300,
                entryCount: 2,
                hasPendingBundle: false,
                latestEntryText: "Applied preview bundle at revision r9",
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    const summaries = await listWorkbookAgentThreadSummaries(queryable, {
      documentId: "doc-1",
      actorUserId: "casey@example.com",
    });

    expect(summaries).toEqual([
      {
        threadId: "thr-shared",
        scope: "shared",
        ownerUserId: "alex@example.com",
        updatedAtUnixMs: 300,
        entryCount: 2,
        hasPendingBundle: false,
        latestEntryText: "Applied preview bundle at revision r9",
      },
    ]);
    expect(queryable.calls.at(-1)?.text).toContain(
      "AND (thread.actor_user_id = $2 OR thread.scope = 'shared')",
    );
  });
});
