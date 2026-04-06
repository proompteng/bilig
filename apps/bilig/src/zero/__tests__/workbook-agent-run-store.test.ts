import { describe, expect, it } from "vitest";
import { appendWorkbookAgentRun, listWorkbookAgentRuns } from "../workbook-agent-run-store.js";
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

function createExecutionRecord() {
  return {
    id: "run-1",
    bundleId: "bundle-1",
    documentId: "doc-1",
    threadId: "thr-1",
    turnId: "turn-1",
    actorUserId: "alex@example.com",
    goalText: "Apply only the selected command",
    planText: "Apply the second command only",
    summary: "Write cells in Sheet1!C3",
    scope: "sheet" as const,
    riskClass: "medium" as const,
    approvalMode: "preview" as const,
    acceptedScope: "partial" as const,
    appliedBy: "user" as const,
    baseRevision: 3,
    appliedRevision: 4,
    createdAtUnixMs: 100,
    appliedAtUnixMs: 200,
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
        kind: "writeRange" as const,
        sheetName: "Sheet1",
        startAddress: "C3",
        values: [[2]],
      },
    ],
    preview: null,
  };
}

describe("workbook-agent-run-store", () => {
  it("persists partial accepted scope in execution rows", async () => {
    const queryable = new FakeQueryable();

    await appendWorkbookAgentRun(queryable, createExecutionRecord());

    const insertQuery = queryable.calls.find((call) =>
      call.text.includes("INSERT INTO workbook_agent_run"),
    );
    expect(insertQuery?.values?.[12]).toBe("partial");
  });

  it("loads partial execution records from stored rows", async () => {
    const record = createExecutionRecord();
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM workbook_agent_run")
          ? [
              {
                id: record.id,
                bundleId: record.bundleId,
                workbookId: record.documentId,
                threadId: record.threadId,
                turnId: record.turnId,
                actorUserId: record.actorUserId,
                goalText: record.goalText,
                planText: record.planText,
                summary: record.summary,
                scope: record.scope,
                riskClass: record.riskClass,
                approvalMode: record.approvalMode,
                acceptedScope: record.acceptedScope,
                appliedBy: record.appliedBy,
                baseRevision: record.baseRevision,
                appliedRevision: record.appliedRevision,
                createdAtUnixMs: record.createdAtUnixMs,
                appliedAtUnixMs: record.appliedAtUnixMs,
                contextJson: record.context,
                commandsJson: record.commands,
                previewJson: record.preview,
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    const records = await listWorkbookAgentRuns(queryable, {
      documentId: "doc-1",
      actorUserId: "alex@example.com",
    });

    expect(records).toEqual([
      expect.objectContaining({
        acceptedScope: "partial",
        commands: [
          {
            kind: "writeRange",
            sheetName: "Sheet1",
            startAddress: "C3",
            values: [[2]],
          },
        ],
      }),
    ]);
  });
});
