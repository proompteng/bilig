import { describe, expect, it } from "vitest";
import {
  listWorkbookThreadWorkflowRuns,
  upsertWorkbookWorkflowRun,
} from "../workbook-workflow-run-store.js";
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
        return { rows: rows.filter((row): row is T => row !== null) };
      }
    }
    return { rows: [] };
  }
}

function createWorkflowRun() {
  return {
    runId: "workflow-1",
    threadId: "thr-1",
    startedByUserId: "alex@example.com",
    workflowTemplate: "summarizeWorkbook" as const,
    title: "Summarize Workbook",
    summary: "Summarized workbook structure across 2 sheets.",
    status: "completed" as const,
    createdAtUnixMs: 100,
    updatedAtUnixMs: 120,
    completedAtUnixMs: 120,
    errorMessage: null,
    artifact: {
      kind: "markdown" as const,
      title: "Workbook Summary",
      text: "## Summary",
    },
  };
}

describe("workbook-workflow-run-store", () => {
  it("persists workflow artifacts in durable rows", async () => {
    const queryable = new FakeQueryable();

    await upsertWorkbookWorkflowRun(queryable, {
      documentId: "doc-1",
      run: createWorkflowRun(),
    });

    const insertQuery = queryable.calls.find((call) =>
      call.text.includes("INSERT INTO workbook_workflow_run"),
    );
    expect(insertQuery?.values?.[12]).toBe(JSON.stringify(createWorkflowRun().artifact));
  });

  it("loads shared workflow runs for collaborator viewers", async () => {
    const run = createWorkflowRun();
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes("FROM workbook_workflow_run AS run") &&
        values?.[1] === "thr-1" &&
        values?.[2] === "casey@example.com"
          ? [
              {
                runId: run.runId,
                workbookId: "doc-1",
                threadId: run.threadId,
                actorUserId: run.startedByUserId,
                workflowTemplate: run.workflowTemplate,
                title: run.title,
                summary: run.summary,
                status: run.status,
                createdAtUnixMs: run.createdAtUnixMs,
                updatedAtUnixMs: run.updatedAtUnixMs,
                completedAtUnixMs: run.completedAtUnixMs,
                errorMessage: run.errorMessage,
                artifactJson: run.artifact,
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: "doc-1",
      actorUserId: "casey@example.com",
      threadId: "thr-1",
    });

    expect(runs).toEqual([
      expect.objectContaining({
        runId: "workflow-1",
        startedByUserId: "alex@example.com",
        workflowTemplate: "summarizeWorkbook",
      }),
    ]);
  });
});
