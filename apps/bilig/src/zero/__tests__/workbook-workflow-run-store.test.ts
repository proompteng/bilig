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
    steps: [
      {
        stepId: "inspect-workbook",
        label: "Inspect workbook structure",
        status: "completed" as const,
        summary: "Read durable workbook structure across 2 sheets.",
        updatedAtUnixMs: 110,
      },
      {
        stepId: "draft-summary",
        label: "Draft summary artifact",
        status: "completed" as const,
        summary: "Prepared the durable workbook summary artifact for the thread.",
        updatedAtUnixMs: 120,
      },
    ],
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
    expect(insertQuery?.values?.[12]).toBe(JSON.stringify(createWorkflowRun().steps));
    expect(insertQuery?.values?.[13]).toBe(JSON.stringify(createWorkflowRun().artifact));
    expect(
      queryable.calls.find((call) => call.text.includes("DELETE FROM workbook_workflow_step")),
    ).toBeDefined();
    expect(
      queryable.calls.filter((call) => call.text.includes("INSERT INTO workbook_workflow_step")),
    ).toHaveLength(createWorkflowRun().steps.length);
    expect(
      queryable.calls.find((call) => call.text.includes("INSERT INTO workbook_workflow_artifact")),
    ).toBeDefined();
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
                stepsJson: run.steps,
                artifactJson: run.artifact,
              } satisfies QueryResultRow,
            ]
          : null,
      (text, values) =>
        text.includes("FROM workbook_workflow_step AS step") &&
        values?.[0] === "doc-1" &&
        Array.isArray(values?.[1])
          ? run.steps.map(
              (step, index) =>
                ({
                  runId: run.runId,
                  stepId: step.stepId,
                  stepOrder: index,
                  label: step.label,
                  status: step.status,
                  summary: step.summary,
                  updatedAtUnixMs: step.updatedAtUnixMs,
                }) satisfies QueryResultRow,
            )
          : null,
      (text, values) =>
        text.includes("FROM workbook_workflow_artifact AS artifact") &&
        values?.[0] === "doc-1" &&
        Array.isArray(values?.[1])
          ? [
              {
                runId: run.runId,
                kind: run.artifact?.kind,
                title: run.artifact?.title,
                text: run.artifact?.text,
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
        steps: expect.arrayContaining([
          expect.objectContaining({
            stepId: "inspect-workbook",
            status: "completed",
          }),
        ]),
        artifact: expect.objectContaining({
          title: "Workbook Summary",
        }),
      }),
    ]);
  });
});
