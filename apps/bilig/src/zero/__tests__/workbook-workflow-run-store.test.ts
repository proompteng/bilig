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
    expect(queryable.calls[0]?.text).toContain("EXISTS (");
    expect(queryable.calls[0]?.text).not.toContain("thread.actor_user_id = run.actor_user_id");
  });

  it("hydrates structural workflow templates from durable rows after reload", async () => {
    const run = {
      ...createWorkflowRun(),
      runId: "workflow-structural-1",
      workflowTemplate: "hideCurrentRow" as const,
      title: "Hide Current Row",
      summary: "Staged a structural change set to hide row 7 on Sheet2.",
      steps: [
        {
          stepId: "resolve-current-row",
          label: "Resolve current row",
          status: "completed" as const,
          summary: "Resolved the selected row as row 7 on Sheet2.",
          updatedAtUnixMs: 110,
        },
        {
          stepId: "stage-row-visibility-preview",
          label: "Stage row visibility preview",
          status: "completed" as const,
          summary: "Staged the semantic preview that hides the current row.",
          updatedAtUnixMs: 120,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Hide Row Preview",
        text: "## Hide Row Preview",
      },
    };
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes("FROM workbook_workflow_run AS run") &&
        values?.[1] === "thr-1" &&
        values?.[2] === "alex@example.com"
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
      actorUserId: "alex@example.com",
      threadId: "thr-1",
    });

    expect(runs).toEqual([
      expect.objectContaining({
        runId: "workflow-structural-1",
        workflowTemplate: "hideCurrentRow",
        title: "Hide Current Row",
        artifact: expect.objectContaining({
          title: "Hide Row Preview",
        }),
        steps: expect.arrayContaining([
          expect.objectContaining({
            stepId: "resolve-current-row",
          }),
          expect.objectContaining({
            stepId: "stage-row-visibility-preview",
          }),
        ]),
      }),
    ]);
  });

  it("hydrates newly added durable workflow templates from rows after reload", async () => {
    const run = {
      ...createWorkflowRun(),
      runId: "workflow-import-1",
      workflowTemplate: "normalizeCurrentSheetNumberFormats" as const,
      title: "Normalize Current Sheet Number Formats",
      summary: "Staged normalized number formats for 3 columns on Imports.",
      artifact: {
        kind: "markdown" as const,
        title: "Number Format Normalization Preview",
        text: "## Number Format Normalization Preview",
      },
    };
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes("FROM workbook_workflow_run AS run") &&
        values?.[1] === "thr-1" &&
        values?.[2] === "alex@example.com"
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
    ]);

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: "doc-1",
      actorUserId: "alex@example.com",
      threadId: "thr-1",
    });

    expect(runs).toEqual([
      expect.objectContaining({
        runId: "workflow-import-1",
        workflowTemplate: "normalizeCurrentSheetNumberFormats",
      }),
    ]);
  });

  it("hydrates formatting workflow templates from durable rows after reload", async () => {
    const run = {
      ...createWorkflowRun(),
      runId: "workflow-formatting-1",
      workflowTemplate: "highlightCurrentSheetOutliers" as const,
      title: "Highlight Current Sheet Outliers",
      summary: "Staged outlier highlights for 2 cells across 1 numeric column on Revenue.",
      artifact: {
        kind: "markdown" as const,
        title: "Current Sheet Outlier Highlights",
        text: "## Highlighted Numeric Outliers",
      },
    };
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes("FROM workbook_workflow_run AS run") &&
        values?.[1] === "thr-1" &&
        values?.[2] === "alex@example.com"
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
    ]);

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: "doc-1",
      actorUserId: "alex@example.com",
      threadId: "thr-1",
    });

    expect(runs).toEqual([
      expect.objectContaining({
        runId: "workflow-formatting-1",
        workflowTemplate: "highlightCurrentSheetOutliers",
      }),
    ]);
  });

  it("loads cancelled workflow runs with cancelled steps", async () => {
    const run = {
      ...createWorkflowRun(),
      summary: "Cancelled workflow: Summarize Workbook",
      status: "cancelled" as const,
      updatedAtUnixMs: 130,
      completedAtUnixMs: 130,
      errorMessage: "Cancelled by alex@example.com.",
      steps: [
        {
          stepId: "inspect-workbook",
          label: "Inspect workbook structure",
          status: "cancelled" as const,
          summary: "Workflow cancelled before this step completed.",
          updatedAtUnixMs: 130,
        },
      ],
      artifact: null,
    };
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes("FROM workbook_workflow_run AS run") &&
        values?.[1] === "thr-1" &&
        values?.[2] === "alex@example.com"
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
          ? []
          : null,
    ]);

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: "doc-1",
      actorUserId: "alex@example.com",
      threadId: "thr-1",
    });

    expect(runs).toEqual([
      expect.objectContaining({
        runId: "workflow-1",
        status: "cancelled",
        errorMessage: "Cancelled by alex@example.com.",
        steps: [
          expect.objectContaining({
            stepId: "inspect-workbook",
            status: "cancelled",
          }),
        ],
        artifact: null,
      }),
    ]);
  });
});
