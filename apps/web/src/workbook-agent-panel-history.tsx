import type { WorkbookAgentExecutionRecord, WorkbookAgentPreviewRange } from "@bilig/agent-api";
import type { WorkbookAgentWorkflowRun } from "@bilig/contracts";
import { cn } from "./cn.js";
import { WorkbookAgentMarkdown } from "./workbook-agent-markdown.js";
import {
  agentPanelBodyTextClass,
  agentPanelLabelTextClass,
  agentPanelMetaTextClass,
} from "./workbook-agent-panel-primitives.js";
import {
  workbookButtonClass,
  workbookInsetClass,
  workbookPillClass,
  workbookSurfaceClass,
} from "./workbook-shell-chrome.js";

const THINKING_LABEL = "Thinking";

export function AssistantProgressRow() {
  return (
    <div
      aria-label="Assistant is thinking"
      className="px-3 pb-3"
      data-testid="workbook-agent-progress-row"
      role="status"
    >
      <div
        className={cn(
          agentPanelBodyTextClass(),
          "inline-flex items-center text-[var(--wb-text-subtle)]",
        )}
      >
        {THINKING_LABEL.split("").map((character, index) => (
          <span
            key={`${character}-${String(index)}`}
            aria-hidden="true"
            className="inline-block animate-[pulse_800ms_ease-in-out_infinite]"
            style={{ animationDelay: `${index * 55}ms` }}
          >
            {character}
          </span>
        ))}
      </div>
    </div>
  );
}

export function PreviewRangeList(props: { readonly ranges: readonly WorkbookAgentPreviewRange[] }) {
  if (props.ranges.length === 0) {
    return null;
  }
  return (
    <div className="mt-2 flex flex-wrap gap-1.5">
      {props.ranges.map((range) => (
        <span
          key={`${range.role}:${range.sheetName}:${range.startAddress}:${range.endAddress}`}
          className={workbookPillClass({
            tone: range.role === "target" ? "accent" : "neutral",
          })}
        >
          {range.sheetName}!{range.startAddress}
          {range.startAddress === range.endAddress ? "" : `:${range.endAddress}`}
        </span>
      ))}
    </div>
  );
}

export function ExecutionRecordRow(props: {
  readonly record: WorkbookAgentExecutionRecord;
  readonly onReplay: () => void;
}) {
  return (
    <div className={cn(workbookSurfaceClass(), "px-3 py-2")}>
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className={cn(agentPanelLabelTextClass(), "font-semibold")}>
            {props.record.summary}
          </div>
          <div className={agentPanelMetaTextClass()}>
            r{String(props.record.appliedRevision)}
          </div>
        </div>
        <span className={workbookPillClass({ tone: "neutral" })}>{props.record.scope}</span>
      </div>
      {(props.record.planText ?? props.record.goalText).trim().length > 0 ? (
        <div className={cn(agentPanelMetaTextClass(), "mt-2")}>
          {props.record.planText ?? props.record.goalText}
        </div>
      ) : null}
      <PreviewRangeList ranges={props.record.preview?.ranges ?? []} />
      <div className="mt-3 flex items-center justify-end">
        <button
          className={workbookButtonClass({ tone: "neutral" })}
          type="button"
          onClick={props.onReplay}
        >
          Replay
        </button>
      </div>
    </div>
  );
}

function workflowStatusTone(
  status: WorkbookAgentWorkflowRun["status"],
): "accent" | "danger" | "neutral" {
  switch (status) {
    case "running":
      return "accent";
    case "failed":
      return "danger";
    case "cancelled":
    case "completed":
      return "neutral";
  }
}

function workflowStatusLabel(status: WorkbookAgentWorkflowRun["status"]): string {
  switch (status) {
    case "running":
      return "Running";
    case "failed":
      return "Failed";
    case "cancelled":
      return "Cancelled";
    case "completed":
      return "Done";
  }
}

function workflowStepTone(
  status: WorkbookAgentWorkflowRun["steps"][number]["status"],
): "accent" | "danger" | "neutral" {
  switch (status) {
    case "running":
      return "accent";
    case "failed":
      return "danger";
    case "pending":
    case "cancelled":
    case "completed":
      return "neutral";
  }
}

function workflowStepLabel(status: WorkbookAgentWorkflowRun["steps"][number]["status"]): string {
  switch (status) {
    case "pending":
      return "Pending";
    case "running":
      return "Running";
    case "completed":
      return "Done";
    case "failed":
      return "Failed";
    case "cancelled":
      return "Cancelled";
  }
}

export function WorkflowRunRow(props: {
  readonly run: WorkbookAgentWorkflowRun;
  readonly isCancelling?: boolean;
  readonly onCancel?: () => void;
}) {
  return (
    <div
      className={cn(workbookSurfaceClass(), "px-3 py-2")}
      data-testid={`workbook-agent-workflow-${props.run.runId}`}
    >
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className={cn(agentPanelLabelTextClass(), "font-semibold")}>{props.run.title}</div>
          <div className={agentPanelMetaTextClass()}>{props.run.summary}</div>
        </div>
        <span
          className={workbookPillClass({
            tone: workflowStatusTone(props.run.status),
            weight: "strong",
          })}
        >
          {workflowStatusLabel(props.run.status)}
        </span>
      </div>
      {props.run.steps.length > 0 ? (
        <div className="mt-2 grid gap-2">
          {props.run.steps.map((step) => (
            <div key={step.stepId} className={cn(workbookInsetClass(), "px-3 py-2")}>
              <div className="flex items-start justify-between gap-2">
                <div className={agentPanelLabelTextClass()}>{step.label}</div>
                <span className={workbookPillClass({ tone: workflowStepTone(step.status) })}>
                  {workflowStepLabel(step.status)}
                </span>
              </div>
              <div className={cn(agentPanelMetaTextClass(), "mt-1")}>
                {step.summary}
              </div>
            </div>
          ))}
        </div>
      ) : null}
      {props.run.artifact ? (
        <div className={cn(workbookInsetClass(), "mt-2 px-3 py-2")}>
          <div className={agentPanelLabelTextClass()}>
            {props.run.artifact.title}
          </div>
          <WorkbookAgentMarkdown
            className="mt-1"
            markdown={props.run.artifact.text}
            title={props.run.artifact.title}
            tone="muted"
          />
        </div>
      ) : null}
      {props.run.errorMessage ? (
        <div
          className={cn(
            agentPanelMetaTextClass(),
            "mt-2",
            props.run.status === "cancelled"
              ? "text-[var(--wb-text-subtle)]"
              : "text-[var(--wb-danger-text)]",
          )}
        >
          {props.run.errorMessage}
        </div>
      ) : null}
      {props.run.status === "running" && props.onCancel ? (
        <div className="mt-3 flex items-center justify-end">
          <button
            className={workbookButtonClass({ tone: "neutral" })}
            data-testid={`workbook-agent-cancel-workflow-${props.run.runId}`}
            disabled={props.isCancelling}
            type="button"
            onClick={props.onCancel}
          >
            {props.isCancelling ? "Cancelling..." : "Cancel"}
          </button>
        </div>
      ) : null}
    </div>
  );
}
