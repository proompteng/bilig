import { useEffect, useRef } from "react";
import type {
  WorkbookAgentCommandBundle,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewSummary,
} from "@bilig/agent-api";
import type {
  WorkbookAgentSessionSnapshot,
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
} from "@bilig/contracts";
import { cn } from "./cn.js";

function contextLabel(context: WorkbookAgentUiContext | null): string {
  if (!context) {
    return "No selection context";
  }
  return `${context.selection.sheetName}!${context.selection.address}`;
}

function ToolStatusPill(props: { readonly status: WorkbookAgentTimelineEntry["toolStatus"] }) {
  const label =
    props.status === "completed" ? "Done" : props.status === "failed" ? "Failed" : "Running";
  return (
    <span
      className={cn(
        "inline-flex h-5 items-center rounded-full px-2 text-[10px] font-semibold uppercase tracking-[0.04em]",
        props.status === "completed"
          ? "bg-[#dcfce7] text-[#166534]"
          : props.status === "failed"
            ? "bg-[#fee2e2] text-[#991b1b]"
            : "bg-[#e0f2fe] text-[#075985]",
      )}
    >
      {label}
    </span>
  );
}

function WorkbookAgentEntryRow(props: { readonly entry: WorkbookAgentTimelineEntry }) {
  const { entry } = props;
  if (entry.kind === "user") {
    return (
      <div className="flex justify-end">
        <div className="max-w-[90%] rounded-[var(--wb-radius-control)] bg-[var(--wb-accent-soft)] px-3 py-2 text-[13px] leading-5 text-[var(--wb-text)]">
          {entry.text}
        </div>
      </div>
    );
  }

  if (entry.kind === "assistant") {
    return (
      <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2 text-[13px] leading-5 text-[var(--wb-text)]">
        {entry.text?.trim().length ? entry.text : "Thinking..."}
      </div>
    );
  }

  if (entry.kind === "plan") {
    return (
      <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-2">
        <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
          Plan
        </div>
        <div className="mt-1 whitespace-pre-wrap text-[12px] leading-5 text-[var(--wb-text-muted)]">
          {entry.text?.trim().length ? entry.text : "Planning..."}
        </div>
      </div>
    );
  }

  if (entry.kind === "tool") {
    return (
      <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2">
        <div className="flex items-center justify-between gap-3">
          <div className="min-w-0">
            <div className="text-[12px] font-semibold text-[var(--wb-text)]">{entry.toolName}</div>
            <div className="text-[11px] text-[var(--wb-text-subtle)]">Workbook tool call</div>
          </div>
          <ToolStatusPill status={entry.toolStatus} />
        </div>
        {entry.argumentsText ? (
          <pre className="mt-2 overflow-x-auto rounded-[var(--wb-radius-control)] bg-[var(--wb-app-bg)] px-2 py-2 text-[11px] leading-5 text-[var(--wb-text-muted)]">
            {entry.argumentsText}
          </pre>
        ) : null}
        {entry.outputText ? (
          <pre className="mt-2 overflow-x-auto rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2 text-[11px] leading-5 text-[var(--wb-text-muted)]">
            {entry.outputText}
          </pre>
        ) : null}
      </div>
    );
  }

  return (
    <div className="rounded-[var(--wb-radius-control)] border border-[#f1b5b5] bg-[#fff7f7] px-3 py-2 text-[12px] leading-5 text-[#991b1b]">
      {entry.text}
    </div>
  );
}

function PreviewRangeList(props: {
  readonly ranges: readonly {
    sheetName: string;
    startAddress: string;
    endAddress: string;
    role: "target" | "source";
  }[];
}) {
  if (props.ranges.length === 0) {
    return null;
  }
  return (
    <div className="mt-2 flex flex-wrap gap-1.5">
      {props.ranges.map((range) => (
        <span
          key={`${range.role}:${range.sheetName}:${range.startAddress}:${range.endAddress}`}
          className={cn(
            "inline-flex items-center rounded-full px-2 py-1 text-[10px] font-medium",
            range.role === "target" ? "bg-[#e0f2fe] text-[#0c4a6e]" : "bg-[#f1f5f9] text-[#475569]",
          )}
        >
          {range.role === "target" ? "Target" : "Source"} {range.sheetName}!{range.startAddress}
          {range.startAddress === range.endAddress ? "" : `:${range.endAddress}`}
        </span>
      ))}
    </div>
  );
}

function PendingBundleCard(props: {
  readonly bundle: WorkbookAgentCommandBundle;
  readonly preview: WorkbookAgentPreviewSummary | null;
  readonly isApplyingBundle: boolean;
  readonly onApply: () => void;
  readonly onDismiss: () => void;
}) {
  const canApply = props.preview !== null && !props.isApplyingBundle;
  const applyLabel =
    props.bundle.approvalMode === "explicit"
      ? props.isApplyingBundle
        ? "Approving..."
        : "Approve and Apply"
      : props.bundle.approvalMode === "auto"
        ? props.isApplyingBundle
          ? "Auto-Applying..."
          : "Apply Now"
        : props.isApplyingBundle
          ? "Applying..."
          : "Apply Preview";
  return (
    <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-surface)] px-3 py-3 shadow-[var(--wb-shadow-sm)]">
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-accent)]">
            Pending Preview
          </div>
          <div className="mt-1 text-[13px] font-semibold text-[var(--wb-text)]">
            {props.bundle.summary}
          </div>
        </div>
        <span className="rounded-full bg-[var(--wb-accent-soft)] px-2 py-1 text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-accent)]">
          {props.bundle.riskClass}
        </span>
      </div>
      <div className="mt-2 text-[12px] leading-5 text-[var(--wb-text-subtle)]">
        Scope: {props.bundle.scope}. Base revision: r{String(props.bundle.baseRevision)}.
        {props.bundle.estimatedAffectedCells === null
          ? ""
          : ` ${String(props.bundle.estimatedAffectedCells)} affected cell${
              props.bundle.estimatedAffectedCells === 1 ? "" : "s"
            }.`}
      </div>
      <div className="mt-2 text-[11px] font-medium text-[var(--wb-text-subtle)]">
        Approval:{" "}
        {props.bundle.approvalMode === "auto"
          ? "auto-apply after local preview"
          : props.bundle.approvalMode === "explicit"
            ? "explicit approval required"
            : "preview required before apply"}
      </div>
      <PreviewRangeList ranges={props.preview?.ranges ?? props.bundle.affectedRanges} />
      {props.preview?.structuralChanges?.length ? (
        <div className="mt-2 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2 text-[11px] leading-5 text-[var(--wb-text-muted)]">
          {props.preview.structuralChanges.join(" · ")}
        </div>
      ) : null}
      {props.preview?.cellDiffs?.length ? (
        <div className="mt-2 overflow-hidden rounded-[var(--wb-radius-control)] border border-[var(--wb-border)]">
          <div className="border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2 py-1 text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
            Local Preview Diff
          </div>
          <div className="max-h-44 overflow-y-auto">
            {props.preview.cellDiffs.map((diff) => (
              <div
                key={`${diff.sheetName}:${diff.address}`}
                className="grid grid-cols-[minmax(0,1fr)_minmax(0,1fr)] gap-x-2 border-t border-[var(--wb-border)] px-2 py-2 text-[11px] leading-5 first:border-t-0"
              >
                <div className="col-span-2 font-medium text-[var(--wb-text)]">
                  {diff.sheetName}!{diff.address}
                </div>
                <div className="text-[var(--wb-text-subtle)]">
                  Before: {(diff.beforeFormula ?? String(diff.beforeInput ?? "")) || "(empty)"}
                </div>
                <div className="text-[var(--wb-text)]">
                  After: {(diff.afterFormula ?? String(diff.afterInput ?? "")) || "(empty)"}
                </div>
              </div>
            ))}
          </div>
        </div>
      ) : null}
      <div className="mt-3 flex items-center justify-end gap-2">
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onDismiss}
        >
          Dismiss
        </button>
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 text-[12px] font-semibold text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)] transition-colors hover:brightness-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
          disabled={!canApply}
          type="button"
          onClick={props.onApply}
        >
          {applyLabel}
        </button>
      </div>
    </div>
  );
}

function ExecutionRecordRow(props: {
  readonly record: WorkbookAgentExecutionRecord;
  readonly onReplay: () => void;
}) {
  return (
    <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="text-[12px] font-semibold text-[var(--wb-text)]">
            {props.record.summary}
          </div>
          <div className="text-[11px] text-[var(--wb-text-subtle)]">
            {props.record.appliedBy === "auto" ? "Auto-applied" : "Applied"} at r
            {String(props.record.appliedRevision)} · {props.record.riskClass} risk ·{" "}
            {props.record.approvalMode}
          </div>
        </div>
        <span className="rounded-full bg-[var(--wb-surface-subtle)] px-2 py-1 text-[10px] font-medium uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
          {props.record.scope}
        </span>
      </div>
      {(props.record.planText ?? props.record.goalText).trim().length > 0 ? (
        <div className="mt-2 text-[11px] leading-5 text-[var(--wb-text-subtle)]">
          {props.record.planText ?? props.record.goalText}
        </div>
      ) : null}
      <PreviewRangeList ranges={props.record.preview?.ranges ?? []} />
      <div className="mt-3 flex items-center justify-end">
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onReplay}
        >
          Replay as Preview
        </button>
      </div>
    </div>
  );
}

export function WorkbookAgentPanel(props: {
  readonly currentContext: WorkbookAgentUiContext | null;
  readonly snapshot: WorkbookAgentSessionSnapshot | null;
  readonly pendingBundle: WorkbookAgentCommandBundle | null;
  readonly preview: WorkbookAgentPreviewSummary | null;
  readonly executionRecords: readonly WorkbookAgentExecutionRecord[];
  readonly draft: string;
  readonly error: string | null;
  readonly isLoading: boolean;
  readonly isApplyingBundle: boolean;
  readonly isOpen: boolean;
  readonly onApplyPendingBundle: () => void;
  readonly onClose: () => void;
  readonly onDraftChange: (value: string) => void;
  readonly onDismissPendingBundle: () => void;
  readonly onInterrupt: () => void;
  readonly onReplayExecutionRecord: (recordId: string) => void;
  readonly onSubmit: () => void;
}) {
  const scrollRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const node = scrollRef.current;
    if (!node) {
      return;
    }
    node.scrollTop = node.scrollHeight;
  }, [props.snapshot?.entries.length, props.snapshot?.status]);

  if (!props.isOpen) {
    return null;
  }

  const isRunning = props.snapshot?.status === "inProgress";

  return (
    <div
      className="flex h-full min-h-0 w-full flex-col"
      data-testid="workbook-agent-panel"
      id="workbook-agent-panel"
    >
      <div className="flex items-center justify-between gap-3 border-b border-[var(--wb-border)] bg-[var(--wb-surface)] px-4 py-3">
        <div className="min-w-0">
          <h2 className="text-[13px] font-semibold text-[var(--wb-text)]">Assistant</h2>
          <p className="text-[11px] text-[var(--wb-text-subtle)]">
            {contextLabel(props.snapshot?.context ?? props.currentContext)}
          </p>
        </div>
        <button
          aria-label="Close assistant"
          className="inline-flex h-8 w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onClose}
        >
          <span aria-hidden="true">×</span>
        </button>
      </div>
      <div ref={scrollRef} className="min-h-0 flex-1 overflow-y-auto px-3 py-3">
        {props.pendingBundle ? (
          <div className="mb-3">
            <PendingBundleCard
              bundle={props.pendingBundle}
              preview={props.preview}
              isApplyingBundle={props.isApplyingBundle}
              onApply={props.onApplyPendingBundle}
              onDismiss={props.onDismissPendingBundle}
            />
          </div>
        ) : null}
        {props.isLoading ? (
          <div className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface)] px-4 py-5 text-sm text-[var(--wb-text-subtle)]">
            Starting workbook assistant...
          </div>
        ) : props.snapshot && props.snapshot.entries.length > 0 ? (
          <div className="flex flex-col gap-2">
            {props.snapshot.entries.map((entry) => (
              <WorkbookAgentEntryRow key={entry.id} entry={entry} />
            ))}
            {props.executionRecords.length > 0 ? (
              <div className="pt-2">
                <div className="mb-2 text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
                  Applied Plans
                </div>
                <div className="flex flex-col gap-2">
                  {props.executionRecords.slice(0, 5).map((record) => (
                    <ExecutionRecordRow
                      key={record.id}
                      record={record}
                      onReplay={() => {
                        props.onReplayExecutionRecord(record.id);
                      }}
                    />
                  ))}
                </div>
              </div>
            ) : null}
          </div>
        ) : (
          <div className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface)] px-4 py-5 text-sm text-[var(--wb-text-subtle)]">
            Ask the assistant to inspect, edit, or restructure this workbook.
          </div>
        )}
      </div>
      {props.error ? (
        <div className="border-t border-[#f1b5b5] bg-[#fff7f7] px-4 py-2 text-[12px] text-[#991b1b]">
          {props.error}
        </div>
      ) : null}
      <form
        className="border-t border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3"
        onSubmit={(event) => {
          event.preventDefault();
          props.onSubmit();
        }}
      >
        <label className="sr-only" htmlFor="workbook-agent-input">
          Ask the workbook assistant
        </label>
        <textarea
          id="workbook-agent-input"
          className="min-h-24 w-full resize-none rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-2 text-[13px] leading-5 text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
          data-testid="workbook-agent-input"
          placeholder="Ask the assistant to update this workbook..."
          value={props.draft}
          onChange={(event) => {
            props.onDraftChange(event.target.value);
          }}
        />
        <div className="mt-3 flex items-center justify-between gap-2">
          <div className="text-[11px] text-[var(--wb-text-subtle)]">
            Uses local workbook tools through the monolith agent runtime.
          </div>
          <div className="flex items-center gap-2">
            {isRunning ? (
              <button
                className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[#f1b5b5] bg-[#fff7f7] px-3 text-[12px] font-medium text-[#991b1b] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[#e58e8e] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#f1b5b5] focus-visible:ring-offset-1"
                data-testid="workbook-agent-interrupt"
                type="button"
                onClick={props.onInterrupt}
              >
                Stop
              </button>
            ) : null}
            <button
              className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 text-[12px] font-semibold text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)] transition-colors hover:brightness-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
              data-testid="workbook-agent-send"
              disabled={props.draft.trim().length === 0 || props.isLoading}
              type="submit"
            >
              Send
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}
