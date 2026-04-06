import { useEffect, useRef } from "react";
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

export function WorkbookAgentPanel(props: {
  readonly snapshot: WorkbookAgentSessionSnapshot | null;
  readonly draft: string;
  readonly error: string | null;
  readonly isLoading: boolean;
  readonly isOpen: boolean;
  readonly onClose: () => void;
  readonly onDraftChange: (value: string) => void;
  readonly onInterrupt: () => void;
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
            {contextLabel(props.snapshot?.context ?? null)}
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
        {props.isLoading ? (
          <div className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface)] px-4 py-5 text-sm text-[var(--wb-text-subtle)]">
            Starting workbook assistant...
          </div>
        ) : props.snapshot && props.snapshot.entries.length > 0 ? (
          <div className="flex flex-col gap-2">
            {props.snapshot.entries.map((entry) => (
              <WorkbookAgentEntryRow key={entry.id} entry={entry} />
            ))}
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
