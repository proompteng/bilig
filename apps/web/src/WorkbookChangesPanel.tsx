import { cn } from "./cn.js";
import {
  formatWorkbookChangeTimestamp,
  type WorkbookChangeEntry,
} from "./workbook-changes-model.js";

const CHANGE_EVENT_TONE_CLASS_NAMES: Record<string, string> = {
  setCellValue:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  setCellFormula:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  clearCell: "border border-[#efc7c7] bg-[#fff6f6] text-[#8f2d2d]",
  clearRange: "border border-[#efc7c7] bg-[#fff6f6] text-[#8f2d2d]",
  fillRange:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  copyRange:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  moveRange:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  updateColumnWidth:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  setRangeStyle:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  clearRangeStyle:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  setRangeNumberFormat:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  clearRangeNumberFormat:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  renderCommit:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  restoreVersion:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
  revertChange: "border border-[#efc7c7] bg-[#fff6f6] text-[#8f2d2d]",
  applyBatch:
    "border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
};

function formatEventLabel(eventKind: string): string {
  switch (eventKind) {
    case "setCellValue":
    case "setCellFormula":
    case "clearCell":
    case "clearRange":
      return "Value";
    case "fillRange":
    case "copyRange":
    case "moveRange":
      return "Range";
    case "updateColumnWidth":
      return "Layout";
    case "setRangeStyle":
    case "clearRangeStyle":
    case "setRangeNumberFormat":
    case "clearRangeNumberFormat":
      return "Format";
    case "renderCommit":
      return "Batch";
    case "restoreVersion":
      return "Version";
    case "revertChange":
      return "Revert";
    case "applyBatch":
      return "Sync";
    default:
      return "Change";
  }
}

function WorkbookChangeRow(props: {
  readonly change: WorkbookChangeEntry;
  readonly isPending: boolean;
  readonly onJump: (sheetName: string, address: string) => void;
  readonly onRevert: (change: WorkbookChangeEntry) => void;
}) {
  const { change, isPending } = props;
  const content = (
    <>
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2">
            <span
              className={cn(
                "inline-flex h-5 items-center rounded-full px-2 text-[10px] font-semibold uppercase tracking-[0.04em]",
                CHANGE_EVENT_TONE_CLASS_NAMES[change.eventKind] ?? "bg-[#e2e8f0] text-[#334155]",
              )}
            >
              {formatEventLabel(change.eventKind)}
            </span>
            <span className="text-[11px] text-[var(--wb-text-subtle)]">r{change.revision}</span>
          </div>
          <div className="mt-2 text-[13px] font-medium leading-5 text-[var(--wb-text)]">
            {change.summary}
          </div>
          <div className="mt-1 flex flex-wrap items-center gap-x-2 gap-y-1 text-[11px] text-[var(--wb-text-subtle)]">
            <span>{change.actorLabel}</span>
            <span aria-hidden="true">•</span>
            <span>{formatWorkbookChangeTimestamp(change.createdAt)}</span>
          </div>
        </div>
      </div>
      {change.targetLabel ? (
        <div className="mt-2 text-[11px] text-[var(--wb-text-subtle)]">{change.targetLabel}</div>
      ) : null}
    </>
  );

  return (
    <div
      className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3"
      data-testid="workbook-change-row"
    >
      {change.isJumpable ? (
        <button
          className="w-full text-left transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={() => {
            if (change.sheetName && change.address) {
              props.onJump(change.sheetName, change.address);
            }
          }}
        >
          {content}
        </button>
      ) : (
        <div className="opacity-85">{content}</div>
      )}
      <div className="mt-3 flex items-center justify-between gap-2">
        {change.revertedByRevision !== null ? (
          <span
            className="text-[11px] font-medium text-[var(--wb-text-subtle)]"
            data-testid="workbook-change-reverted"
          >
            Reverted in r{change.revertedByRevision}
          </span>
        ) : change.revertsRevision !== null ? (
          <span className="text-[11px] text-[var(--wb-text-subtle)]">
            Reverted r{change.revertsRevision}
          </span>
        ) : (
          <span />
        )}
        {change.canRevert ? (
          <button
            className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[#efc7c7] bg-[#fffafa] px-3 text-[12px] font-medium text-[#8f2d2d] transition-colors hover:bg-[#fff6f6] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#f1b5b5] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
            data-testid="workbook-change-revert"
            disabled={isPending}
            type="button"
            onClick={() => {
              props.onRevert(change);
            }}
          >
            {isPending ? "Reverting..." : "Revert"}
          </button>
        ) : null}
      </div>
    </div>
  );
}

export function WorkbookChangesPanel(props: {
  readonly changes: readonly WorkbookChangeEntry[];
  readonly onJump: (sheetName: string, address: string) => void;
  readonly onRevert: (change: WorkbookChangeEntry) => void;
  readonly pendingRevisions: readonly number[];
}) {
  return (
    <div
      aria-label="Workbook changes"
      className="flex h-full min-h-0 flex-col"
      data-testid="workbook-changes-panel"
      id="workbook-changes-panel"
    >
      <div className="min-h-0 flex-1 overflow-y-auto px-3 py-3">
        {props.changes.length === 0 ? (
          <div />
        ) : (
          <div className="flex flex-col gap-1.5">
            {props.changes.map((change) => (
              <WorkbookChangeRow
                key={`${change.revision}:${change.summary}`}
                change={change}
                isPending={props.pendingRevisions.includes(change.revision)}
                onJump={props.onJump}
                onRevert={props.onRevert}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
