import { cn } from "./cn.js";
import {
  formatWorkbookChangeTimestamp,
  type WorkbookChangeEntry,
} from "./workbook-changes-model.js";

const CHANGE_EVENT_TONE_CLASS_NAMES: Record<string, string> = {
  setCellValue: "bg-[var(--color-mauve-100)] text-[var(--color-mauve-800)]",
  setCellFormula: "bg-[var(--color-mauve-100)] text-[var(--color-mauve-800)]",
  clearCell: "bg-[#fee2e2] text-[#991b1b]",
  clearRange: "bg-[#fee2e2] text-[#991b1b]",
  fillRange: "bg-[#dcfce7] text-[#166534]",
  copyRange: "bg-[#dcfce7] text-[#166534]",
  moveRange: "bg-[#dcfce7] text-[#166534]",
  updateColumnWidth: "bg-[#ede9fe] text-[#5b21b6]",
  setRangeStyle: "bg-[#fef3c7] text-[#92400e]",
  clearRangeStyle: "bg-[#fef3c7] text-[#92400e]",
  setRangeNumberFormat: "bg-[#fef3c7] text-[#92400e]",
  clearRangeNumberFormat: "bg-[#fef3c7] text-[#92400e]",
  renderCommit: "bg-[#e2e8f0] text-[#334155]",
  restoreVersion: "bg-[var(--color-mauve-100)] text-[var(--color-mauve-800)]",
  revertChange: "bg-[#fee2e2] text-[#991b1b]",
  applyBatch: "bg-[#e2e8f0] text-[#334155]",
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
      <div className="mt-2 text-[11px] text-[var(--wb-text-subtle)]">
        {change.targetLabel ? (
          change.isJumpable ? (
            <>{change.targetLabel}</>
          ) : (
            <>{change.targetLabel}</>
          )
        ) : (
          <></>
        )}
      </div>
    </>
  );

  return (
    <div
      className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3 shadow-[var(--wb-shadow-sm)]"
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
            className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[#f0c2c2] bg-[#fff7f7] px-3 text-[12px] font-medium text-[#991b1b] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[#e58e8e] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#f1b5b5] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
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
          <div className="flex flex-col gap-2">
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
