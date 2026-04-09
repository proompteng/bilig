import { cva } from "class-variance-authority";
import { cn } from "./cn.js";
import {
  formatWorkbookChangeTimestamp,
  type WorkbookChangeEntry,
} from "./workbook-changes-model.js";
import { workbookButtonClass, workbookPillClass } from "./workbook-shell-chrome.js";

const changeEventToneClass = cva("", {
  variants: {
    tone: {
      neutral: workbookPillClass({ tone: "neutral", weight: "strong" }),
      danger: workbookPillClass({ tone: "danger", weight: "strong" }),
    },
  },
});

const CHANGE_EVENT_TONES: Record<string, "neutral" | "danger"> = {
  setCellValue: "neutral",
  setCellFormula: "neutral",
  clearCell: "danger",
  clearRange: "danger",
  fillRange: "neutral",
  copyRange: "neutral",
  moveRange: "neutral",
  updateRowMetadata: "neutral",
  updateColumnMetadata: "neutral",
  updateColumnWidth: "neutral",
  setFreezePane: "neutral",
  setRangeStyle: "neutral",
  clearRangeStyle: "neutral",
  setRangeNumberFormat: "neutral",
  clearRangeNumberFormat: "neutral",
  renderCommit: "neutral",
  restoreVersion: "neutral",
  revertChange: "danger",
  applyBatch: "neutral",
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
    case "updateRowMetadata":
    case "updateColumnMetadata":
    case "updateColumnWidth":
    case "setFreezePane":
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
                changeEventToneClass({
                  tone: CHANGE_EVENT_TONES[change.eventKind] ?? "neutral",
                }),
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
            className={workbookButtonClass({ tone: "danger" })}
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
