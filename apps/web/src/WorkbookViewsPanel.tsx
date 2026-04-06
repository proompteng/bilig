import { cn } from "./cn.js";
import {
  formatWorkbookSheetViewTimestamp,
  type WorkbookSheetViewEntry,
  type WorkbookSheetViewVisibility,
} from "./workbook-views-model.js";

function WorkbookViewSection(props: {
  readonly title: string;
  readonly views: readonly WorkbookSheetViewEntry[];
  readonly onApply: (view: WorkbookSheetViewEntry) => void;
  readonly onUpdate: (view: WorkbookSheetViewEntry) => void;
  readonly onDelete: (view: WorkbookSheetViewEntry) => void;
}) {
  if (props.views.length === 0) {
    return null;
  }

  return (
    <section className="flex flex-col gap-2">
      <h3 className="px-1 text-[11px] font-semibold uppercase tracking-[0.08em] text-[var(--wb-text-subtle)]">
        {props.title}
      </h3>
      <div className="flex flex-col gap-2">
        {props.views.map((view) => (
          <div
            key={view.id}
            className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3 shadow-[var(--wb-shadow-sm)]"
            data-testid="workbook-view-row"
          >
            <div className="flex items-start justify-between gap-3">
              <div className="min-w-0">
                <div className="flex items-center gap-2">
                  <span
                    className={cn(
                      "inline-flex rounded-full px-2 py-0.5 text-[10px] font-semibold uppercase tracking-[0.08em]",
                      view.visibility === "private"
                        ? "bg-[var(--wb-accent-soft)] text-[var(--wb-accent)]"
                        : "bg-[var(--wb-surface-subtle)] text-[var(--wb-text-muted)]",
                    )}
                  >
                    {view.visibility}
                  </span>
                  <span className="truncate text-[13px] font-semibold text-[var(--wb-text)]">
                    {view.name}
                  </span>
                </div>
                <div className="mt-1 text-[11px] text-[var(--wb-text-subtle)]">
                  {view.visibility === "shared"
                    ? `Shared by ${view.ownerLabel}`
                    : "Visible only to you"}
                </div>
              </div>
            </div>
            <div className="mt-2 flex flex-wrap items-center gap-x-2 gap-y-1 text-[11px] text-[var(--wb-text-subtle)]">
              <span>{view.targetLabel}</span>
              <span>•</span>
              <span>{formatWorkbookSheetViewTimestamp(view.updatedAt)}</span>
            </div>
            <div className="mt-3 flex items-center gap-2">
              <button
                className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[var(--wb-accent-ring)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
                disabled={!view.isApplyable}
                type="button"
                onClick={() => {
                  props.onApply(view);
                }}
              >
                Apply
              </button>
              <button
                className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
                disabled={!view.canManage}
                type="button"
                onClick={() => {
                  props.onUpdate(view);
                }}
              >
                Update
              </button>
              <button
                className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[#f0c2c2] bg-[#fff7f7] px-3 text-[12px] font-medium text-[#991b1b] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[#e58e8e] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#f1b5b5] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
                disabled={!view.canManage}
                type="button"
                onClick={() => {
                  props.onDelete(view);
                }}
              >
                Delete
              </button>
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

export function WorkbookViewsPanel(props: {
  readonly isOpen: boolean;
  readonly views: readonly WorkbookSheetViewEntry[];
  readonly draftName: string;
  readonly draftVisibility: WorkbookSheetViewVisibility;
  readonly onDraftNameChange: (value: string) => void;
  readonly onDraftVisibilityChange: (visibility: WorkbookSheetViewVisibility) => void;
  readonly onSave: () => void;
  readonly onClose: () => void;
  readonly onApply: (view: WorkbookSheetViewEntry) => void;
  readonly onUpdate: (view: WorkbookSheetViewEntry) => void;
  readonly onDelete: (view: WorkbookSheetViewEntry) => void;
}) {
  if (!props.isOpen) {
    return null;
  }

  const privateViews = props.views.filter((view) => view.visibility === "private");
  const sharedViews = props.views.filter((view) => view.visibility === "shared");

  return (
    <aside
      aria-label="Workbook views"
      className="absolute inset-y-3 right-[24rem] z-20 flex w-[20rem] flex-col overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[0_20px_48px_rgba(15,23,42,0.16)]"
      data-testid="workbook-views-panel"
      id="workbook-views-panel"
    >
      <div className="flex items-center justify-between gap-3 border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-4 py-3">
        <div className="min-w-0">
          <h2 className="text-[13px] font-semibold text-[var(--wb-text)]">Views</h2>
          <p className="text-[11px] text-[var(--wb-text-subtle)]">
            Save the current sheet, cell, and visible frame as a private or shared view.
          </p>
        </div>
        <button
          aria-label="Close workbook views"
          className="inline-flex h-8 w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onClose}
        >
          ×
        </button>
      </div>
      <div className="border-b border-[var(--wb-border)] px-4 py-3">
        <label className="flex flex-col gap-2">
          <span className="text-[11px] font-medium text-[var(--wb-text-subtle)]">Name</span>
          <input
            className="h-9 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[13px] text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] outline-none transition-colors focus:border-[var(--wb-accent-ring)] focus:ring-2 focus:ring-[var(--wb-accent-ring)] focus:ring-offset-1"
            data-testid="workbook-view-name"
            placeholder="Current focus"
            type="text"
            value={props.draftName}
            onChange={(event) => {
              props.onDraftNameChange(event.currentTarget.value);
            }}
          />
        </label>
        <div className="mt-3 flex items-center gap-2">
          {(["private", "shared"] as const).map((visibility) => (
            <button
              key={visibility}
              className={cn(
                "inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border px-3 text-[12px] font-medium shadow-[var(--wb-shadow-sm)] transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1",
                props.draftVisibility === visibility
                  ? "border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] text-[var(--wb-accent)]"
                  : "border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] hover:text-[var(--wb-text)]",
              )}
              type="button"
              onClick={() => {
                props.onDraftVisibilityChange(visibility);
              }}
            >
              {visibility === "private" ? "Private" : "Shared"}
            </button>
          ))}
          <button
            className="ml-auto inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 text-[12px] font-semibold text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)] transition-colors hover:bg-[var(--wb-accent-soft)]/80 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
            data-testid="workbook-view-save"
            type="button"
            onClick={props.onSave}
          >
            Save current
          </button>
        </div>
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto px-3 py-3">
        {props.views.length === 0 ? (
          <div className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-4 py-5 text-sm text-[var(--wb-text-subtle)]">
            Save a view to reopen the same workbook frame later without disturbing anyone else.
          </div>
        ) : (
          <div className="flex flex-col gap-4">
            <WorkbookViewSection
              title="Private"
              views={privateViews}
              onApply={props.onApply}
              onDelete={props.onDelete}
              onUpdate={props.onUpdate}
            />
            <WorkbookViewSection
              title="Shared"
              views={sharedViews}
              onApply={props.onApply}
              onDelete={props.onDelete}
              onUpdate={props.onUpdate}
            />
          </div>
        )}
      </div>
    </aside>
  );
}
