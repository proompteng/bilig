import {
  formatWorkbookVersionTimestamp,
  type WorkbookVersionEntry,
} from "./workbook-versions-model.js";

function WorkbookVersionRow(props: {
  readonly version: WorkbookVersionEntry;
  readonly onRestore: (version: WorkbookVersionEntry) => void;
  readonly onDelete: (version: WorkbookVersionEntry) => void;
}) {
  const { version } = props;
  return (
    <div
      className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3 shadow-[var(--wb-shadow-sm)]"
      data-testid="workbook-version-row"
    >
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2">
            <span className="inline-flex rounded-full bg-[var(--wb-accent-soft)] px-2 py-0.5 text-[10px] font-semibold uppercase tracking-[0.08em] text-[var(--wb-accent)]">
              Version
            </span>
            <span className="truncate text-[13px] font-semibold text-[var(--wb-text)]">
              {version.name}
            </span>
          </div>
          <div className="mt-1 text-[11px] text-[var(--wb-text-subtle)]">
            Saved by {version.ownerLabel}
          </div>
        </div>
        <span className="text-[11px] text-[var(--wb-text-subtle)]">r{version.revision}</span>
      </div>
      <div className="mt-2 flex flex-wrap items-center gap-x-2 gap-y-1 text-[11px] text-[var(--wb-text-subtle)]">
        {version.targetLabel ? (
          <span>{version.targetLabel}</span>
        ) : (
          <span>Workbook-wide restore</span>
        )}
        <span>•</span>
        <span>{formatWorkbookVersionTimestamp(version.updatedAt)}</span>
      </div>
      <div className="mt-3 flex items-center gap-2">
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[var(--wb-accent-ring)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={() => {
            props.onRestore(version);
          }}
        >
          Restore
        </button>
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[#f0c2c2] bg-[#fff7f7] px-3 text-[12px] font-medium text-[#991b1b] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[#e58e8e] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#f1b5b5] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
          disabled={!version.canDelete}
          type="button"
          onClick={() => {
            props.onDelete(version);
          }}
        >
          Delete
        </button>
      </div>
    </div>
  );
}

export function WorkbookVersionsPanel(props: {
  readonly isOpen: boolean;
  readonly versions: readonly WorkbookVersionEntry[];
  readonly draftName: string;
  readonly onDraftNameChange: (value: string) => void;
  readonly onSave: () => void;
  readonly onClose: () => void;
  readonly onRestore: (version: WorkbookVersionEntry) => void;
  readonly onDelete: (version: WorkbookVersionEntry) => void;
}) {
  if (!props.isOpen) {
    return null;
  }

  return (
    <aside
      aria-label="Workbook versions"
      className="absolute inset-y-3 right-[45rem] z-20 flex w-[20rem] flex-col overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[0_20px_48px_rgba(15,23,42,0.16)]"
      data-testid="workbook-versions-panel"
      id="workbook-versions-panel"
    >
      <div className="flex items-center justify-between gap-3 border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-4 py-3">
        <div className="min-w-0">
          <h2 className="text-[13px] font-semibold text-[var(--wb-text)]">Versions</h2>
          <p className="text-[11px] text-[var(--wb-text-subtle)]">
            Save named checkpoints of the authoritative workbook and restore them later.
          </p>
        </div>
        <button
          aria-label="Close workbook versions"
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
            data-testid="workbook-version-name"
            placeholder="April close"
            type="text"
            value={props.draftName}
            onChange={(event) => {
              props.onDraftNameChange(event.currentTarget.value);
            }}
          />
        </label>
        <div className="mt-3 flex items-center gap-2">
          <button
            className="ml-auto inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 text-[12px] font-semibold text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)] transition-colors hover:bg-[var(--wb-accent-soft)]/80 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
            data-testid="workbook-version-save"
            type="button"
            onClick={props.onSave}
          >
            Save current
          </button>
        </div>
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto px-3 py-3">
        {props.versions.length === 0 ? (
          <div className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-4 py-5 text-sm text-[var(--wb-text-subtle)]">
            Save a named version before risky workbook edits so authoritative restore stays one
            click away.
          </div>
        ) : (
          <div className="flex flex-col gap-2">
            {props.versions.map((version) => (
              <WorkbookVersionRow
                key={version.id}
                version={version}
                onDelete={props.onDelete}
                onRestore={props.onRestore}
              />
            ))}
          </div>
        )}
      </div>
    </aside>
  );
}
