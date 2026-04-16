import { toEditorValue, toResolvedValue, type WorkbookEditorConflict } from './worker-workbook-app-model.js'

function ConflictValueBlock(props: {
  readonly label: string
  readonly primaryValue: string
  readonly secondaryValue?: string | undefined
}) {
  return (
    <div className="min-w-0 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3">
      <div className="text-[11px] font-semibold uppercase tracking-[0.08em] text-[var(--wb-text-subtle)]">{props.label}</div>
      <div className="mt-2 overflow-hidden text-ellipsis whitespace-pre-wrap break-words rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2 font-mono text-[12px] text-[var(--wb-text)]">
        {props.primaryValue || '∅'}
      </div>
      {props.secondaryValue ? <div className="mt-2 text-[11px] text-[var(--wb-text-subtle)]">Resolved: {props.secondaryValue}</div> : null}
    </div>
  )
}

export function WorkbookEditorConflictBanner(props: {
  readonly conflict: WorkbookEditorConflict
  readonly localDraft: string
  readonly onReview: () => void
  readonly onApplyMine: () => void
  readonly onKeepAuthoritative: () => void
  readonly onKeepDraftLocal: () => void
}) {
  const targetLabel = `${props.conflict.sheetName}!${props.conflict.address}`
  const authoritativeEditorValue = toEditorValue(props.conflict.authoritativeSnapshot)
  const authoritativeResolvedValue = toResolvedValue(props.conflict.authoritativeSnapshot)

  if (props.conflict.phase === 'badge') {
    return (
      <div className="border-b border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 py-2" data-testid="editor-conflict-banner">
        <div className="flex flex-wrap items-center gap-3">
          <div className="min-w-0 flex-1 text-sm text-[var(--wb-accent)]">
            Remote update detected in {targetLabel} while you were editing. Your draft is still local.
          </div>
          <button
            className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-surface)] px-3 text-[12px] font-semibold text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)] transition-colors hover:bg-[var(--wb-surface-subtle)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
            data-testid="editor-conflict-review"
            type="button"
            onClick={props.onReview}
          >
            Review conflict
          </button>
        </div>
      </div>
    )
  }

  return (
    <div className="border-b border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 py-3" data-testid="editor-conflict-banner">
      <div className="flex flex-col gap-3">
        <div className="flex flex-wrap items-center gap-3">
          <div className="min-w-0 flex-1">
            <div className="text-[13px] font-semibold text-[var(--wb-accent)]">{targetLabel} changed while you were editing</div>
            <div className="mt-1 text-[12px] text-[var(--wb-text-subtle)]">
              Compare your local draft with the latest authoritative value before deciding what to keep.
            </div>
          </div>
        </div>
        <div className="grid gap-3 md:grid-cols-2">
          <ConflictValueBlock label="My Draft" primaryValue={props.localDraft} />
          <ConflictValueBlock
            label="Authoritative"
            primaryValue={authoritativeEditorValue}
            secondaryValue={authoritativeResolvedValue !== authoritativeEditorValue ? authoritativeResolvedValue : undefined}
          />
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button
            className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-surface)] px-3 text-[12px] font-semibold text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)] transition-colors hover:bg-[var(--wb-surface-subtle)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
            data-testid="editor-conflict-apply-mine"
            type="button"
            onClick={props.onApplyMine}
          >
            Apply mine
          </button>
          <button
            className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[var(--wb-accent-ring)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
            data-testid="editor-conflict-keep-authoritative"
            type="button"
            onClick={props.onKeepAuthoritative}
          >
            Keep authoritative
          </button>
          <button
            className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
            data-testid="editor-conflict-keep-draft-local"
            type="button"
            onClick={props.onKeepDraftLocal}
          >
            Keep draft local
          </button>
        </div>
      </div>
    </div>
  )
}
