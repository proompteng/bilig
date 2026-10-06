import type { WorkbookSaveState } from './workbook-save-state.js'

const SAVE_LABELS = {
  loading: 'Loading…',
  'read-only': 'Read only',
  error: 'Save failed',
  saving: 'Saving…',
  local: 'Saved on this device',
  offline: 'Offline',
  saved: 'Saved',
} satisfies Record<WorkbookSaveState, string>

export function WorkbookSaveStatus({ state }: { state: WorkbookSaveState }) {
  const className =
    state === 'error'
      ? 'inline-flex h-8 items-center text-[12px] font-medium text-[var(--wb-danger-text)]'
      : state === 'offline' || state === 'read-only'
        ? 'inline-flex h-8 items-center text-[12px] font-medium text-[var(--wb-warning)]'
        : 'sr-only'

  return (
    <span className={className} data-testid="workbook-save-status" role="status">
      {SAVE_LABELS[state]}
    </span>
  )
}
