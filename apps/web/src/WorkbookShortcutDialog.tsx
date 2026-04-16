import { cva } from 'class-variance-authority'
import { useMemo, useRef } from 'react'
import { Search, X } from 'lucide-react'
import { Dialog } from '@base-ui/react/dialog'
import {
  getWorkbookShortcutLabel,
  getWorkbookShortcutParts,
  groupWorkbookShortcutEntries,
  searchWorkbookShortcutEntries,
} from './shortcut-registry.js'

const shortcutChordClass = cva('flex items-center gap-1.5')

const shortcutKeyClass = cva(
  'inline-flex min-h-7 min-w-7 items-center justify-center rounded-md border border-[var(--color-mauve-300)] bg-white px-2 text-[11px] font-semibold leading-none text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]',
  {
    variants: {
      tokenType: {
        symbol: 'font-sans',
        text: 'font-sans',
      },
    },
    defaultVariants: {
      tokenType: 'text',
    },
  },
)

function ShortcutKeyChord(props: { readonly shortcutId: string }) {
  const parts = getWorkbookShortcutParts(props.shortcutId)
  const label = getWorkbookShortcutLabel(props.shortcutId)
  if (parts.length === 0) {
    return null
  }
  const showPlusSeparators = label.includes('+')
  return (
    <div aria-label={label} className={shortcutChordClass()} data-testid="workbook-shortcut-chord" title={label}>
      {parts.map((part, index) => {
        const isSymbol = part.length === 1 && /[⌘⇧⌥⌃]/.test(part)
        const chordKey = `${props.shortcutId}:${parts.slice(0, index + 1).join('::')}`
        return (
          <div className="flex items-center gap-1.5" key={chordKey}>
            {index > 0 && showPlusSeparators ? <span className="text-[10px] font-medium text-[var(--color-mauve-500)]">+</span> : null}
            <kbd
              className={shortcutKeyClass({
                tokenType: isSymbol ? 'symbol' : 'text',
              })}
            >
              {part}
            </kbd>
          </div>
        )
      })}
    </div>
  )
}

export function WorkbookShortcutDialog(props: {
  open: boolean
  query: string
  onOpenChange(this: void, open: boolean): void
  onQueryChange(this: void, next: string): void
}) {
  const searchInputRef = useRef<HTMLInputElement | null>(null)
  const filteredEntries = useMemo(() => searchWorkbookShortcutEntries(props.query), [props.query])
  const groupedEntries = useMemo(() => groupWorkbookShortcutEntries(filteredEntries), [filteredEntries])

  return (
    <Dialog.Root open={props.open} onOpenChange={props.onOpenChange}>
      <Dialog.Portal>
        <Dialog.Backdrop className="fixed inset-0 z-[1200] bg-black/35" />
        <Dialog.Popup
          aria-label="Keyboard shortcuts"
          className="fixed left-1/2 top-1/2 z-[1201] flex w-[min(42rem,calc(100vw-2rem))] -translate-x-1/2 -translate-y-1/2 flex-col overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-sm)]"
          data-testid="workbook-shortcut-dialog"
          initialFocus={searchInputRef}
        >
          <div className="flex items-start justify-between gap-4 border-b border-[var(--wb-border)] px-4 py-3">
            <div className="min-w-0">
              <Dialog.Title className="text-[14px] font-semibold text-[var(--wb-text)]">Keyboard shortcuts</Dialog.Title>
              <Dialog.Description className="mt-1 text-[12px] text-[var(--wb-text-subtle)]">
                Search shortcuts and commands already supported by the workbook shell.
              </Dialog.Description>
            </div>
            <Dialog.Close
              aria-label="Close shortcuts"
              className="inline-flex h-8 w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-transparent text-[var(--wb-text-muted)] transition-colors hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
            >
              <X className="h-4 w-4" />
            </Dialog.Close>
          </div>

          <div className="border-b border-[var(--wb-border)] px-4 py-3">
            <label className="sr-only" htmlFor="workbook-shortcut-search">
              Search shortcuts
            </label>
            <div className="flex h-9 items-center gap-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3">
              <Search className="h-4 w-4 shrink-0 text-[var(--wb-text-subtle)]" />
              <input
                aria-label="Search shortcuts"
                className="min-w-0 flex-1 border-0 bg-transparent text-[13px] text-[var(--wb-text)] outline-none placeholder:text-[var(--wb-text-subtle)]"
                data-testid="workbook-shortcut-search"
                id="workbook-shortcut-search"
                placeholder="Search actions, keys, or categories"
                ref={searchInputRef}
                type="text"
                value={props.query}
                onChange={(event) => {
                  props.onQueryChange(event.target.value)
                }}
              />
            </div>
          </div>

          <div className="max-h-[28rem] overflow-y-auto px-4 py-3">
            {groupedEntries.length === 0 ? (
              <div
                className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] px-3 py-6 text-center text-[12px] text-[var(--wb-text-subtle)]"
                data-testid="workbook-shortcut-empty"
              >
                No shortcuts match that search.
              </div>
            ) : (
              <div className="flex flex-col gap-4">
                {groupedEntries.map((group) => (
                  <section key={group.category}>
                    <h3 className="mb-2 text-[11px] font-semibold uppercase tracking-[0.08em] text-[var(--wb-text-subtle)]">
                      {group.category}
                    </h3>
                    <div className="grid gap-1.5">
                      {group.entries.map((entry) => (
                        <div
                          className="flex items-center justify-between gap-3 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-muted)] px-3 py-2"
                          data-testid="workbook-shortcut-entry"
                          key={entry.id}
                        >
                          <div className="min-w-0">
                            <div className="text-[13px] font-medium text-[var(--wb-text)]">{entry.label}</div>
                          </div>
                          <ShortcutKeyChord shortcutId={entry.id} />
                        </div>
                      ))}
                    </div>
                  </section>
                ))}
              </div>
            )}
          </div>
        </Dialog.Popup>
      </Dialog.Portal>
    </Dialog.Root>
  )
}
