import { cva } from 'class-variance-authority'
import { useEffect, useMemo, useRef, type CSSProperties } from 'react'
import { Search, X } from 'lucide-react'
import { Dialog } from '@base-ui/react/dialog'
import {
  getWorkbookShortcutLabel,
  getWorkbookShortcutParts,
  groupWorkbookShortcutEntries,
  searchWorkbookShortcutEntries,
} from './shortcut-registry.js'

const shortcutChordClass = cva('flex items-center gap-1.5')

const shortcutDialogThemeStyle: CSSProperties & Record<`--${string}`, string> = {
  '--wb-app-bg': 'var(--color-mauve-50)',
  '--wb-surface': 'white',
  '--wb-surface-subtle': 'var(--color-mauve-50)',
  '--wb-surface-muted': 'var(--color-mauve-100)',
  '--wb-border': 'var(--color-mauve-200)',
  '--wb-border-strong': 'var(--color-mauve-300)',
  '--wb-grid-border': 'var(--color-mauve-100)',
  '--wb-text': 'var(--color-mauve-900)',
  '--wb-text-muted': 'var(--color-mauve-700)',
  '--wb-text-subtle': 'var(--color-mauve-600)',
  '--wb-accent': 'var(--color-mauve-900)',
  '--wb-accent-soft': 'var(--color-mauve-100)',
  '--wb-accent-ring': 'var(--color-mauve-400)',
  '--wb-hover': 'var(--color-mauve-100)',
  '--wb-shadow-sm': '0 1px 2px rgba(15, 23, 42, 0.04)',
  '--wb-shadow-md': '0 16px 40px rgba(15, 23, 42, 0.12)',
}

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

  useEffect(() => {
    if (!props.open) {
      return
    }

    const focusSearchInput = () => {
      searchInputRef.current?.focus({ preventScroll: true })
    }

    focusSearchInput()
    const timeoutId = window.setTimeout(focusSearchInput, 0)

    return () => {
      window.clearTimeout(timeoutId)
    }
  }, [props.open])

  return (
    <Dialog.Root open={props.open} onOpenChange={props.onOpenChange}>
      <Dialog.Portal>
        <Dialog.Backdrop className="fixed inset-0 z-[1200] bg-black/35" />
        <Dialog.Popup
          aria-label="Keyboard shortcuts"
          aria-modal="true"
          className="fixed left-1/2 top-1/2 z-[1201] flex w-[min(56rem,calc(100vw-3rem))] max-w-[calc(100vw-3rem)] -translate-x-1/2 -translate-y-1/2 flex-col overflow-hidden rounded-xl border border-[var(--wb-border)] bg-[var(--color-mauve-50)] shadow-[var(--wb-shadow-md)]"
          data-testid="workbook-shortcut-dialog"
          initialFocus={searchInputRef}
          style={shortcutDialogThemeStyle}
        >
          <div className="flex items-start justify-between gap-4 border-b border-[var(--wb-border)] bg-[var(--color-mauve-50)] px-5 py-4">
            <div className="min-w-0">
              <Dialog.Title className="text-[16px] font-semibold text-[var(--wb-text)]">Keyboard shortcuts</Dialog.Title>
              <Dialog.Description className="mt-1 max-w-[48ch] text-[13px] text-[var(--wb-text-subtle)]">
                Search shortcuts and commands already supported by the workbook shell.
              </Dialog.Description>
            </div>
            <Dialog.Close
              aria-label="Close shortcuts"
              className="inline-flex h-9 w-9 items-center justify-center rounded-md border border-transparent text-[var(--wb-text-muted)] transition-colors hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
            >
              <X className="h-4 w-4" />
            </Dialog.Close>
          </div>

          <div className="border-b border-[var(--wb-border)] bg-white px-5 py-4">
            <label className="sr-only" htmlFor="workbook-shortcut-search">
              Search shortcuts
            </label>
            <div className="flex h-10 items-center gap-2 rounded-md border border-[var(--wb-border)] bg-white px-3 shadow-[var(--wb-shadow-sm)]">
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

          <div className="max-h-[34rem] overflow-y-auto bg-[var(--color-mauve-50)] px-5 py-4">
            {groupedEntries.length === 0 ? (
              <div
                className="rounded-md border border-dashed border-[var(--wb-border)] bg-white px-4 py-8 text-center text-[12px] text-[var(--wb-text-subtle)]"
                data-testid="workbook-shortcut-empty"
              >
                No shortcuts match that search.
              </div>
            ) : (
              <div className="flex flex-col gap-4">
                {groupedEntries.map((group) => (
                  <section key={group.category}>
                    <h3 className="mb-2 px-1 text-[11px] font-semibold uppercase tracking-[0.08em] text-[var(--wb-text-subtle)]">
                      {group.category}
                    </h3>
                    <div className="overflow-hidden rounded-lg border border-[var(--color-mauve-200)] bg-white">
                      {group.entries.map((entry) => (
                        <div
                          className="flex items-center justify-between gap-4 border-b border-[var(--color-mauve-200)] bg-transparent px-4 py-3 last:border-b-0"
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
