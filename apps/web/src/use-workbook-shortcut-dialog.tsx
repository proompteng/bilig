import { useCallback, useEffect, useMemo, useState } from 'react'
import { Button } from '@base-ui/react/button'
import { CircleHelp } from 'lucide-react'
import { WorkbookShortcutDialog } from './WorkbookShortcutDialog.js'
import { isTextEntryTarget } from './worker-workbook-app-model.js'

export function useWorkbookShortcutDialog() {
  const [isOpen, setIsOpen] = useState(false)
  const [query, setQuery] = useState('')

  const closeShortcutDialog = useCallback(() => {
    setIsOpen(false)
    setQuery('')
  }, [])

  const openShortcutDialog = useCallback(() => {
    setIsOpen(true)
  }, [])

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.altKey || event.ctrlKey || event.metaKey || isTextEntryTarget(event.target)) {
        return
      }
      if (event.key !== '?') {
        return
      }
      event.preventDefault()
      setIsOpen(true)
    }

    window.addEventListener('keydown', handleWindowKeyDown, true)
    return () => {
      window.removeEventListener('keydown', handleWindowKeyDown, true)
    }
  }, [])

  const shortcutHelpButton = useMemo(
    () => (
      <Button
        aria-label="Show keyboard shortcuts"
        className="inline-flex h-8 w-8 items-center justify-center rounded-full border-0 bg-transparent p-0 text-[var(--color-mauve-700)] shadow-none transition-colors hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)]"
        data-testid="workbook-shortcut-button"
        type="button"
        onClick={openShortcutDialog}
      >
        <CircleHelp aria-hidden="true" className="size-4" strokeWidth={1.9} />
      </Button>
    ),
    [openShortcutDialog],
  )

  const shortcutDialog = useMemo(
    () => (
      <WorkbookShortcutDialog
        open={isOpen}
        query={query}
        onOpenChange={(open) => {
          if (open) {
            setIsOpen(true)
            return
          }
          closeShortcutDialog()
        }}
        onQueryChange={setQuery}
      />
    ),
    [closeShortcutDialog, isOpen, query],
  )

  return {
    closeShortcutDialog,
    openShortcutDialog,
    shortcutDialog,
    shortcutHelpButton,
  }
}
