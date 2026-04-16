import { Button } from '@base-ui/react/button'
import { cn } from './cn.js'
import type { WorkbookCollaboratorPresence } from './workbook-presence-model.js'
import { workbookHeaderSurfaceClass } from './workbook-header-controls.js'

const PRESENCE_TONE_CLASS_NAMES = [
  {
    avatar: 'bg-[var(--color-mauve-100)] text-[var(--color-mauve-900)]',
  },
  {
    avatar: 'bg-[var(--color-mauve-200)] text-[var(--color-mauve-900)]',
  },
  {
    avatar: 'bg-[var(--color-mauve-300)] text-[var(--color-mauve-900)]',
  },
  {
    avatar: 'bg-[var(--color-mauve-200)] text-[var(--color-mauve-900)]',
  },
  {
    avatar: 'bg-[var(--color-mauve-100)] text-[var(--color-mauve-900)]',
  },
  {
    avatar: 'bg-[var(--color-mauve-300)] text-[var(--color-mauve-900)]',
  },
  {
    avatar: 'bg-[var(--color-mauve-200)] text-[var(--color-mauve-900)]',
  },
  {
    avatar: 'bg-[var(--color-mauve-100)] text-[var(--color-mauve-900)]',
  },
] as const

export function WorkbookPresenceBar(props: {
  collaborators: readonly WorkbookCollaboratorPresence[]
  onJump: (sheetName: string, address: string) => void
}) {
  if (props.collaborators.length === 0) {
    return null
  }

  return (
    <div className="flex items-center gap-2" data-testid="ax-presence">
      {props.collaborators.map((collaborator) => {
        const tone = PRESENCE_TONE_CLASS_NAMES[collaborator.toneIndex % PRESENCE_TONE_CLASS_NAMES.length]!
        return (
          <Button
            key={collaborator.sessionId}
            aria-label={`Jump to ${collaborator.label} at ${collaborator.sheetName}!${collaborator.address}`}
            className={cn(
              workbookHeaderSurfaceClass,
              'gap-2 px-2.5 text-[12px] font-medium text-[var(--color-mauve-700)] transition-colors hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)] focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-[var(--color-mauve-50)]',
            )}
            data-testid="ax-presence-chip"
            title={`${collaborator.label} • ${collaborator.sheetName}!${collaborator.address}`}
            type="button"
            onClick={() => props.onJump(collaborator.sheetName, collaborator.address)}
          >
            <span
              aria-hidden="true"
              className={cn('inline-flex h-5 w-5 items-center justify-center rounded-full text-[10px] font-semibold', tone.avatar)}
            >
              {collaborator.initials}
            </span>
            <span className="max-w-28 truncate">{collaborator.label}</span>
            <span className="hidden text-[11px] text-[var(--color-mauve-500)] xl:inline">
              {collaborator.sheetName}!{collaborator.address}
            </span>
          </Button>
        )
      })}
    </div>
  )
}
