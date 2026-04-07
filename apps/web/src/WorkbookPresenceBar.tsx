import { cn } from "./cn.js";
import type { WorkbookCollaboratorPresence } from "./workbook-presence-model.js";
import { workbookSurfaceClass } from "./workbook-shell-chrome.js";

const PRESENCE_TONE_CLASS_NAMES = [
  {
    avatar: "bg-[var(--color-mauve-200)] text-[var(--color-mauve-900)]",
  },
  {
    avatar: "bg-[#e8efe8] text-[#31533d]",
  },
  {
    avatar: "bg-[#f3ede1] text-[#7a5b20]",
  },
  {
    avatar: "bg-[#efe7f6] text-[#654786]",
  },
  {
    avatar: "bg-[#f5e6e6] text-[#8a3b3b]",
  },
  {
    avatar: "bg-[#e3efef] text-[#2f6260]",
  },
  {
    avatar: "bg-[#e8ebf6] text-[#40518c]",
  },
  {
    avatar: "bg-[#f0e8f4] text-[#744488]",
  },
] as const;

export function WorkbookPresenceBar(props: {
  collaborators: readonly WorkbookCollaboratorPresence[];
  onJump: (sheetName: string, address: string) => void;
}) {
  if (props.collaborators.length === 0) {
    return null;
  }

  return (
    <div className="flex items-center gap-2 pl-2" data-testid="ax-rail">
      {props.collaborators.map((collaborator) => {
        const tone =
          PRESENCE_TONE_CLASS_NAMES[collaborator.toneIndex % PRESENCE_TONE_CLASS_NAMES.length]!;
        return (
          <button
            key={collaborator.sessionId}
            aria-label={`Jump to ${collaborator.label} at ${collaborator.sheetName}!${collaborator.address}`}
            className={cn(
              workbookSurfaceClass({ emphasis: "raised" }),
              "inline-flex h-8 items-center gap-2 px-2.5 text-[12px] font-medium text-[var(--wb-text-muted)] transition-colors hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1",
            )}
            data-testid="ax-presence-chip"
            title={`${collaborator.label} • ${collaborator.sheetName}!${collaborator.address}`}
            type="button"
            onClick={() => props.onJump(collaborator.sheetName, collaborator.address)}
          >
            <span
              aria-hidden="true"
              className={cn(
                "inline-flex h-5 w-5 items-center justify-center rounded-full text-[10px] font-semibold",
                tone.avatar,
              )}
            >
              {collaborator.initials}
            </span>
            <span className="max-w-28 truncate">{collaborator.label}</span>
            <span className="hidden text-[11px] text-[var(--wb-text-subtle)] xl:inline">
              {collaborator.sheetName}!{collaborator.address}
            </span>
          </button>
        );
      })}
    </div>
  );
}
