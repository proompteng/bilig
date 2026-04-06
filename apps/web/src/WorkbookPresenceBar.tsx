import { cn } from "./cn.js";
import type { WorkbookCollaboratorPresence } from "./workbook-presence-model.js";

const PRESENCE_TONE_CLASS_NAMES = [
  {
    avatar: "bg-[#dbeafe] text-[#1d4ed8]",
    chip: "border-[#bfdbfe] hover:border-[#60a5fa]",
  },
  {
    avatar: "bg-[#dcfce7] text-[#15803d]",
    chip: "border-[#bbf7d0] hover:border-[#4ade80]",
  },
  {
    avatar: "bg-[#fef3c7] text-[#b45309]",
    chip: "border-[#fde68a] hover:border-[#f59e0b]",
  },
  {
    avatar: "bg-[#fae8ff] text-[#a21caf]",
    chip: "border-[#f5d0fe] hover:border-[#d946ef]",
  },
  {
    avatar: "bg-[#fee2e2] text-[#b91c1c]",
    chip: "border-[#fecaca] hover:border-[#f87171]",
  },
  {
    avatar: "bg-[#cffafe] text-[#0f766e]",
    chip: "border-[#a5f3fc] hover:border-[#22d3ee]",
  },
  {
    avatar: "bg-[#e0e7ff] text-[#4338ca]",
    chip: "border-[#c7d2fe] hover:border-[#818cf8]",
  },
  {
    avatar: "bg-[#f3e8ff] text-[#7e22ce]",
    chip: "border-[#e9d5ff] hover:border-[#c084fc]",
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
              "inline-flex h-8 items-center gap-2 rounded-[var(--wb-radius-control)] border bg-[var(--wb-surface)] px-2.5 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1",
              tone.chip,
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
