import { cn } from "./cn.js";
import type { FormulaHelpEntry } from "./formulaAssist.js";

interface FormulaArgumentHintProps {
  readonly entry: FormulaHelpEntry;
  readonly activeArgumentIndex: number;
}

export function FormulaArgumentHint({ entry, activeArgumentIndex }: FormulaArgumentHintProps) {
  return (
    <div
      aria-live="polite"
      className="flex min-h-7 items-center gap-2 overflow-x-auto rounded-[var(--wb-radius-control)] border border-[var(--wb-border-subtle)] bg-[var(--wb-surface-muted)] px-2.5 text-[11px] text-[var(--wb-text-subtle)]"
      data-testid="formula-arg-hint"
    >
      <span className="shrink-0 font-semibold text-[var(--wb-text)]">{entry.name}</span>
      <span className="shrink-0">(</span>
      <span className="flex min-w-0 items-center gap-1.5">
        {entry.args.length === 0 ? (
          <span className="text-[var(--wb-text-muted)]">no arguments</span>
        ) : (
          entry.args.map((arg, index) => (
            <span
              className={cn(
                "shrink-0",
                index === activeArgumentIndex ? "font-semibold text-[var(--wb-text)]" : undefined,
              )}
              key={`${entry.name}:${arg.label}`}
            >
              {index > 0 ? ", " : ""}
              {arg.optional ? `[${arg.label.replace(/^\[(.*)\]$/, "$1")}]` : arg.label}
            </span>
          ))
        )}
        {entry.variadic ? <span className="shrink-0 text-[var(--wb-text-muted)]">, …</span> : null}
      </span>
      <span className="shrink-0">)</span>
    </div>
  );
}
