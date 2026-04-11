import { cn } from "./cn.js";
import type { FormulaHelpEntry } from "./formulaAssist.js";
import { formulaHintClass } from "./formula-bar-theme.js";

interface FormulaArgumentHintProps {
  readonly entry: FormulaHelpEntry;
  readonly activeArgumentIndex: number;
}

export function FormulaArgumentHint({ entry, activeArgumentIndex }: FormulaArgumentHintProps) {
  return (
    <div aria-live="polite" className={formulaHintClass()} data-testid="formula-arg-hint">
      <span className="shrink-0 font-semibold text-[var(--color-mauve-950)]">{entry.name}</span>
      <span className="shrink-0">(</span>
      <span className="flex min-w-0 items-center gap-1.5">
        {entry.args.length === 0 ? (
          <span className="text-[var(--color-mauve-600)]">no arguments</span>
        ) : (
          entry.args.map((arg, index) => (
            <span
              className={cn(
                "shrink-0",
                index === activeArgumentIndex
                  ? "font-semibold text-[var(--color-mauve-950)]"
                  : undefined,
              )}
              key={`${entry.name}:${arg.label}`}
            >
              {index > 0 ? ", " : ""}
              {arg.optional ? `[${arg.label.replace(/^\[(.*)\]$/, "$1")}]` : arg.label}
            </span>
          ))
        )}
        {entry.variadic ? (
          <span className="shrink-0 text-[var(--color-mauve-600)]">, …</span>
        ) : null}
      </span>
      <span className="shrink-0">)</span>
    </div>
  );
}
