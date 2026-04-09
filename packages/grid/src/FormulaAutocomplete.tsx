import { cn } from "./cn.js";
import type { FormulaSuggestion } from "./formulaAssist.js";

interface FormulaAutocompleteProps {
  readonly suggestions: readonly FormulaSuggestion[];
  readonly highlightedIndex: number;
  readonly onSelect: (index: number) => void;
}

export function FormulaAutocomplete({
  suggestions,
  highlightedIndex,
  onSelect,
}: FormulaAutocompleteProps) {
  if (suggestions.length === 0) {
    return null;
  }

  return (
    <div
      aria-label="Formula suggestions"
      className="absolute left-0 right-0 top-[calc(100%+0.375rem)] z-40 overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-app-bg)] shadow-[0_12px_28px_rgba(15,23,42,0.16)]"
      data-testid="formula-autocomplete"
      id="formula-autocomplete"
      role="listbox"
    >
      <ul className="max-h-72 overflow-auto py-1">
        {suggestions.map((suggestion, index) => {
          const active = index === highlightedIndex;
          return (
            <li
              aria-selected={active}
              className={cn(
                "cursor-pointer px-3 py-2",
                active ? "bg-[var(--wb-accent-soft)]" : "hover:bg-[var(--wb-muted)]",
              )}
              data-testid="formula-autocomplete-option"
              id={`formula-autocomplete-option-${index}`}
              key={`${suggestion.kind}:${suggestion.name}`}
              role="option"
              onMouseDown={(event) => {
                event.preventDefault();
                onSelect(index);
              }}
            >
              <div className="flex items-center justify-between gap-3">
                <div className="min-w-0">
                  <div className="truncate text-[12px] font-semibold text-[var(--wb-text)]">
                    {suggestion.name}
                  </div>
                  <div className="truncate text-[11px] text-[var(--wb-text-subtle)]">
                    {suggestion.kind === "function" ? suggestion.signature : suggestion.summary}
                  </div>
                </div>
                <span className="shrink-0 text-[10px] font-medium uppercase tracking-[0.08em] text-[var(--wb-text-muted)]">
                  {suggestion.kind === "function" ? suggestion.category : "Name"}
                </span>
              </div>
              {suggestion.kind === "function" ? (
                <div className="mt-1 truncate text-[11px] text-[var(--wb-text-subtle)]">
                  {suggestion.summary}
                </div>
              ) : null}
            </li>
          );
        })}
      </ul>
    </div>
  );
}
